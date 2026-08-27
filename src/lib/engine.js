import { normalizeText, normalizeTarget, toNumber, toRate, mapReportColumns, reportDateRange, bulkDateRangeFromFileName } from "./parsers.js";

const clamp = (n, min, max) => Math.min(max, Math.max(min, n));
const ratio = (a, b) => b > 0 ? a / b : null;
const key = (...parts) => parts.map(v => normalizeText(v)).join("||");
const fmtMoney = n => `$${Number(n || 0).toFixed(2)}`;

function rowValue(row, names) {
  const entries = Object.entries(row || {});
  for (const name of names) {
    const found = entries.find(([k]) => normalizeText(k) === normalizeText(name));
    if (found) {
      const value = found[1];
      if (value !== undefined && value !== null && String(value).trim() !== "") {
        return value;
      }
    }
  }
  return "";
}

function targetType(targeting, matchType = "", entity = "") {
  const t = normalizeTarget(targeting);
  const m = normalizeText(matchType);
  const e = normalizeText(entity);
  if (["close-match", "loose-match", "substitutes", "complements"].includes(t)) return "AUTO";
  if (/^asin=|^asin\"|asin=\"|category=|brand=/.test(t) || e.includes("product targeting")) return "PRODUCT";
  if (["exact", "phrase", "broad"].includes(m) || e.includes("keyword")) return "KEYWORD";
  return "OTHER";
}

function canonicalTarget(row) {
  const keyword = String(rowValue(row, ["Keyword text"])).trim();
  const expr = String(rowValue(row, ["Product targeting expression"])).trim();
  return keyword || expr;
}

export function buildBulkModel(rows) {
  const campaigns = new Map();
  const adGroups = new Map();
  const targets = [];
  const negatives = [];
  const placements = [];
  const productAds = [];

  for (const row of rows) {
    const entity = normalizeText(rowValue(row, ["Entity"]));
    const campaignId = String(rowValue(row, ["Campaign ID"])).trim();
    const adGroupId = String(rowValue(row, ["Ad Group ID"])).trim();
    const keywordId = String(rowValue(row, ["Keyword ID"])).trim();
    const productTargetId = String(rowValue(row, ["Product Targeting ID"])).trim();
    const campaignName = String(rowValue(row, ["Campaign name", "Campaign name (Informational only)"])).trim();
    const adGroupName = String(rowValue(row, ["Ad group name", "Ad group name (Informational only)"])).trim();
    const state = normalizeText(rowValue(row, ["State"]));
    const campaignState = normalizeText(rowValue(row, ["Campaign state (Informational only)"]));
    const adGroupState = normalizeText(rowValue(row, ["Ad group state (Informational only)"]));
    const matchType = normalizeText(rowValue(row, ["Match type"]));
    const targeting = canonicalTarget(row);
    const bid = toNumber(rowValue(row, ["Bid"]));
    const defaultBid = toNumber(rowValue(row, ["Ad Group Default Bid", "Ad Group Default Bid (Informational only)"]));
    const placement = String(rowValue(row, ["Placement"])).trim();
    const percentage = toNumber(rowValue(row, ["Percentage"]));
    const spend = toNumber(rowValue(row, ["Spend"]));
    const sales = toNumber(rowValue(row, ["Sales"]));
    const orders = toNumber(rowValue(row, ["Orders"]));
    const clicks = toNumber(rowValue(row, ["Clicks"]));
    const impressions = toNumber(rowValue(row, ["Impressions"]));
    const acos = sales > 0 ? spend / sales : null;
    const cpc = clicks > 0 ? spend / clicks : toNumber(rowValue(row, ["CPC"]));

    if (entity === "campaign" && campaignId) {
      campaigns.set(campaignId, { campaignId, campaignName, state, campaignState, dailyBudget: toNumber(rowValue(row, ["Daily budget"])), row });
    }
    if (entity === "ad group" && adGroupId) {
      adGroups.set(adGroupId, { campaignId, campaignName, adGroupId, adGroupName, state, adGroupState, defaultBid, row });
    }
    if (entity === "bidding adjustment") {
      placements.push({ campaignId, campaignName, placement, percentage, spend, sales, orders, clicks, impressions, acos, cpc, row });
    }
    if (entity === "product ad") {
      productAds.push({ campaignId, campaignName, adGroupId, adGroupName, sku: rowValue(row, ["SKU"]), asin: rowValue(row, ["ASIN (Informational only)"]), state, row });
    }

    const isNegative = entity.includes("negative");
    const isTarget = entity.includes("keyword") || entity.includes("product targeting");
    if (isTarget && targeting) {
      const rec = {
        entity, campaignId, adGroupId, keywordId, productTargetId,
        campaignName, adGroupName, state, campaignState, adGroupState,
        targeting, normalizedTarget: normalizeTarget(targeting), matchType, bid,
        targetType: targetType(targeting, matchType, entity), row,
      };
      if (isNegative) negatives.push(rec); else targets.push(rec);
    }
  }

  const byEntityId = new Map();
  const byStructure = new Map();
  const byFallbackStructure = new Map();

  for (const t of targets) {
    const id = t.keywordId || t.productTargetId;
    if (id) byEntityId.set(String(id), t);

    const exactKey = key(t.campaignName, t.adGroupName, t.normalizedTarget, t.matchType);
    if (!byStructure.has(exactKey)) byStructure.set(exactKey, []);
    byStructure.get(exactKey).push(t);

    const fallbackKey = key(t.campaignName, t.normalizedTarget);
    if (!byFallbackStructure.has(fallbackKey)) byFallbackStructure.set(fallbackKey, []);
    byFallbackStructure.get(fallbackKey).push(t);
  }

  return {
    campaigns,
    adGroups,
    targets,
    negatives,
    placements,
    productAds,
    byEntityId,
    byStructure,
    byFallbackStructure,
  };
}

function findBulkTarget(metric, bulk) {
  const entityId = metric.keywordId || metric.productTargetId;

  if (entityId && bulk.byEntityId.has(String(entityId))) {
    return {
      target: bulk.byEntityId.get(String(entityId)),
      quality: "Exact ID",
    };
  }

  const exactKey = key(
    metric.campaignName,
    metric.adGroupName,
    normalizeTarget(metric.targeting),
    metric.matchType
  );

  const exactMatches = bulk.byStructure.get(exactKey) || [];
  if (exactMatches.length === 1) {
    return {
      target: exactMatches[0],
      quality: "Name + Target",
    };
  }

  // First fallback: ignore match type, but keep campaign + ad group + target.
  const adGroupFallback = bulk.targets.filter(t =>
    normalizeText(t.campaignName) === normalizeText(metric.campaignName) &&
    normalizeText(t.adGroupName) === normalizeText(metric.adGroupName) &&
    normalizeTarget(t.targeting) === normalizeTarget(metric.targeting)
  );

  if (adGroupFallback.length === 1) {
    return {
      target: adGroupFallback[0],
      quality: "Name + Target (fallback)",
    };
  }

  // Second fallback: campaign + target only. This handles Amazon rows where
  // the ad-group metadata is unavailable or represented differently.
  const fallbackKey = key(metric.campaignName, normalizeTarget(metric.targeting));
  const campaignFallback = bulk.byFallbackStructure.get(fallbackKey) || [];

  if (campaignFallback.length === 1) {
    return {
      target: campaignFallback[0],
      quality: "Name + Target (fallback)",
    };
  }

  return {
    target: null,
    quality: "Unmatched",
  };
}

function isActive(t) {
  const states = [t.state, t.campaignState, t.adGroupState].filter(Boolean);
  return !states.some(s => ["paused", "archived"].includes(normalizeText(s)));
}

function existingExact(term, bulk) {
  const n = normalizeTarget(term);
  return bulk.targets.some(t => isActive(t) && t.targetType === "KEYWORD" && t.matchType === "exact" && t.normalizedTarget === n);
}

function existingAny(term, bulk) {
  const n = normalizeTarget(term);
  return bulk.targets.some(t => isActive(t) && t.normalizedTarget === n);
}

function existingNegative(term, campaignName, adGroupName, bulk) {
  const n = normalizeTarget(term);
  return bulk.negatives.some(t => {
    if (t.normalizedTarget !== n) return false;
    const sameCampaign = !t.campaignName || normalizeText(t.campaignName) === normalizeText(campaignName);
    const sameAdGroup = !t.adGroupName || normalizeText(t.adGroupName) === normalizeText(adGroupName);
    return sameCampaign && sameAdGroup;
  });
}

function aggregateSearchTerms(rows, cols, bulk) {
  const map = new Map();
  for (const r of rows) {
    const term = String(r[cols.term] ?? "").trim();
    if (!term) continue;
    const campaignName = cols.campaignName ? String(r[cols.campaignName] ?? "").trim() : "";
    const adGroupName = cols.adGroupName ? String(r[cols.adGroupName] ?? "").trim() : "";
    const campaignId = cols.campaignId ? String(r[cols.campaignId] ?? "").trim() : "";
    const adGroupId = cols.adGroupId ? String(r[cols.adGroupId] ?? "").trim() : "";
    const keywordId = cols.keywordId ? String(r[cols.keywordId] ?? "").trim() : "";
    const productTargetId = cols.productTargetId ? String(r[cols.productTargetId] ?? "").trim() : "";
    const targeting = cols.targeting ? String(r[cols.targeting] ?? "").trim() : "";
    const matchType = cols.matchType ? normalizeText(r[cols.matchType]) : "";
    const k = key(campaignId || campaignName, adGroupId || adGroupName, targeting, matchType, term);
    if (!map.has(k)) map.set(k, {
      term, campaignId, adGroupId, keywordId, productTargetId, campaignName, adGroupName, targeting, matchType,
      spend: 0, sales: 0, orders: 0, clicks: 0, impressions: 0,
    });
    const x = map.get(k);
    x.spend += toNumber(r[cols.spend]);
    x.sales += toNumber(r[cols.sales]);
    x.orders += toNumber(r[cols.orders]);
    x.clicks += toNumber(r[cols.clicks]);
    if (cols.impressions) x.impressions += toNumber(r[cols.impressions]);
  }

  return [...map.values()].map(x => ({
    ...x,
    acos: ratio(x.spend, x.sales),
    cvr: ratio(x.orders, x.clicks),
    cpc: ratio(x.spend, x.clicks),
    queryType: /^b0[a-z0-9]{8}$/i.test(normalizeText(x.term)) ? "ASIN" : "KEYWORD",
    existingExact: existingExact(x.term, bulk),
    existingAny: existingAny(x.term, bulk),
    existingNegative: existingNegative(x.term, x.campaignName, x.adGroupName, bulk),
  }));
}

function aggregateTargets(rows, cols, bulk) {
  const map = new Map();
  for (const r of rows) {
    const targeting = String(r[cols.targeting] ?? "").trim();
    if (!targeting) continue;
    const campaignName = cols.campaignName ? String(r[cols.campaignName] ?? "").trim() : "";
    const adGroupName = cols.adGroupName ? String(r[cols.adGroupName] ?? "").trim() : "";
    const campaignId = cols.campaignId ? String(r[cols.campaignId] ?? "").trim() : "";
    const adGroupId = cols.adGroupId ? String(r[cols.adGroupId] ?? "").trim() : "";
    const matchType = cols.matchType ? normalizeText(r[cols.matchType]) : "";
    const k = key(campaignId || campaignName, adGroupId || adGroupName, targeting, matchType);
    if (!map.has(k)) map.set(k, { campaignId, adGroupId, campaignName, adGroupName, targeting, matchType, spend:0, sales:0, orders:0, clicks:0, impressions:0, tosWeighted:0, tosWeight:0 });
    const x = map.get(k);
    const clicks = toNumber(r[cols.clicks]);
    x.spend += toNumber(r[cols.spend]);
    x.sales += toNumber(r[cols.sales]);
    x.orders += toNumber(r[cols.orders]);
    x.clicks += clicks;
    if (cols.impressions) x.impressions += toNumber(r[cols.impressions]);
    if (cols.tosShare) {
      const share = toRate(r[cols.tosShare]);
      if (share != null) { x.tosWeighted += share * Math.max(clicks, 1); x.tosWeight += Math.max(clicks, 1); }
    }
  }

  return [...map.values()].map(x => {
    const match = findBulkTarget(x, bulk);
    const target = match.target;
    return {
      ...x,
      acos: ratio(x.spend, x.sales), cvr: ratio(x.orders, x.clicks), cpc: ratio(x.spend, x.clicks),
      tosShare: x.tosWeight ? x.tosWeighted / x.tosWeight : null,
      currentBid: target?.bid || 0,
      state: target?.state || "",
      targetType: target?.targetType || targetType(x.targeting, x.matchType),
      bulkTarget: target,
      matchQuality: match.quality,
    };
  });
}

function relevance(term, brandTerms = []) {
  const n = normalizeText(term);
  if (brandTerms.some(b => b && n.includes(normalizeText(b)))) return "BRAND";
  if (/^b0[a-z0-9]{8}$/i.test(n)) return "ASIN";
  return "UNKNOWN";
}

function confidenceForZeroOrder(x, targetCpa, minClicks) {
  const cpaMultiple = targetCpa > 0 ? x.spend / targetCpa : 0;
  if (x.clicks >= Math.max(20, minClicks * 2) && cpaMultiple >= 2) return "High";
  if (x.clicks >= minClicks && cpaMultiple >= 1.25) return "Medium";
  if (x.clicks >= Math.max(12, minClicks) || cpaMultiple >= 1.5) return "Medium";
  return "Low";
}

export function validateInputs({ searchWorkbook, targetWorkbook, bulkWorkbook, searchSheet, targetSheet, bulkSheet, bulkSearchSheet }) {
  const stCols = mapReportColumns(searchSheet.rows, "searchTerm");
  const tgCols = mapReportColumns(targetSheet.rows, "targeting");
  const bulk = buildBulkModel(bulkSheet.rows);
  const stRange = reportDateRange(searchSheet.rows, stCols);
  const tgRange = reportDateRange(targetSheet.rows, tgCols);
  const bulkRange = bulkDateRangeFromFileName(bulkWorkbook.fileName);
  const targetAgg = aggregateTargets(targetSheet.rows, tgCols, bulk);
  const matched = targetAgg.filter(x => x.matchQuality !== "Unmatched").length;
  const matchRate = targetAgg.length ? matched / targetAgg.length : 0;

  const matchQualityCounts = targetAgg.reduce((acc, x) => {
    acc[x.matchQuality] = (acc[x.matchQuality] || 0) + 1;
    return acc;
  }, {});
  const reportCampaigns = new Set(targetAgg.map(x => normalizeText(x.campaignName)).filter(Boolean));
  const bulkCampaignsByName = new Set([...bulk.campaigns.values()].map(x => normalizeText(x.campaignName)).filter(Boolean));
  const missingCampaigns = [...reportCampaigns].filter(x => !bulkCampaignsByName.has(x));
  const warnings = [];

  if (bulkRange && stRange && bulkRange.end < stRange.end) {
    const days = Math.round((stRange.end - bulkRange.end) / 86400000);
    warnings.push(`Bulk Sheet ends ${days} day${days === 1 ? "" : "s"} before the Search Term Report. Newer campaign/target changes may be unmatched.`);
  }
  if (matchRate < 0.95) warnings.push(`${((1 - matchRate) * 100).toFixed(1)}% of Targeting Report entities could not be matched to the uploaded Bulk Sheet.`);
  if (missingCampaigns.length) warnings.push(`${missingCampaigns.length} campaign${missingCampaigns.length === 1 ? "" : "s"} appear in performance reports but not in the Bulk Sheet.`);

  const entityCounts = {};
  for (const r of bulkSheet.rows) {
    const e = String(rowValue(r, ["Entity"]) || "Unknown");
    entityCounts[e] = (entityCounts[e] || 0) + 1;
  }

  return {
    search: { file: searchWorkbook.fileName, sheet: searchSheet.name, rows: searchSheet.rows.length, range: stRange },
    targeting: { file: targetWorkbook.fileName, sheet: targetSheet.name, rows: targetSheet.rows.length, range: tgRange },
    bulk: { file: bulkWorkbook.fileName, sheet: bulkSheet.name, rows: bulkSheet.rows.length, range: bulkRange, embeddedSearchRows: bulkSearchSheet?.rows?.length || 0 },
    counts: { campaigns: bulk.campaigns.size, adGroups: bulk.adGroups.size, positiveTargets: bulk.targets.length, negatives: bulk.negatives.length, placements: bulk.placements.length, productAds: bulk.productAds.length },
    entityCounts,
    matchRate,
    matchQualityCounts,
    matchedTargets: matched,
    totalTargetRows: targetAgg.length,
    missingCampaigns,
    warnings,
  };
}

export function analyseAccount({ searchRows, targetingRows, bulkRows, targetAcosPct = 30, minClicks = 6, brandTerms = [] }) {
  const targetAcos = clamp(Number(targetAcosPct) / 100, 0.05, 0.80);
  const stCols = mapReportColumns(searchRows, "searchTerm");
  const tgCols = mapReportColumns(targetingRows, "targeting");
  const bulk = buildBulkModel(bulkRows);
  const searchTerms = aggregateSearchTerms(searchRows, stCols, bulk);
  const targets = aggregateTargets(targetingRows, tgCols, bulk);

  const totalSpend = searchTerms.reduce((s, x) => s + x.spend, 0);
  const totalSales = searchTerms.reduce((s, x) => s + x.sales, 0);
  const totalOrders = searchTerms.reduce((s, x) => s + x.orders, 0);
  const totalClicks = searchTerms.reduce((s, x) => s + x.clicks, 0);
  const accountAov = totalOrders > 0 ? totalSales / totalOrders : 0;
  const targetCpa = accountAov * targetAcos;

  const negativeCandidates = [];
  const watch = [];
  const harvest = [];

  for (const t of searchTerms) {
    const rel = relevance(t.term, brandTerms);
    const confidence = confidenceForZeroOrder(t, targetCpa, minClicks);
    const cpaMultiple = targetCpa > 0 ? t.spend / targetCpa : null;

    if (t.orders === 0 && t.spend > 0) {
      if (t.existingNegative) {
        watch.push({ ...t, relevance: rel, confidence: "Low", cpaMultiple, action: "Already negative", reason: "This query is already represented by an existing negative in the Bulk Sheet." });
      } else if (rel === "BRAND") {
        watch.push({ ...t, relevance: rel, confidence: confidence === "High" ? "Medium" : confidence, cpaMultiple, action: "Lower bid / review", reason: "Brand-relevant zero-order query; do not auto-negate without review." });
      } else if (["High", "Medium"].includes(confidence)) {
        negativeCandidates.push({ ...t, relevance: rel, confidence, cpaMultiple, action: t.queryType === "ASIN" ? "Negative product target / review" : "Add Negative Exact / review source target", reason: `${t.clicks} clicks, ${fmtMoney(t.spend)} spend, 0 orders${targetCpa ? `; ${cpaMultiple.toFixed(1)}x target CPA` : ""}.` });
      } else {
        watch.push({ ...t, relevance: rel, confidence, cpaMultiple, action: "Watch", reason: "Zero orders but not enough spend/click evidence for an automatic negative recommendation." });
      }
    }

    if (t.queryType === "KEYWORD" && t.orders >= 2 && t.acos != null && t.acos <= targetAcos * 1.10 && !t.existingExact) {
      harvest.push({
        ...t,
        relevance: rel,
        confidence: t.orders >= 5 ? "High" : "Medium",
        action: t.existingAny ? "Add Exact (already targeted elsewhere)" : "Create Exact keyword",
        reason: `${t.orders} orders at ${(t.acos * 100).toFixed(1)}% ACoS; no active Exact target found in Bulk Sheet.`,
      });
    }
  }

  const pauseTargets = [];
  const bidDown = [];
  const bidUp = [];
  const targetWatch = [];
  for (const t of targets) {
    const confidence = confidenceForZeroOrder(t, targetCpa, minClicks);
    const cpaMultiple = targetCpa > 0 ? t.spend / targetCpa : null;
    const active = !t.bulkTarget || isActive(t.bulkTarget);

    if (t.orders === 0 && t.spend > 0 && active) {
      if (["High", "Medium"].includes(confidence)) {
        pauseTargets.push({ ...t, confidence, cpaMultiple, action: t.matchQuality === "Unmatched" ? "Review (Bulk target unmatched)" : "Pause / lower bid", reason: `${t.clicks} clicks, ${fmtMoney(t.spend)} spend, 0 orders${targetCpa ? `; ${cpaMultiple.toFixed(1)}x target CPA` : ""}.` });
      } else {
        targetWatch.push({ ...t, confidence: "Low", cpaMultiple, action: "Watch", reason: "Target has zero orders but insufficient evidence for a pause." });
      }
      continue;
    }

    if (t.orders >= 2 && t.acos != null && t.currentBid > 0) {
      if (t.acos > targetAcos * 1.05) {
        const cut = clamp(1 - (targetAcos / t.acos), 0.05, 0.40);
        const suggestedBid = t.currentBid * (1 - cut);
        bidDown.push({
          ...t, confidence: t.orders >= 5 ? "High" : "Medium",
          suggestedBid, changePct: -cut,
          action: `Lower bid ${(cut * 100).toFixed(0)}%`,
          reason: `${(t.acos * 100).toFixed(1)}% ACoS vs ${targetAcosPct}% target; cut capped at 40%.`,
        });
      } else if (t.acos < targetAcos * 0.75) {
        let increase = clamp((targetAcos / Math.max(t.acos, 0.01)) - 1, 0.05, 0.30);
        if (t.tosShare != null && t.tosShare >= 0.70) increase = Math.min(increase, 0.10);
        const suggestedBid = t.currentBid * (1 + increase);
        bidUp.push({
          ...t, confidence: t.orders >= 5 ? "High" : "Medium",
          suggestedBid, changePct: increase,
          action: `Raise bid ${(increase * 100).toFixed(0)}%`,
          reason: `${(t.acos * 100).toFixed(1)}% ACoS vs ${targetAcosPct}% target${t.tosShare != null ? `; TOS share ${(t.tosShare * 100).toFixed(1)}%` : ""}.`,
        });
      }
    }
  }

  const overlapMap = new Map();
  for (const t of bulk.targets) {
    if (!isActive(t)) continue;
    const k = key(t.normalizedTarget, t.matchType);
    if (!overlapMap.has(k)) overlapMap.set(k, []);
    overlapMap.get(k).push(t);
  }
  const overlaps = [...overlapMap.values()]
    .filter(arr => arr.length > 1)
    .map(arr => ({
      targeting: arr[0].targeting,
      matchType: arr[0].matchType || "-",
      targetType: arr[0].targetType,
      count: arr.length,
      locations: arr,
      action: "Review query routing",
      reason: "Same active target/match type exists in multiple campaign/ad-group locations; overlap is not automatically waste.",
    }))
    .sort((a, b) => b.count - a.count);

  const placements = bulk.placements
    .filter(p => p.spend > 0 || p.orders > 0)
    .map(p => {
      let action = "Monitor";
      if (p.orders >= 3 && p.acos != null && p.acos > targetAcos * 1.20) action = "Review / reduce placement adjustment";
      if (p.orders >= 3 && p.acos != null && p.acos < targetAcos * 0.80) action = "Review / increase placement adjustment";
      return { ...p, action };
    })
    .sort((a, b) => b.spend - a.spend);

  const wastedSpend = negativeCandidates.reduce((s, x) => s + x.spend, 0);
  const bidDownSpend = bidDown.reduce((s, x) => s + x.spend, 0);
  const highConfidenceWaste = negativeCandidates.filter(x => x.confidence === "High").reduce((s, x) => s + x.spend, 0);

  const actionPlan = [
    negativeCandidates.length ? { priority: 1, title: "Eliminate high-confidence wasted queries", impact: highConfidenceWaste || wastedSpend, detail: `${negativeCandidates.length} negative candidates; ${fmtMoney(wastedSpend)} total flagged spend.` } : null,
    bidDown.length ? { priority: 2, title: "Reduce bids on inefficient converting targets", impact: bidDownSpend, detail: `${bidDown.length} targets above Target ACoS with ${fmtMoney(bidDownSpend)} analyzed spend.` } : null,
    harvest.length ? { priority: 3, title: "Harvest proven customer search terms", impact: harvest.reduce((s,x)=>s+x.sales,0), detail: `${harvest.length} converting queries have no active Exact target.` } : null,
    overlaps.length ? { priority: 4, title: "Review target overlap and query routing", impact: overlaps.length, detail: `${overlaps.length} active target/match combinations exist in multiple locations.` } : null,
    bidUp.length ? { priority: 5, title: "Scale efficient targets selectively", impact: bidUp.reduce((s,x)=>s+x.sales,0), detail: `${bidUp.length} targets are materially below Target ACoS and have conversion evidence.` } : null,
  ].filter(Boolean);

  return {
    settings: { targetAcos, targetAcosPct, targetCpa, accountAov, minClicks, brandTerms },
    summary: { totalSpend, totalSales, totalOrders, totalClicks, overallAcos: ratio(totalSpend, totalSales), accountAov, targetCpa, wastedSpend, highConfidenceWaste, bidDownSpend },
    counts: { campaigns: bulk.campaigns.size, adGroups: bulk.adGroups.size, positiveTargets: bulk.targets.length, negatives: bulk.negatives.length },
    negativeCandidates: negativeCandidates.sort((a,b)=>b.spend-a.spend),
    bidDown: bidDown.sort((a,b)=>b.spend-a.spend),
    bidUp: bidUp.sort((a,b)=>b.sales-a.sales),
    harvest: harvest.sort((a,b)=>b.orders-a.orders),
    pauseTargets: pauseTargets.sort((a,b)=>b.spend-a.spend),
    watch: [...watch, ...targetWatch].sort((a,b)=>(b.spend||0)-(a.spend||0)),
    overlaps,
    placements,
    actionPlan,
  };
}
