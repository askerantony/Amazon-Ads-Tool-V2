import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { parseWorkbook, pickBestSheet } from "./lib/parsers.js";
import { analyseAccount, validateInputs } from "./lib/engine.js";

const TABS = [
  ["summary", "Executive Summary"],
  ["negative", "Waste / Negate"],
  ["bidDown", "Lower Bid"],
  ["bidUp", "Increase Bid"],
  ["harvest", "Harvest"],
  ["pause", "Pause Targets"],
  ["watch", "Watch"],
  ["overlap", "Structure / Overlap"],
  ["placement", "Placement"],
  ["plan", "Action Plan"],
];

const money = n => `$${Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const pct = n => n == null || !Number.isFinite(n) ? "—" : `${(n * 100).toFixed(1)}%`;
const fmtDate = d => d ? new Date(d).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }) : "—";
const changePct = n => n == null ? "—" : `${n >= 0 ? "+" : ""}${(n * 100).toFixed(0)}%`;

function UploadCard({ title, hint, file, onChange }) {
  const ref = useRef(null);
  const [drag, setDrag] = useState(false);
  const choose = f => f && onChange(f);
  return (
    <div className={`uploadCard ${file ? "ready" : ""} ${drag ? "drag" : ""}`}
      onClick={() => ref.current?.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); choose(e.dataTransfer.files?.[0]); }}>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" onChange={e => choose(e.target.files?.[0])} />
      <div className="uploadIcon">{file ? "✓" : "↥"}</div>
      <strong>{file ? file.name : title}</strong>
      <span>{file ? "Ready for validation" : hint}</span>
    </div>
  );
}

function Badge({ children, tone = "neutral" }) { return <span className={`badge ${tone}`}>{children}</span>; }
function Confidence({ value }) { return <Badge tone={String(value || "").toLowerCase()}>{value || "—"}</Badge>; }

function Kpi({ label, value, tone = "" }) {
  return <div className={`kpi ${tone}`}><span>{label}</span><strong>{value}</strong></div>;
}

function ValidationCard({ title, info, tone = "good" }) {
  return (
    <div className="validationCard">
      <div className="validationTop"><Badge tone={tone}>{tone === "good" ? "Validated" : "Review"}</Badge><strong>{title}</strong></div>
      <div className="validationGrid">
        <span>File</span><b title={info.file}>{info.file}</b>
        <span>Sheet</span><b>{info.sheet}</b>
        <span>Rows</span><b>{Number(info.rows || 0).toLocaleString()}</b>
        <span>Date range</span><b>{info.range ? `${fmtDate(info.range.start)} → ${fmtDate(info.range.end)}` : "Not available"}</b>
      </div>
    </div>
  );
}

function Table({ columns, rows, limit = 100 }) {
  const [shown, setShown] = useState(limit);
  if (!rows?.length) return <div className="empty">No items found for this section.</div>;
  return <>
    <div className="tableWrap"><table><thead><tr>{columns.map(c => <th key={c.key}>{c.label}</th>)}</tr></thead>
      <tbody>{rows.slice(0, shown).map((r, i) => <tr key={i}>{columns.map(c => <td key={c.key}>{c.render ? c.render(r) : String(r[c.key] ?? "")}</td>)}</tr>)}</tbody>
    </table></div>
    {shown < rows.length && <button className="ghostBtn more" onClick={() => setShown(x => x + 100)}>Show 100 more ({rows.length - shown} remaining)</button>}
  </>;
}

const commonSearchCols = [
  { key: "campaignName", label: "Campaign" },
  { key: "adGroupName", label: "Ad Group" },
  { key: "targeting", label: "Source Target" },
  { key: "term", label: "Customer Search Term" },
  { key: "matchType", label: "Match" },
  { key: "clicks", label: "Clicks" },
  { key: "spend", label: "Spend", render: r => money(r.spend) },
  { key: "sales", label: "Sales", render: r => money(r.sales) },
  { key: "orders", label: "Orders" },
  { key: "acos", label: "ACoS", render: r => pct(r.acos) },
  { key: "confidence", label: "Confidence", render: r => <Confidence value={r.confidence} /> },
  { key: "action", label: "Recommended Action", render: r => <Badge tone="action">{r.action}</Badge> },
  { key: "reason", label: "Reason" },
];

const targetCols = [
  { key: "campaignName", label: "Campaign" },
  { key: "adGroupName", label: "Ad Group" },
  { key: "targeting", label: "Targeting" },
  { key: "targetType", label: "Type", render: r => <Badge>{r.targetType}</Badge> },
  { key: "matchType", label: "Match" },
  { key: "clicks", label: "Clicks" },
  { key: "spend", label: "Spend", render: r => money(r.spend) },
  { key: "sales", label: "Sales", render: r => money(r.sales) },
  { key: "orders", label: "Orders" },
  { key: "acos", label: "ACoS", render: r => pct(r.acos) },
  { key: "currentBid", label: "Current Bid", render: r => r.currentBid ? money(r.currentBid) : "—" },
  { key: "suggestedBid", label: "Suggested Bid", render: r => r.suggestedBid ? money(r.suggestedBid) : "—" },
  { key: "changePct", label: "Change", render: r => changePct(r.changePct) },
  { key: "matchQuality", label: "Bulk Match", render: r => <Badge tone={r.matchQuality === "Unmatched" ? "medium" : "good"}>{r.matchQuality}</Badge> },
  { key: "confidence", label: "Confidence", render: r => <Confidence value={r.confidence} /> },
  { key: "action", label: "Action", render: r => <Badge tone="action">{r.action}</Badge> },
  { key: "reason", label: "Reason" },
];

function exportRecommendations(results, validation) {
  const wb = XLSX.utils.book_new();
  const clean = rows => rows.map(r => {
    const out = {};
    for (const [k, v] of Object.entries(r)) {
      if (["row", "bulkTarget", "locations"].includes(k)) continue;
      if (typeof v === "object" && v !== null) continue;
      out[k] = v;
    }
    return out;
  });
  const summary = [
    { Metric: "Total Spend", Value: results.summary.totalSpend },
    { Metric: "Total Sales", Value: results.summary.totalSales },
    { Metric: "Orders", Value: results.summary.totalOrders },
    { Metric: "Overall ACoS", Value: results.summary.overallAcos },
    { Metric: "Target ACoS", Value: results.settings.targetAcos },
    { Metric: "Target CPA", Value: results.summary.targetCpa },
    { Metric: "Flagged Waste", Value: results.summary.wastedSpend },
    { Metric: "Targeting-to-Bulk Match Rate", Value: validation.matchRate },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Executive Summary");
  const tabs = [
    ["Waste Negate", results.negativeCandidates], ["Lower Bid", results.bidDown], ["Increase Bid", results.bidUp],
    ["Harvest", results.harvest], ["Pause Targets", results.pauseTargets], ["Watch", results.watch],
    ["Placement", results.placements],
  ];
  for (const [name, rows] of tabs) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(clean(rows)), name.slice(0, 31));
  const overlapRows = results.overlaps.map(o => ({
    Targeting: o.targeting, MatchType: o.matchType, TargetType: o.targetType, Count: o.count,
    Locations: o.locations.map(x => `${x.campaignName} > ${x.adGroupName}`).join(" | "), Action: o.action, Reason: o.reason,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(overlapRows), "Structure Overlap");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.actionPlan.map(x => ({ Priority:x.priority, Action:x.title, Impact:x.impact, Detail:x.detail }))), "Action Plan");
  XLSX.writeFile(wb, `Amazon-Ads-Audit-V2-${new Date().toISOString().slice(0,10)}.xlsx`);
}

function Summary({ results, validation }) {
  const s = results.summary;
  return <>
    <div className="kpiGrid">
      <Kpi label="Total Spend" value={money(s.totalSpend)} />
      <Kpi label="Attributed Sales" value={money(s.totalSales)} tone="green" />
      <Kpi label="Overall ACoS" value={pct(s.overallAcos)} tone={s.overallAcos > results.settings.targetAcos ? "red" : "green"} />
      <Kpi label="Orders" value={Number(s.totalOrders).toLocaleString()} />
      <Kpi label="Target CPA" value={money(s.targetCpa)} />
      <Kpi label="Flagged Waste" value={money(s.wastedSpend)} tone="red" />
      <Kpi label="Bid-Down Spend" value={money(s.bidDownSpend)} tone="orange" />
      <Kpi label="Bulk Match Rate" value={pct(validation.matchRate)} tone={validation.matchRate < .9 ? "red" : "green"} />
    </div>
    <div className="twoCol">
      <div className="panel inset"><h3>Recommendation inventory</h3><div className="statList">
        <div><span>Waste / negative candidates</span><b>{results.negativeCandidates.length}</b></div>
        <div><span>Bid reductions</span><b>{results.bidDown.length}</b></div>
        <div><span>Bid increases</span><b>{results.bidUp.length}</b></div>
        <div><span>Harvest opportunities</span><b>{results.harvest.length}</b></div>
        <div><span>Pause / review targets</span><b>{results.pauseTargets.length}</b></div>
        <div><span>Watch items</span><b>{results.watch.length}</b></div>
      </div></div>
      <div className="panel inset"><h3>Account structure</h3><div className="statList">
        <div><span>Campaigns</span><b>{validation.counts.campaigns}</b></div>
        <div><span>Ad groups</span><b>{validation.counts.adGroups}</b></div>
        <div><span>Positive targets</span><b>{validation.counts.positiveTargets}</b></div>
        <div><span>Negative targets</span><b>{validation.counts.negatives}</b></div>
        <div><span>Placement records</span><b>{validation.counts.placements}</b></div>
        <div><span>Product ads</span><b>{validation.counts.productAds}</b></div>
      </div></div>
    </div>
    {validation.warnings.length > 0 && <div className="warningBox"><strong>Data-quality warnings</strong>{validation.warnings.map((w,i)=><div key={i}>• {w}</div>)}</div>}
  </>;
}

function ActionPlan({ rows }) {
  if (!rows.length) return <div className="empty">No prioritized actions generated.</div>;
  return <div className="roadmap">{rows.map(r => <div className="roadmapItem" key={r.priority}><div className="priority">P{r.priority}</div><div><h3>{r.title}</h3><p>{r.detail}</p></div><div className="impact">{typeof r.impact === "number" && r.impact > 100 ? money(r.impact) : r.impact}</div></div>)}</div>;
}

export default function App() {
  const [searchFile, setSearchFile] = useState(null);
  const [targetFile, setTargetFile] = useState(null);
  const [bulkFile, setBulkFile] = useState(null);
  const [targetAcos, setTargetAcos] = useState(30);
  const [minClicks, setMinClicks] = useState(6);
  const [brandTerms, setBrandTerms] = useState("active green pro, agp");
  const [phase, setPhase] = useState("upload");
  const [parsed, setParsed] = useState(null);
  const [validation, setValidation] = useState(null);
  const [results, setResults] = useState(null);
  const [tab, setTab] = useState("summary");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const tabCounts = useMemo(() => results ? {
    negative: results.negativeCandidates.length, bidDown: results.bidDown.length, bidUp: results.bidUp.length,
    harvest: results.harvest.length, pause: results.pauseTargets.length, watch: results.watch.length,
    overlap: results.overlaps.length, placement: results.placements.length,
  } : {}, [results]);

  const reset = () => { setPhase("upload"); setParsed(null); setValidation(null); setResults(null); setTab("summary"); setError(""); };

  async function validateFiles() {
    if (!searchFile || !targetFile || !bulkFile) return;
    setLoading(true); setError("");
    try {
      const [searchWorkbook, targetWorkbook, bulkWorkbook] = await Promise.all([
        parseWorkbook(searchFile), parseWorkbook(targetFile), parseWorkbook(bulkFile),
      ]);
      const searchSheet = pickBestSheet(searchWorkbook, "searchTerm");
      const targetSheet = pickBestSheet(targetWorkbook, "targeting");
      const bulkSheet = pickBestSheet(bulkWorkbook, "bulk");
      const bulkSearchSheet = pickBestSheet(bulkWorkbook, "bulkSearch", false);
      const v = validateInputs({ searchWorkbook, targetWorkbook, bulkWorkbook, searchSheet, targetSheet, bulkSheet, bulkSearchSheet });
      setParsed({ searchWorkbook, targetWorkbook, bulkWorkbook, searchSheet, targetSheet, bulkSheet, bulkSearchSheet });
      setValidation(v); setPhase("validation");
    } catch (e) { setError(e.message || String(e)); }
    finally { setLoading(false); }
  }

  function runAnalysis() {
    try {
      const r = analyseAccount({
        searchRows: parsed.searchSheet.rows,
        targetingRows: parsed.targetSheet.rows,
        bulkRows: parsed.bulkSheet.rows,
        targetAcosPct: targetAcos,
        minClicks,
        brandTerms: brandTerms.split(",").map(x => x.trim()).filter(Boolean),
      });
      setResults(r); setPhase("results"); setTab("summary");
    } catch (e) { setError(e.message || String(e)); }
  }

  return <div className="app">
    <header><div className="mark"/><div><strong>Amazon Ads Audit Tool</strong><span>Version 2.0</span></div><div className="headerMeta">SEARCH TERM + TARGETING + BULK STRUCTURE</div></header>
    <main>
      {phase === "upload" && <>
        <section className="hero"><div className="eyebrow">ACCOUNT-LEVEL PPC ANALYSIS</div><h1>Validate the account first.<br/>Then build the action plan.</h1><p>V2 uses your actual Amazon Bulk Operations structure to interpret Search Term and Targeting performance.</p></section>
        <div className="uploadGrid">
          <UploadCard title="Search Term Report" hint="Sponsored Products Search Term report" file={searchFile} onChange={setSearchFile}/>
          <UploadCard title="Targeting Report" hint="Sponsored Products Targeting report" file={targetFile} onChange={setTargetFile}/>
          <UploadCard title="Bulk Sheet" hint="Amazon Sponsored Products Bulk Operations workbook" file={bulkFile} onChange={setBulkFile}/>
        </div>
        <div className="settings panel">
          <label><span>Target ACoS %</span><input type="number" min="5" max="80" value={targetAcos} onChange={e=>setTargetAcos(Number(e.target.value))}/><small>Drives Target CPA and bid recommendations.</small></label>
          <label><span>Minimum Click Evidence</span><input type="number" min="3" max="50" value={minClicks} onChange={e=>setMinClicks(Number(e.target.value))}/><small>Manual plan used &gt;5 clicks; V2 also checks spend vs Target CPA.</small></label>
          <label className="brandInput"><span>Brand terms</span><input value={brandTerms} onChange={e=>setBrandTerms(e.target.value)}/><small>Comma-separated; brand queries are reviewed rather than auto-negated.</small></label>
        </div>
        {error && <div className="errorBox">{error}</div>}
        <button className="primary full" disabled={!searchFile || !targetFile || !bulkFile || loading} onClick={validateFiles}>{loading ? "Validating Amazon files…" : "Validate Files →"}</button>
      </>}

      {phase === "validation" && validation && <>
        <div className="pageTop"><div><div className="eyebrow">STEP 1 · DATA VALIDATION</div><h2>Files parsed successfully</h2><p>Review the date alignment and account-structure match before generating recommendations.</p></div><button className="ghostBtn" onClick={reset}>← New Upload</button></div>
        <div className="validationCards">
          <ValidationCard title="Search Term Report" info={validation.search}/>
          <ValidationCard title="Targeting Report" info={validation.targeting}/>
          <ValidationCard title="Bulk Operations" info={validation.bulk} tone={validation.warnings.length ? "medium" : "good"}/>
        </div>
        <div className="kpiGrid validationKpis">

  <Kpi
    label="Targeting → Bulk Match"
    value={pct(validation.matchRate)}
    tone={validation.matchRate >= .9 ? "green" : "red"}
  />

  <Kpi
    label="Exact Structure Matches"
    value={Number(
      validation.matchQualityCounts?.["Name + Target"] || 0
    ).toLocaleString()}
    tone="green"
  />

  <Kpi
    label="Fallback Matches"
    value={Number(
      validation.matchQualityCounts?.["Name + Target (fallback)"] || 0
    ).toLocaleString()}
    tone="orange"
  />

  <Kpi
    label="Unmatched Targets"
    value={Number(
      validation.matchQualityCounts?.Unmatched || 0
    ).toLocaleString()}
    tone={
      (validation.matchQualityCounts?.Unmatched || 0)
        ? "orange"
        : "green"
    }
  />

  <Kpi
    label="Campaigns"
    value={validation.counts.campaigns}
  />

  <Kpi
    label="Ad Groups"
    value={validation.counts.adGroups}
  />

  <Kpi
    label="Positive Targets"
    value={validation.counts.positiveTargets.toLocaleString()}
  />

  <Kpi
    label="Negatives"
    value={validation.counts.negatives.toLocaleString()}
  />

  <Kpi
    label="Embedded Bulk STR Rows"
    value={validation.bulk.embeddedSearchRows.toLocaleString()}
  />

</div>
        {validation.warnings.length ? <div className="warningBox"><strong>Proceed with awareness</strong>{validation.warnings.map((w,i)=><div key={i}>• {w}</div>)}</div> : <div className="successBox">✓ No material alignment warnings detected.</div>}
        {validation.missingCampaigns.length > 0 && <div className="panel"><h3>Campaigns present in performance reports but missing from Bulk Sheet</h3><div className="chipList">{validation.missingCampaigns.map(c=><Badge key={c} tone="medium">{c}</Badge>)}</div></div>}
        {error && <div className="errorBox">{error}</div>}
        <div className="actionRow"><button className="ghostBtn" onClick={reset}>Replace Files</button><button className="primary" onClick={runAnalysis}>Proceed With Analysis →</button></div>
      </>}

      {phase === "results" && results && <>
        <div className="pageTop"><div><div className="eyebrow">AMAZON ADS OPTIMIZATION PLAN · V2</div><h2>Advertising audit & action plan</h2><p>Target ACoS {targetAcos}% · Target CPA {money(results.summary.targetCpa)} · {validation.counts.positiveTargets.toLocaleString()} positive targets analyzed against Bulk structure.</p></div><div className="topActions"><button className="ghostBtn" onClick={reset}>New Audit</button><button className="primary" onClick={()=>exportRecommendations(results, validation)}>Export Excel</button></div></div>
        <nav className="tabs">{TABS.map(([id,label])=><button key={id} onClick={()=>setTab(id)} className={tab===id?"active":""}>{label}{tabCounts[id] != null && <span>{tabCounts[id]}</span>}</button>)}</nav>
        <section className="panel resultsPanel">
          {tab === "summary" && <Summary results={results} validation={validation}/>} 
          {tab === "negative" && <><SectionTitle title="Waste / Negative Candidates" subtitle="Zero-order customer queries with click and Target-CPA evidence. Existing negatives and brand queries are excluded from automatic negative recommendations."/><Table rows={results.negativeCandidates} columns={commonSearchCols}/></>}
          {tab === "bidDown" && <><SectionTitle title="Lower Bid" subtitle="Converting targets above Target ACoS. Suggested bid cuts use the Bulk Sheet current bid and are capped at 40%, matching the logic of your manual optimization plan."/><Table rows={results.bidDown} columns={targetCols}/></>}
          {tab === "bidUp" && <><SectionTitle title="Increase Bid" subtitle="Efficient converting targets materially below Target ACoS. Increases are capped at 30% and further constrained when Top-of-Search share is already high."/><Table rows={results.bidUp} columns={targetCols}/></>}
          {tab === "harvest" && <><SectionTitle title="Harvest Opportunities" subtitle="Converting customer search terms with no active Exact target found in the Bulk Sheet."/><Table rows={results.harvest} columns={commonSearchCols}/></>}
          {tab === "pause" && <><SectionTitle title="Pause / Review Targets" subtitle="Deliberate targets with zero orders and sufficient evidence. Unmatched Bulk entities are marked for review rather than direct action."/><Table rows={results.pauseTargets} columns={targetCols}/></>}
          {tab === "watch" && <><SectionTitle title="Watch List" subtitle="Low-evidence, brand-relevant, already-negative, or otherwise uncertain items that should not be automatically changed."/><Table rows={results.watch} columns={results.watch.some(x=>x.term) ? commonSearchCols : targetCols}/></>}
          {tab === "overlap" && <><SectionTitle title="Structure / Target Overlap" subtitle="Same active target and match type in multiple locations. This is a routing review—not an assumption that Amazon is bidding against itself."/><Table rows={results.overlaps.map(x=>({...x,locationsText:x.locations.map(l=>`${l.campaignName} > ${l.adGroupName}`).join(" | ")}))} columns={[
            {key:"targeting",label:"Targeting"},{key:"targetType",label:"Type"},{key:"matchType",label:"Match"},{key:"count",label:"Locations"},{key:"locationsText",label:"Campaign / Ad Group Locations"},{key:"action",label:"Action",render:r=><Badge tone="action">{r.action}</Badge>},{key:"reason",label:"Reason"}
          ]}/></>}
          {tab === "placement" && <><SectionTitle title="Placement Review" subtitle="Uses the Bidding Adjustment entities already contained in the Bulk Sheet. V2 flags direction for review rather than pretending placement percentages can be optimized from ACoS alone."/><Table rows={results.placements} columns={[
            {key:"campaignName",label:"Campaign"},{key:"placement",label:"Placement"},{key:"percentage",label:"Current Adjustment",render:r=>`${Number(r.percentage||0).toFixed(0)}%`},{key:"clicks",label:"Clicks"},{key:"spend",label:"Spend",render:r=>money(r.spend)},{key:"sales",label:"Sales",render:r=>money(r.sales)},{key:"orders",label:"Orders"},{key:"acos",label:"ACoS",render:r=>pct(r.acos)},{key:"action",label:"Recommendation",render:r=><Badge tone="action">{r.action}</Badge>}
          ]}/></>}
          {tab === "plan" && <><SectionTitle title="Prioritized Action Plan" subtitle="Automatically summarizes the highest-impact categories in the same spirit as your manual Amazon Ads roadmap."/><ActionPlan rows={results.actionPlan}/></>}
        </section>
      </>}
    </main>
  </div>;
}

function SectionTitle({ title, subtitle }) {
  return <div className="sectionTitle"><h2>{title}</h2><p>{subtitle}</p></div>;
}
