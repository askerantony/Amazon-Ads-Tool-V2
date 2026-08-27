--- a/src/lib/engine.js
+++ b/src/lib/engine.js
@@
 function rowValue(row, names) {
   const entries = Object.entries(row || {});
   for (const name of names) {
     const found = entries.find(([k]) => normalizeText(k) === normalizeText(name));
-    if (found) return found[1];
+    // Amazon Bulk files often include both an editable field and an
+    // "(Informational only)" field. On keyword/product-target rows the
+    // editable Campaign/Ad Group name is commonly blank while the
+    // informational field contains the actual value. Do not stop at a
+    // matching-but-empty column; continue to the next alias.
+    if (found) {
+      const value = found[1];
+      if (value !== null && value !== undefined && String(value).trim() !== "") {
+        return value;
+      }
+    }
   }
   return "";
 }
@@
 export function validateInputs({ searchWorkbook, targetWorkbook, bulkWorkbook, searchSheet, targetSheet, bulkSheet, bulkSearchSheet }) {
@@
   const targetAgg = aggregateTargets(targetSheet.rows, tgCols, bulk);
   const matched = targetAgg.filter(x => x.matchQuality !== "Unmatched").length;
   const matchRate = targetAgg.length ? matched / targetAgg.length : 0;
+  const matchQualityCounts = targetAgg.reduce((acc, x) => {
+    acc[x.matchQuality] = (acc[x.matchQuality] || 0) + 1;
+    return acc;
+  }, {});
   const reportCampaigns = new Set(targetAgg.map(x => normalizeText(x.campaignName)).filter(Boolean));
-  const bulkCampaignsByName = new Set([...bulk.campaigns.values()].map(x => normalizeText(x.campaignName)).filter(Boolean));
+  // Campaign entities are not the only reliable place Amazon exposes the
+  // campaign name. Include ad groups and positive targets so validation does
+  // not label a campaign missing merely because the Campaign entity is sparse.
+  const bulkCampaignsByName = new Set([
+    ...[...bulk.campaigns.values()].map(x => normalizeText(x.campaignName)),
+    ...[...bulk.adGroups.values()].map(x => normalizeText(x.campaignName)),
+    ...bulk.targets.map(x => normalizeText(x.campaignName)),
+  ].filter(Boolean));
   const missingCampaigns = [...reportCampaigns].filter(x => !bulkCampaignsByName.has(x));
@@
-  if (matchRate < 0.95) warnings.push(`${((1 - matchRate) * 100).toFixed(1)}% of Targeting Report entities could not be matched to the uploaded Bulk Sheet.`);
+  if (matchRate < 0.95) warnings.push(`${((1 - matchRate) * 100).toFixed(1)}% of Targeting Report entities could not be matched to the uploaded Bulk Sheet. Unmatched rows will be review-only for actions that require a current bid or target state.`);
@@
     matchRate,
     matchedTargets: matched,
     totalTargetRows: targetAgg.length,
+    matchQualityCounts,
     missingCampaigns,
     warnings,
   };
 }
