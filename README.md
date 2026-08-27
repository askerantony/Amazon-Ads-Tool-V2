# Amazon Ads Audit Tool — Version 2.0

This build replaces the original two-report audit logic with a three-source Amazon Ads model grounded in the actual exports used for Active Green Pro.

## Inputs
1. Sponsored Products Search Term Report
2. Sponsored Products Targeting Report
3. Amazon Bulk Operations workbook containing `Sponsored Products Campaigns`

## V2 workflow
1. Upload all three files.
2. Validate report schemas, dates, Bulk structure and Targeting-to-Bulk match rate.
3. Review warnings before analysis.
4. Generate recommendations:
   - Executive Summary
   - Waste / Negative Candidates
   - Lower Bid
   - Increase Bid
   - Harvest
   - Pause Targets
   - Watch
   - Structure / Overlap
   - Placement
   - Prioritized Action Plan
5. Export recommendations to Excel.

## Key logic changes
- Repairs malformed Amazon worksheet ranges before SheetJS conversion (`!ref=A1` issue).
- Bulk Sheet is the account-structure source of truth.
- Search Term Report is the customer-query performance source.
- Targeting Report is the target-level performance source.
- Target ACoS drives Target CPA and bid calculations.
- Suggested bid reductions use current Bulk bid, not observed CPC; reductions capped at 40%.
- Bid increases capped at 30%, and constrained when Top-of-Search impression share is already high.
- Zero-order traffic requires evidence from both clicks and spend vs Target CPA.
- Brand terms are routed to review instead of automatic negation.
- Existing negatives and existing Exact targets are checked before recommendations.
- Keyword, Auto and Product targets are classified separately.
- Structural overlap is labeled as routing review, not assumed self-competition.
- Direct browser-to-Anthropic API code is removed.

## Install / run
```bash
npm install
npm run dev
```

## Build
```bash
npm run build
```

## Deployment
Replace your current Vite app files with this project, commit to GitHub, and let Vercel redeploy from the connected repository.

## Important
V2 exports recommendations only. It intentionally does not yet create a directly uploadable Amazon Bulk Operations change file. Add human Accept/Reject controls before implementing that phase.
