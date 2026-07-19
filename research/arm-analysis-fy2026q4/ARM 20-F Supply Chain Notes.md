# Arm Holdings plc — SEC Cross-Check Notes (FY2026 20-F + Q4 6-K + Form 4s)

CIK: 0001973239. Arm is a Foreign Private Issuer (UK/Cambridge) — files Form 20-F (annual) and Form 6-K (quarterly earnings/other events), not 10-K/10-Q.

## 20-F (FY ended March 31, 2026; filed 2026-05-26, accession 0001973239-26-000097)

### Customer concentration (Item 3.D Risk Factors)
- Top 5 customers (including Arm China and SoftBank Group) collectively accounted for **57%, 56%, 54%** of total revenue for FY26, FY25, FY24 respectively.
- Largest single customer, **Arm China**, accounted for **16%, 17%, 21%** of total revenue in FY26/FY25/FY24 — a declining trend even as the top-5 share ticked up, meaning concentration among the *rest* of the top 5 (likely SoftBank Group and a few large hyperscalers) is rising.
- A portion of near-term revenue is dependent on commercial arrangements with SoftBank Group and/or its affiliates (see $200m/quarter SoftBank technology-licensing contribution in the Q4 6-K).

### PRC revenue exposure
- Revenue from the PRC (incl. Hong Kong/Macau, excl. Taiwan) accounted for **18%, 19%, 22%** of total revenue for FY26/FY25/FY24 — a steadily declining share, both direct and via the Arm China relationship.

### Supply chain risk language (general, no single named foundry disclosed as sole-source in the risk-factor text itself)
- Arm's silicon products (Arm AGI CPU, CSS, chiplets) depend on **third-party foundry partners and contract manufacturers**; risk factors cite manufacturing yield risk, long lead times, wafer/packaging/testing purchase-obligation risk, and "complex supply chain dependencies on memory suppliers, packaging vendors, and system integrators."
- Separately, the **Q4 FY26 shareholder letter (6-K, Exhibit 99.1, filed 2026-05-06)** names specific ecosystem partners supporting the silicon expansion: *AWS, Broadcom, Google Cloud, Marvell, Microsoft, Micron, NVIDIA, Oracle, Samsung, SK Hynix, and TSMC* — this is the highest-confidence source for naming TSMC/Samsung/SK Hynix/Micron as foundry/memory partners (company's own investor letter), though exact wafer/capacity allocation by partner is not disclosed.
- Geopolitical/Taiwan concentration risk is separately flagged: "a significant portion of the world's semiconductor manufacturing is in Taiwan," and escalation in PRC-Taiwan tensions could disrupt the global semiconductor supply chain — relevant given TSMC's role as a named silicon partner.

### Confidence grading used in Analysis.html supply-chain tab
- **High confidence**: Arm China = 16% of FY26 revenue; PRC = 18% of FY26 revenue; top-5 = 57%; SoftBank = $200m/quarter license contribution; TSMC/Samsung/SK Hynix/Micron named as silicon-expansion partners; Supermicro/Lenovo/Quanta/ASRock named as AGI CPU rack ODM partners; Meta named as AGI CPU lead co-development partner — all sourced verbatim from the 20-F or the Q4 6-K shareholder letter.
- **Low confidence / illustrative**: any per-customer dollar allocation among AWS/Google/Microsoft/NVIDIA/enterprise licensees in the Sankey diagram — Arm does not disclose per-licensee revenue splits beyond Arm China and SoftBank, so those flows are illustrative relative-scale estimates only, clearly labeled as such in the visualization.

## Form 4 insider-trading signal (2026-05-13 through 2026-06-04, i.e. the ~4 weeks immediately following the May 6 earnings call)

All transactions in this window were **sales (code S) or tax-withholding dispositions (code F) — zero open-market purchases (code P)** by any officer or director:

| Date | Filer | Role | Activity |
|---|---|---|---|
| 2026-05-13 | Spencer Collins | Chief Legal Officer | Sold ~51,961 sh across 8 tranches, $207.77–$214.41 |
| 2026-05-19 | Rene Haas + 7 other officers/directors | CEO + CFO/CLO/CCO/CPO/Chief Architect/CAO + 4 non-employee directors | Annual RSU vest: awards (AA), net-share-settlement withholding (FD, all priced at $209.16), several directors' derivative-only vests |
| 2026-05-21 | Collins, Abbey, Child, Eaton | CLO, CCO, CFO, CPO | Sold ~99,917 sh combined, $215.00–$282.77 |
| 2026-05-22 | Abbey, Eaton | CCO, CPO | Sold ~12,460 sh combined, $287.03–$291.08 |
| 2026-05-26 | Abbey | CCO | Sold 2,300 sh @ $305.82 |
| 2026-06-01 | Abbey | CCO | Sold 4,200 sh @ $343.81 |
| 2026-06-03 | Abbey | CCO | Sold 6,566 sh combined, $383.73–$415.52 |
| 2026-06-04 | Bartels | CAO | Sold 11,306 sh @ $392.70 |

**Read**: the stock traded from ~$209 (May 19, the RSU-vest tax-withholding reference price) to ~$415 (June 3) — roughly a **+98% run in six weeks** following the earnings beat and the AGI CPU demand-doubling headline. Multiple officers (most visibly CCO William Abbey) sold progressively larger share counts at each successively higher price band on the way up. This pattern is consistent with **scheduled 10b5-1 sales tied to price triggers or routine post-vest tax-withholding**, not necessarily a bearish signal — but the sheer velocity and breadth (every reporting officer sold, zero bought) during a period when management was publicly reiterating unchanged, supply-constrained AGI CPU revenue guidance despite doubled "demand" is worth flagging in the risk matrix as a neutral-to-cautionary data point: insiders were net sellers into the entire post-earnings rally with no offsetting purchases.
