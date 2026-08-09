# ATI Inc. (NYSE: ATI) — SEC Filing Cross-Reference Notes (Q2 FY2026)

CIK: 0001018963 (formerly Allegheny Technologies Inc / Allegheny Teledyne Inc). Fiscal year ends Sunday closest to Dec 31 (4-4-5 calendar).

## Filings pulled (SEC EDGAR)
- 10-Q, period ended 2026-06-28, filed 2026-08-06, accession 0001628280-26-054233 (the quarter this report covers)
- 10-Q, period ended 2026-03-29, filed 2026-04-30, accession 0001628280-26-028684
- 10-Q, period ended 2025-09-28, filed 2025-10-28, accession 0001628280-25-046675
- 10-K, fiscal year ended 2025-12-28, filed 2026-02-20, accession 0001628280-26-010140
- Form 4s: Kim Fields (Chair/President/CEO) 2026-06-02, 2026-06-22/23/24, 2026-07-07, 2026-07-28; James Robert Foster (SVP Finance & CFO) 2026-07-01

Quarterly figures not directly reported (Q4'25, and each quarter's discrete cash-flow/segment splits) were derived by subtracting YTD periods from the underlying XBRL facts (`companyconcept`/`companyfacts` API), cross-checked against the 10-Q "R" exhibit tables (Business Segments — Schedule of Sales and Profit by Segment) pulled from each filing's FilingSummary.xml index.

## Segment data (verified against transcript)
Two segments only: HPMC (High Performance Materials & Components) and AA&S (Advanced Alloys & Solutions) — no third/fourth segment, unlike Howmet's four-segment structure. Q2'26 HPMC external sales $637.1M / EBITDA $153.5M (24.10% margin); AA&S external sales $624.0M / EBITDA $147.6M (23.65% margin, or 22.05% excluding the $10M asset-sale gain) — both tie out exactly to the percentages Kim Fields/Rob Foster stated on the call.

"Adjusted EBITDA" per the call ($284M) reconciles to Segment EBITDA ($301.1M) minus Corporate expenses (-$14.9M) minus Closed operations and other income (-$1.8M) = $284.4M — near-exact match, used as the primary 雙層互證 anchor.

## Customer concentration
10-K, Note 1: "No single customer accounted for more than 10% of sales for any year presented." This is a genuine structural difference from peers like Howmet Aerospace (GE Aerospace ~14%, RTX ~10% of Q1'26 sales) — worth flagging as a risk-matrix positive, not just an omission.

## Raw materials / supply chain (10-K Item 1)
Source countries: Nickel (Canada, Norway, Japan, Finland, South Africa); Zirconium & Hafnium (U.S. and China); Cobalt (Norway, Japan); Chromium (U.K., South Africa, Germany, Turkey); Niobium (Brazil); Molybdenum (U.S., Brazil, China); Titanium sponge (Japan, Kazakhstan, Saudi Arabia, China).
Competitors — HPMC: Precision Castparts (Berkshire Hathaway), Howmet Aerospace, Carpenter Technology, Aubert & Duval. AA&S: Haynes International, VDM Metals (Acerinox).
Backlog at FYE2025 (10-K): total $3.7B (HPMC $3.1B / AA&S $0.6B), down from $3.9B at FYE2024 — this makes the Q2'26 quarter-end $4.4B figure (+18% YoY per the call) a sharp build during H1'26, not a continuation of a pre-existing uptrend.

## Debt / capital structure (10-Q Note — Debt)
Issued $450M 5.875% Senior Notes due 2033 on 2026-06-03 (net proceeds $443.1M); ~$350M of proceeds used to redeem the $350M 5.875% Senior Notes due 2027, which were called and fully redeemed 2026-07-08 (a subsequent event — the $350M was still sitting in current liabilities on the June 28 balance sheet, which is why short-term debt jumped from $33.2M to $383.6M quarter over quarter). Other notes outstanding: 7.25% due 2030 ($425M), 5.125% due 2031 ($350M), 4.875% due 2029 ($325M, next maturity Q4 FY2029). $200M term loan; $0 drawn on the $600M revolver. No credit-rating change disclosed this quarter (unlike Howmet's Fitch upgrade last quarter) — omitted from balance-sheet tab rather than guessed at.

Share repurchase: $700M program (Sep 2024) fully utilized + $5M of the new $500M program (Feb 2026) used as of June 28, 2026 → **$495M remaining**, matches Kim Fields' statement exactly. Q2'26 buyback $50M (0.3M shares); H1'26 $125M (0.8M shares).

No dividend — `PaymentsOfDividendsCommonStock`/`CommonStockDividendsPerShareDeclared` show $0 since 2017; ATI suspended its dividend that year and has not reinstated it.

## Form 4 insider activity (context for risk matrix)
CEO Kim Fields sold steadily under a Rule 10b5-1 plan across four filings, June 2 – July 28, 2026: ~59,749 sh @ $177.97–$182.75 (6/2) → ~20,693 sh @ $197.21–$202.48 (6/22–24) → 40,000 sh @ $179.25–$187.50 (7/7) → 31,757 sh @ $190.99 (7/28). Direct holding fell from an unknown pre-June base to 125,564 sh by 7/28 — a steady, pre-scheduled diversification program (all but the CFO's 7/1 filing are checked as 10b5-1(c) transactions), not a cluster of discretionary sales timed around the print. CFO James Foster: 163 sh disposed 7/1/2026 via code "M" (tax withholding on option exercise), balance 57,824 sh.

These trade prices (range $177.97–$193.82, rising over the two months) are the only market-price data available to this report and were used as the valuation-tab proxy for ATI's actual share price, which this report did not otherwise source.
