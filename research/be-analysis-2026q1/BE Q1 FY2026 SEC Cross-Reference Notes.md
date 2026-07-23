# Bloom Energy Corporation (BE) Q1 FY2026 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK 0001664703) via `data.sec.gov` and `www.sec.gov/Archives`, fetched directly with a descriptive User-Agent (WebFetch was blocked by SEC with HTTP 403 on every attempt; all data below is a primary-source pull, not a mirror).

**Filings used:**
- **10-Q** filed 2026-04-29, period ended 2026-03-31 — accession `0001628280-26-028021` (`be-20260331.htm`)
- **10-K** filed 2026-02-09, FY2025 (period ended 2025-12-31) — accession `0001628280-26-006516` (`be-20251231.htm`)
- **8-K + Ex-99.1 press release + Ex-99.2 investor deck** filed 2026-04-28 (day of the print) — accession `0001628280-26-027913`
- **8-K** filed 2026-04-27 — accession `0001193125-26-179381`, re: prospectus supplement
- **424B7** filed 2026-04-27 — accession `0001193125-26-179296` — resale prospectus for the Oracle warrant shares
- **Form 3** — Simon Edwards (CFO), event date 2026-04-13, filed 2026-04-16 — accession `0001193125-26-167765`
- **Form 4s**: 40 filings in the 2026-03-01 to 2026-07-23 window, all read in full (full census, not a sample)
- Earnings call transcript at `research/be-analysis-2026q1/BE Q1 2026 Earnings Call Script.md` (Quartr, call 2026-04-28 5:00pm ET) — used for cross-referencing only

---

## 1. 10-Q Key Data — Three Months Ended March 31, 2026 vs. March 31, 2025

### Income Statement (GAAP, $ thousands except per-share)

| Line | Q1 2025 | Q1 2026 | YoY |
|---|--:|--:|--:|
| Product revenue | 211,869 | 653,348 | +208.4% |
| Installation revenue | 33,651 | 25,931 | −23.0% |
| Service revenue | 53,548 | 61,879 | +15.6% |
| Electricity revenue | 26,953 | 9,896 | −63.3% |
| **Total revenue** | 326,021 | **751,054** | **+130.4%** |
| Total cost of revenue | 237,314 | 525,510 | +121.4% |
| **Gross profit** | 88,707 | **225,544** | +154.3% |
| Gross margin % | 27.2% | **30.0%** | +2.8pp |
| R&D | 40,612 | 56,849 | +40.0% |
| Sales & marketing | 22,265 | 38,439 | +72.6% |
| G&A | 44,900 | 58,066 | +29.3% |
| **Income (loss) from operations** | (19,070) | **72,190** | — |
| Interest income | 8,553 | 20,601 | +140.9% |
| Interest expense | (14,411) | (8,604) | −40.3% |
| Equity in loss of unconsolidated affiliates | — | (17,002) | new |
| Other income, net | 2,048 | 6,197 | +202.6% |
| Gain (loss) on revaluation of embedded derivatives | (103) | 754 | — |
| **Income (loss) before income taxes** | (22,983) | **74,136** | — |
| Income tax provision | 431 | 445 | — |
| **Net income (loss)** | (23,414) | **73,691** | — |
| Less: NCI (Korean JV) | 400 | 3,038 | — |
| **Net income (loss) attributable to common stockholders** | **(23,814)** | **70,653** | — |
| Basic EPS | $(0.10) | **$0.25** | — |
| Diluted EPS | $(0.10) | **$0.23** | — |
| Weighted-avg diluted shares (000) | 230,210 | 319,708 | +38.9% |

Bloom was GAAP-profitable for the first time in a Q1 period. Effective tax rate 0.6% (full valuation allowance against U.S. DTAs still in place).

### GAAP vs. Non-GAAP (reconciliation from 8-K Ex-99.1, not the 10-Q itself)

| Metric | GAAP (10-Q) | Non-GAAP (8-K Ex-99.1) | Gap |
|---|--:|--:|--:|
| Gross margin | 30.0% | 31.5% | +1.5pp |
| Operating income | $72.19M | $129.71M | +$57.5M (+79.7%) |
| Operating margin | 9.6% | 17.3% | +7.7pp |
| Diluted EPS | $0.23 | $0.44 | +91% |
| Net income / Adj. Net Profit | $70.65M | $138.06M | +$67.4M |
| Adjusted EBITDA | n/a | $142.99M | — |

Bridge from GAAP operating income to non-GAAP: stock-based comp ($57,004K) + restructuring ($306K) + other ($211K) ≈ $57,521K ≈ $129,710K. **SBC alone is ~99% of the operating-income adjustment.** Total SBC recognized in the income statement: $57,004K (+77.0% YoY from $32,201K); cash-flow-statement SBC line is $48,215K (~$8.8M capitalized into inventory/PP&E, a normal reconciling item).

### Balance Sheet ($ thousands)

| Line | Dec 31, 2025 | Mar 31, 2026 | Δ |
|---|--:|--:|--:|
| Cash and cash equivalents | 2,454,108 | 2,491,433 | +37,325 |
| Restricted cash (current + LT) | 27,472 | 26,851 | −621 |
| Accounts receivable, net | 371,796 | 359,406 | −12,390 |
| Contract assets (current) | 178,928 | 242,595 | +63,667 |
| Inventories | 643,306 | 732,528 | +89,222 |
| Total current assets | 3,730,567 | 3,954,536 | +223,969 |
| Property, plant & equipment, net | 398,507 | 401,088 | +2,581 |
| Investments in unconsolidated affiliates | 10,037 | 23,261 | +13,224 |
| Total assets | 4,396,711 | 4,664,729 | +268,018 |
| Deferred revenue & customer deposits (current) | 100,975 | 194,094 | +93,119 (+92%) |
| Total current liabilities | 623,832 | 786,804 | +162,972 |
| Recourse debt (long-term) | 2,613,726 | 2,598,676 | −15,050 |
| Total liabilities | 3,603,748 | 3,716,721 | +112,973 |
| Total stockholders' equity attrib. to common | 768,641 | 921,469 | +152,828 |
| Total stockholders' equity | 792,963 | 948,008 | +155,045 |

Deferred revenue and customer deposits (current) grew 92% QoQ — corroborates the CFO's "customer prepayments to reserve capacity" comment.

**Debt (Note 8, $ thousands unpaid principal):**
- 0% Convertible Senior Notes due Nov 2030: $2,500,000 (net carrying $2,444,939)
- 3.0% Green Convertible Senior Notes due June 2029: $75,125 ($73,594)
- 3.0% Green Convertible Senior Notes due June 2028: $81,236 ($80,143) — down from $99,655 at Dec 31, 2025 because $18,163K of this tranche converted to equity in Q1 (976,992 shares issued)
- Total recourse debt: $2,598,676K net carrying; non-recourse (Korean JV term loans): $3,959K
- All covenants in compliance both period-ends

**PP&E breakdown** ($000): Vehicles/machinery/equipment 216,525; Energy Server systems 147,689; Leasehold improvements 131,857; Construction-in-progress 91,541; Buildings 53,513; Computers/software 35,644; Furniture 10,693; gross 687,462 less accumulated depreciation (286,374) = 401,088 net.

### Cash Flow ($ thousands)

| Line | Q1 2025 | Q1 2026 |
|---|--:|--:|
| Net cash from operating activities | (110,682) | 73,610 |
| Purchase of PP&E | (14,259) | (26,182) |
| Investments in unconsolidated affiliates | — | (19,848) |
| Net cash used in investing activities | (14,216) | (45,939) |
| Proceeds from issuance of common stock | 7,651 | 15,835 |
| Repayment of financing obligations | (2,671) | (7,972) |
| Net cash from financing activities | 5,130 | 7,057 |
| Net increase (decrease) in cash | (119,613) | 36,704 |
| End of period cash + restricted cash | 831,358 | 2,518,284 |

"$2.52 billion in total cash" and "$73.6 million operating cash inflow" (Simon Edwards, call) both match the 10-Q exactly.

### Share Count / Dilution (Note 15)

- Weighted-avg basic shares: 281,719K; diluted: 319,708K (dilutive effect: convertible notes +21,371K, stock options/awards +16,618K)
- Diluted EPS numerator required a $4,200K debt-interest add-back, net of tax (if-converted method)
- Shares outstanding Mar 31, 2026: 284,207,963 (all Class A; zero Class B; zero preferred)
- Q1 2025 (net-loss period): 67,220K potential shares excluded as antidilutive — fully diluted share count roughly triples once profitable and converts go in the money

### Customer / Credit Concentration (Note 1)

> "During the three months ended March 31, 2026, revenue from two customers, the first of which is our related party, accounted for approximately 50% and 12% of our total revenue. During the three months ended March 31, 2025, two customers represented approximately 37% and 29% of our total revenue."

Ties to the revenue-statement footnote: "related party revenue of $373.3 million" for Q1 2026 — $373.3M / $751.054M ≈ 49.7% ≈ the disclosed "50%." The related party is Bloom's own equity-method investee — **Bolt US Class A JVCo LLC / Bolt US JVCo LLC**, Brookfield-sponsored "AI Fund" joint ventures in which Bloom holds 9.9% stakes (Note 7). **Half of Q1 revenue was product sold into project-financing vehicles Bloom itself co-owns and has a financial interest in**, not an arm's-length third party.

AR credit risk: one customer (same related party) ≈30% of AR at Mar 31, 2026 (vs. three customers at 41%/17%/15% at Dec 31, 2025 — concentration into a single counterparty increased sequentially even as total AR fell).

Geographic: U.S. revenue rose from 56% (Q1 2025) to 91% (Q1 2026) of total. **No dollar-level segment breakout of "data center" vs. "C&I" revenue exists anywhere in the 10-Q** — Bloom discloses revenue only by product/installation/service/electricity category, never by end-market vertical.

### Notable MD&A / Note Items Not Mentioned on the Call

- **$19.7 million specific warranty reserve for "identified product issues"**, booked in Q1 2026 within cost of product revenue: *"As of March 31, 2026, Bloom established a specific warranty reserve of $19.7 million for identified product issues, accounted for as an assurance-type warranty."* No further detail on what the issues are. Recomputed product gross margin excluding this reserve ≈ 37.3%, vs. 34.3% GAAP / 35.3% non-GAAP as reported — roughly 300bp higher absent the charge.
- Purchase commitments with suppliers: "no material open purchase orders... expected to be realized within more than a 12-month period and are not cancellable" at both period-ends — unlike peers with large disclosed multi-year supply commitments, Bloom discloses essentially none, consistent with a 90-day cancellable PO supplier model.
- Performance guarantee payments: $8.4M (Q1 2026) vs. $11.6M (Q1 2025) — favorable YoY trend, moving opposite to the new warranty reserve above.
- Item 1A Risk Factors: "There were no material changes in risk factors as disclosed in our 2025 Form 10-K."

---

## 2. Insider Transactions (Form 4) — Full Census, 2026-03-01 to 2026-07-23 (40 filings)

**Reporting persons:** KR Sridhar (Founder/Chairman/CEO), Simon Stephen Edwards (CFO, Form 3 event 04/13/2026), Maciej Kurzymski (CAO, Acting PFO pre-Edwards), Shawn Marie Soderberg (Chief Legal Officer, signs most other insiders' Form 4s as attorney-in-fact), Satish Chitoori (COO), Aman Joshi (CCO), plus non-employee directors (Zervigon, Pinkus, Warner, Immelt, Chambers, Bush, Snabe, Boskin, Burger).

**CEO KR Sridhar:** sold zero shares on the open market in this window. Only activity was equity-plan vesting — 300,000 PSUs vested 03/03/2026 but deferred to Jan 1, 2030 under the deferred comp plan (no sale); 80,000 RSUs vested 05/21/2026, no corresponding sale.

**CFO Simon Edwards:** Form 3 (event 04/13/2026) shows zero beneficial ownership at start. Only Form 4 (filed 05/22/2026) is a grant of 10,000 new-hire RSUs — no sale.

**CAO/Acting PFO, COO, CCO, CLO — 10b5-1 sales**, all executed under plans adopted November 2025 (pre-print), i.e., mechanically scheduled, not discretionarily timed:

| Date | Insider | Shares sold | Price |
|---|---|--:|--:|
| 03/19/2026 | Soderberg (CLO) | 15,410 | $150.47 |
| 04/01/2026 | Joshi (CCO) | 10,000 | $135.88 |
| 04/14/2026 | Chitoori (COO) | 20,000 | $204.23 |
| 04/14/2026 | Soderberg (CLO) | 30,000 | $204.23 |
| 04/15/2026 | Soderberg (CLO) | 25,000 | $225.13 |
| — earnings call 04/28/2026 5:00pm ET — | | | |
| 04/29/2026 | Soderberg (CLO, via trust) | 35,000 | $279.00 |
| 05/07/2026 | Bush (director) | 25,000 | $266.96 |
| 05/13/2026 | Kurzymski (CAO) | 6,229 | $293.36 |
| 05/28/2026 | Chambers (director, via LLC; plan adopted 02/26/2026) | 55,000 | $297.69 |
| 07/01/2026 | Joshi (CCO) | 8,343 | $300.37 |

Stock price trajectory implied by Form 4 prices: $150.47 (03/19) → $135.88 dip (04/01) → $204.23–$225.13 (04/14–15, two weeks pre-print) → $231.17 (04/24, per 424B7 reference price) → **$279.00 the day after the print (04/29, +21%)** → sustained $259–$300 through July (still rising: $297.69 on 05/28, $300.37 on 07/01). This is roughly a doubling from ~$136–150 in March to ~$280–300 by late April/July, concentrated around the earnings print — the inverse pattern of a post-print selloff; here the filed insider-sale prices corroborate a genuine, sustained post-earnings re-rating. Derived from Form 4 transaction prices only, not an independent market feed.

Housekeeping flag: 6 director Form 4s filed 2026-04-02 report a transaction dated 05/14/2025 (prior-year RSU grant) — ~10.5 months late, no explanation given beyond standard boilerplate. Routine equity awards, not sales, no valuation signal. A Schedule 13G/A was filed 2026-04-24 by a passive institutional holder (not read in detail, outside Section-16 scope).

---

## 3. Supply Chain & Manufacturing Disclosure (FY2025 10-K, confidence: medium — largely qualitative)

**Manufacturing footprint:**
- Fremont, CA: 326,000 sq ft total (89,000 operational since Apr 2021 + 164,000 leased since Jul 2022, expires Feb 2036 + 73,000 sq ft research/hydrogen center opened Jun 2022)
- Newark, DE: 454,000 sq ft leased + 178,000 sq ft owned ("first purpose-built... copy-exact duplication" center, 25 additional acres available for expansion/supplier co-location) + Repair & Overhaul facility 133,000 sq ft (leases expiring Dec 2026/Apr 2027/Jul 2030)
- Republic of Korea: full-assembly JV with SK ecoplant
- Sunnyvale (60,000 sq ft) and Mountain View (44,000 sq ft), CA facilities vacated during 2025, consolidated into Fremont

**Suppliers:** no suppliers named individually anywhere in the 10-K or 10-Q, no supplier-concentration percentages disclosed. Generic language: "high-quality suppliers that support automotive, semiconductors and other traditional manufacturing organizations... rare earth elements, specialty alloys and industrial commodities... generally have multiple sources of supply... except in cases where we have specialized technology and material property requirements." Global supply base across North America, Asia, Europe, India.

**China/Taiwan — internally inconsistent within the same 10-K:**
- Item 1 Business: "our supply chain does not have significant exposure to China"
- Item 1A Risk Factors: "China as a country supplies multiple components including rare earth metals and compounds... that are part of our tier 2 and tier 3 sub-assembly suppliers"; "our products contain components which are supplied principally from Taiwan"

The most coherent read is "not significant at tier-1 (direct) suppliers, but present and unquantified two tiers upstream" — but the filing does not reconcile the two statements itself, and neither gives a percentage. Tariff risk is discussed qualitatively only, no dollar impact disclosed; Q1 10-Q MD&A names steel alloys, specialty metals, electronic components, natural-gas-linked inputs, and rare-earth-dependent materials as experiencing price fluctuations.

**Customer concentration cross-reference:** no customers named by name in the revenue-concentration note (only "two customers," one being the related-party JV). The only customer named anywhere in the 10-Q is **Oracle**, and only in the warrant/strategic-partnership note (Section 4) — never in the revenue-concentration disclosure, meaning Oracle-driven revenue cannot be separately sized from the filing.

---

## 4. Discrepancies Between the Earnings Call and the SEC Filings (ranked by materiality)

**(1) The marquee announcement of the call — Oracle "Project Jupiter," "up to 2.45 GW," "100% Bloom" — appears nowhere in any Bloom SEC filing found.** Searched exhaustively across the full 10-Q, Ex-99.1 press release, and 8-K body: the strings "Jupiter," "2.45," "gigawatt," and "GW" do not appear anywhere. "Oracle" appears only in the 10-Q's warrant footnote, never in the press release or 8-K body. No separate 8-K announcing Project Jupiter was found under Bloom's own CIK (it is possible Oracle filed any Jupiter-specific disclosure itself, which would not appear under Bloom's CIK — this research covered only Bloom's own EDGAR filings). Treat the 2.45 GW / "100% Bloom" figures as call-only claims, unconfirmed by any Bloom SEC filing found in this research.

**(2) The Oracle warrant is a large, rapidly growing reduction of revenue, never mentioned on the call.** 10-Q Note 3 ("Commitment to Issue Share-Based Consideration Payable to Customer's Customer") discloses a warrant for 3,531,073 shares at $113.28 strike tied to the October 2025 Oracle strategic partnership. Estimated fair value: $55.9M (Dec 31, 2025) → $183.6M (Mar 31, 2026) → $261.3M (April 9, 2026 grant date, Note 16 Subsequent Events). **$12.8 million was already recognized as a reduction of revenue in Q1 2026** from this instrument, and the full $261.3M grant-date value will continue reducing revenue as Energy Server systems are delivered under the Oracle arrangement. The 424B7 (filed 04/27/2026, day before the call) confirms Oracle as counterparty and the exact share count, with BE trading at $231.17 on 04/24/2026 as reference price. Neither KR Sridhar nor Simon Edwards mentioned this warrant or its status as a contra-revenue item directly netted against the Oracle relationship being celebrated as the quarter's highlight.

**(3) A $19.7M "specific warranty reserve for identified product issues" was booked in Q1 2026 and never mentioned on the call.** See Section 1. Product margin excluding the reserve would be roughly 300bp higher than either the GAAP or non-GAAP figures discussed. No detail on what the "identified product issues" are or whether resolved.

**(4) GAAP vs. non-GAAP gap is large and driven almost entirely by one line (SBC), and the GAAP milestone itself was underplayed.** Diluted EPS $0.23 GAAP vs. $0.44 non-GAAP (call) — a 91% gap, nearly all stock-based compensation (+77% YoY). The reconciliation is fully filed same-day (not a disclosure failure), but this is arguably the first quarter Bloom has ever generated positive GAAP operating income and operating cash flow in a seasonally weak Q1 — a bigger, filing-native story than the non-GAAP beat-and-raise framing used on the call.

**(5) "70% to 80% of our business" from repeat customers, and "well more than half of our current data center backlog" from non-Oracle hyperscalers — unverifiable against any filed figure.** Bloom discloses no revenue segmentation by end-market vertical or by new-vs-repeat customer anywhere in the 10-Q or 10-K.

**(6) Minor governance-timing detail:** the April 28 8-K is still signed by Maciej Kurzymski as "Acting Principal Financial Officer," even though Simon Edwards' Form 3 event date (04/13/2026) predates the filing and he spoke on the call as CFO — the certifying-officer role had evidently not yet formally transferred to him as of the filing date.

**Cleanly corroborated (no discrepancy):** revenue $751.1M/+130.4% YoY (GAAP = the only revenue figure, no non-GAAP distinction exists); operating cash flow $73.6M; total cash $2.52B; gross margin improvement ~280bp (true on both a GAAP and non-GAAP basis — a rare case where GAAP and non-GAAP tell an identical trend story despite differing absolute levels); deferred revenue/customer deposits +92% QoQ (corroborates "customer prepayments to reserve capacity"); CFO's "joined two weeks ago" (Form 3 event date is exactly 15 days before the call).

---

僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
