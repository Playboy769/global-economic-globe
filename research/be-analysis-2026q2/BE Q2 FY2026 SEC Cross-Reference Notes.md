# Bloom Energy Corporation (BE) Q2 FY2026 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK **0001664703**, verified via `data.sec.gov/submissions/CIK0001664703.json` — `entityName` returned exactly **"Bloom Energy Corp"**, SIC 3620) via `data.sec.gov` and `www.sec.gov/Archives`, fetched directly with `curl` and a descriptive `User-Agent` header (the WebFetch tool was blocked by SEC with HTTP 403 on every attempt, consistent with the Q1 notes — all data below is a primary-source pull via direct HTTP fetch, not a mirror or paraphrase tool).

**Filings used:**
- **10-Q** filed 2026-07-28, period ended 2026-06-30 — accession `0001628280-26-050247` (`be-20260630.htm`)
- **8-K + Ex-99.1 press release + Ex-99.2 investor deck** filed 2026-07-28 (day of the print) — accession `0001628280-26-050150`
- **10-Q** filed 2025-10-28, period ended 2025-09-30 — accession `0001628280-25-046844` (`be-20250930.htm`) — fetched to resolve an XBRL EPS tagging anomaly (see trend table notes) and to get Q3 2025 net-income-to-common directly from the statement
- **10-K** filed 2026-02-09, FY2025 (period ended 2025-12-31) — accession `0001628280-26-006516` (`be-20251231.htm`) — same filing already covered in the Q1 notes; re-used here for FY2025 full-year figures and to confirm no new 10-K has been filed since Q1
- **Form 4s**: 8 filings in the 2026-06-01 to 2026-08-05 window (data pull performed 2026-07-29, one day after the earnings call — see Section 2 for the resulting coverage gap)
- **XBRL** `companyfacts` API (`data.sec.gov/api/xbrl/companyfacts/CIK0001664703.json`) and `companyconcept` API — used for the 6-quarter trend table
- Earnings call transcript at `research/be-analysis-2026q2/BE Q2 2026 Earnings Call Script.md` (Quartr, call 2026-07-28 5:00pm ET) — used only for cross-referencing specific claims against the filings, per the framework's "雙層互證" requirement

---

## 過去 6 季財務趨勢（10-Q XBRL）— NEW, per updated SEC cross-reference rule for US companies

**Coverage note (read before using this table):** the SEC's XBRL `companyfacts`/`companyconcept` APIs had **not yet ingested the Q2 FY2026 10-Q** as of this pull (2026-07-29, one day after the 2026-07-28 filing) — every tag queried for the period 2026-04-01/2026-06-30 came back empty. Q2 FY2026 figures in this table are therefore sourced **directly from the filed 10-Q's financial statements** (Item 1), not from XBRL, and are cited as such. All other quarters are XBRL-sourced. **All figures below are GAAP** (XBRL only carries GAAP tags) — the earnings call and 8-K press release speak in **non-GAAP** terms (Adjusted EBITDA, non-GAAP operating income, non-GAAP EPS, etc.); do not compare these GAAP figures directly against call commentary without checking Section 1's GAAP-vs-non-GAAP reconciliation table below.

| Metric ($000 unless noted) | Q1 FY25 (3mo end 3/31/25) | Q2 FY25 (3mo end 6/30/25) | Q3 FY25 (3mo end 9/30/25) | Q4 FY25 (derived, 3mo end 12/31/25) | Q1 FY26 (3mo end 3/31/26) | Q2 FY26 (3mo end 6/30/26) |
|---|--:|--:|--:|--:|--:|--:|
| Total revenue | 326,021 | 401,242 | 519,048 | 777,683 (derived) | 751,054 | 1,065,365 |
| Total cost of revenue | 237,314 | 294,119 | 367,373 | 537,788 (derived) | 525,510 | 709,793 |
| Gross profit | 88,707 | 107,123 | 151,675 | 239,895 (derived) | 225,544 | 355,572 |
| Gross margin % | 27.2% | 26.7% | 29.2% | 30.8% (derived) | 30.0% | 33.4% |
| R&D expense | 40,612 | 40,768 | 48,724 | 55,889 (derived) | 56,849 | 58,873 |
| G&A expense | 44,900 | 45,792 | 53,110 | 54,575 (derived) | 58,066 | 71,417 |
| Sales & marketing expense | 22,265 | 24,066 | 41,995 | 41,902 (derived) | 38,439 | 43,045 |
| Operating income (loss) | (19,070) | (3,503) | 7,846 | 87,529 (derived) | 72,190 | 182,237 |
| Net income (loss) attrib. to common stockholders | (23,814) | (42,619) | (23,093) | 1,092 (derived) | 70,653 | 196,290 |
| Diluted EPS | $(0.10) | $(0.18) | $(0.10)* | not derived† | $0.23 | $0.62 |
| Cash and cash equivalents (period-end) | 794,751 | 574,764 | 595,055 | 2,454,108 | 2,491,433 | 2,666,859 |
| Net cash from operating activities | (110,682) | (213,111) (derived) | 19,669 (derived) | 418,073 (derived) | 73,610 | 226,432 (derived) |
| Total assets (period-end) | 2,607,984 | 2,530,422 | 2,638,199 | 4,396,711 | 4,664,729 | 5,628,401 |
| Total liabilities (period-end) | 2,006,529 | 1,910,992 | 1,960,720 | 3,603,748 | 3,716,721 | 3,987,730 |
| Stockholders' equity attrib. to common (period-end) | 578,271 | 594,581 | 653,070 | 768,641 | 921,469 | 1,611,998 |
| Recourse debt, long-term (period-end) | 1,016,182 | 1,129,190 | 1,130,892 | 2,613,726 | 2,598,676 | 2,470,704 |

\* **XBRL tagging anomaly flagged and resolved by hand.** The `companyfacts` API reports Q3 2025 diluted EPS as **-100** (three months) and **-380** (nine months) under tag `EarningsPerShareDiluted` — off by exactly 1000x from the true values. Confirmed by pulling the actual Q3 2025 10-Q (accession `0001628280-25-046844`) statement of operations directly: *"Net loss per share available to common stockholders, basic and diluted $(0.10) $(0.06) $(0.38) $(0.59)"* for three/nine months ended 9/30/2025 vs. 9/30/2024. The table above uses the **10-Q-sourced, correct value of $(0.10)**, not the erroneous XBRL tag. This is a filer-side inline-XBRL scaling/decimals error in the original submission, not a transcription error on my part — flagging so whoever builds charts off the raw XBRL API doesn't silently plot -100 and -380.

† **Q4 FY2025 diluted EPS is not independently derivable** by subtracting 9-month YTD EPS from FY EPS — EPS is not additive across periods because the weighted-average diluted share count differs quarter to quarter (and swings especially hard here: ~232.6M weighted avg diluted shares for 9mo 2025 vs. a full-year 240.4M, itself blended against a share count that later roughly triples with post-Q1-2026 convertible-note conversions). Only the FY2025 net-income and share-count denominators are on record; a true Q4-only weighted-average share count was not found in the filings reviewed. Q4 FY2025 net income to common ($1,092K, derived) implies Q4 2025 was — narrowly — Bloom's first sequentially profitable quarter, i.e., one quarter *before* the "first GAAP-profitable Q1" framing noted in the Q1 FY2026 cross-reference notes; that framing was specifically about Q1-over-Q1 seasonality, and remains accurate on that basis, but Q4 2025 was already GAAP net-income-positive as a plain calendar quarter.

**Derivation method for "(derived)" cells:** Q4 FY2025 = FY2025 (10-K, `2025-01-01`–`2025-12-31`) minus 9-month YTD 2025 (`2025-01-01`–`2025-09-30`, from the Q3 2025 10-Q). Operating cash flow for Q2 FY2025, Q3 FY2025, and Q2 FY2026 = each quarter's cumulative YTD figure minus the prior quarter's cumulative YTD figure (only cumulative YTD operating cash flow is XBRL-tagged in any given 10-Q; Bloom's 10-Qs never present a discrete-3-month cash flow column). Example: Q2 FY26 OCF = 6-month YTD OCF per the Q2 10-Q ($300,042K) minus Q1 FY26 OCF ($73,610K, itself a directly-tagged discrete quarter) = $226,432K — this matches the CFO's call statement of "$226 million" operating cash flow for the quarter almost exactly, a clean corroboration.

---

## 1. 10-Q Key Data — Three and Six Months Ended June 30, 2026 vs. 2025

### Income Statement (GAAP, $ thousands except per-share)

| Line | Q2 2025 | Q2 2026 | YoY | 6mo 2025 | 6mo 2026 |
|---|--:|--:|--:|--:|--:|
| Product revenue | 296,611 | 935,413 | +215.4% | 508,480 | 1,588,761 |
| Installation revenue | 37,372 | 50,978 | +36.4% | 71,023 | 76,909 |
| Service revenue | 54,449 | 69,023 | +26.8% | 107,997 | 130,902 |
| Electricity revenue | 12,810 | 9,951 | −22.3% | 39,763 | 19,847 |
| **Total revenue** | 401,242 | **1,065,365** | **+165.6%** | 727,263 | **1,816,419** |
| Total cost of revenue | 294,119 | 709,793 | +141.3% | 531,433 | 1,235,303 |
| **Gross profit** | 107,123 | **355,572** | +231.9% | 195,830 | **581,116** |
| Gross margin % | 26.7% | **33.4%** | +668bp | 26.9% | **32.0%** |
| R&D | 40,768 | 58,873 | +44.4% | 81,380 | 115,722 |
| Sales & marketing | 24,066 | 43,045 | +78.9% | 46,331 | 81,484 |
| G&A | 45,792 | 71,417 | +56.0% | 90,692 | 129,483 |
| **Income (loss) from operations** | (3,503) | **182,237** | — | (22,573) | **254,427** |
| Interest income | 6,623 | 20,881 | +215.3% | 15,176 | 41,482 |
| Interest expense | (14,440) | (8,906) | −38.3% | (28,851) | (17,510) |
| Equity in earnings (loss) of unconsolidated affiliates | — | 4,346 | new | — | (12,656) |
| Other income, net | 2,373 | 2,307 | −2.8% | 4,421 | 8,504 |
| Loss on extinguishment of debt | (32,340) | — | — | (32,340) | — |
| Gain (loss) on revaluation of embedded derivatives | 112 | (539) | — | 9 | 215 |
| **Income (loss) before income taxes** | (41,175) | **200,326** | — | (64,158) | **274,462** |
| Income tax provision | 1,017 | 1,470 | — | 1,448 | 1,915 |
| **Net income (loss)** | (42,192) | **198,856** | — | (65,606) | **272,547** |
| Less: NCI (Korean JV) | 427 | 2,566 | — | 827 | 5,604 |
| **Net income (loss) attributable to common stockholders** | **(42,619)** | **196,290** | — | (66,433) | **266,943** |
| Basic EPS | $(0.18) | **$0.68** | — | $(0.29) | **$0.94** |
| Diluted EPS | $(0.18) | **$0.62** | — | $(0.29) | **$0.85** |
| Weighted-avg diluted shares (000) | 232,542 | 323,331 | +39.0% | 231,383 | 323,649 |

Bloom's first-ever quarter over $1 billion in revenue and first GAAP-profitable Q2 (the Q1 FY2026 notes already flagged Q1's first-GAAP-profitable-Q1 milestone; per the derived Q4 FY2025 figure in the trend table above, Q4 2025 was actually the first plain-calendar GAAP-profitable quarter, narrowly, at +$1.1M).

### GAAP vs. Non-GAAP (reconciliation from 8-K Ex-99.1, not the 10-Q itself)

| Metric | GAAP (10-Q) | Non-GAAP (8-K Ex-99.1) | Gap |
|---|--:|--:|--:|
| Revenue | $1,065,365K | $1,065,365K (no GAAP/non-GAAP distinction on revenue) | — |
| Cost of revenue | $709,793K | $700,002K | −$9,791K |
| Gross profit | $355,572K | $365,363K | +$9,791K |
| Gross margin | 33.4% | 34.3% | +90bp |
| Operating income | $182,237K | $239,642K | +$57,405K (+31.5%) |
| Operating margin | 17.1% | 22.5% | +540bp |
| Diluted EPS | $0.62 | $0.78 | +25.8% |
| Net income / Adjusted Net Profit | $196,290K | $248,209K | +$51,919K (+26.5%) |
| Adjusted EBITDA | n/a (not a GAAP measure) | $253,388K (23.8% of revenue) | — |

Bridge from GAAP operating income to non-GAAP operating income: stock-based comp $56,402K + restructuring $848K + other $153K = $57,403K ≈ $57,405K disclosed gap (rounding). **SBC alone is ~98.3% of the operating-income adjustment**, essentially identical in composition to Q1's ~99%. Note the GAAP-vs-non-GAAP *percentage* gap on diluted EPS narrowed sharply from Q1 (+91%) to Q2 (+26%) — not because SBC shrank (it's flat sequentially at $56.4M vs $57.0M) but because the GAAP base got much larger (net income to common $196.3M vs $70.7M), so the same-size SBC add-back matters proportionally less.

### Balance Sheet ($ thousands)

| Line | Dec 31, 2025 | Jun 30, 2026 | Δ |
|---|--:|--:|--:|
| Cash and cash equivalents | 2,454,108 | 2,666,859 | +212,751 |
| Restricted cash (current + LT) | 27,472 | 21,649 | −5,823 |
| Accounts receivable, net | 371,796 | 458,126 | +86,330 |
| Contract assets (current) | 178,928 | 365,461 | +186,533 |
| Inventories | 643,306 | 758,188 | +114,882 |
| Customer consideration asset (current + LT, new line, Oracle warrant) | — | 306,500 | new |
| Total current assets | 3,730,567 | 4,590,062 | +859,495 |
| Property, plant & equipment, net | 398,507 | 443,388 | +44,881 |
| Investments in unconsolidated affiliates | 10,037 | 28,090 | +18,053 |
| Total assets | 4,396,711 | 5,628,401 | +1,231,690 |
| Deferred revenue & customer deposits (current) | 100,975 | 327,145 | +224,170 (+222%) |
| Accrued warranty | 20,013 | 77,797 | +57,784 |
| Total current liabilities | 623,832 | 1,123,187 | +499,355 |
| Recourse debt (long-term) | 2,613,726 | 2,470,704 | −143,022 |
| Total liabilities | 3,603,748 | 3,987,730 | +383,982 |
| Total stockholders' equity attrib. to common | 768,641 | 1,611,998 | +843,357 |
| Total stockholders' equity | 792,963 | 1,640,671 | +847,708 |
| Common shares issued & outstanding | 280,045,459 | 293,354,001 | +13,308,542 |

Deferred revenue and customer deposits (current) grew 222% since year-end — an even sharper acceleration than Q1's 92% QoQ growth, continuing to corroborate "customer prepayments to reserve capacity." Share count grew ~13.3M shares in the quarter/half; at least 2,154,231 of those shares are directly traceable to the Oracle warrant cashless exercise on May 1, 2026 (1,905,433 base + 248,798 inducement shares, see below); the remainder reflects continued convertible-note conversions (see Debt) plus ordinary equity compensation issuance.

**Debt (Note 8, $ thousands unpaid principal):**
- 0% Convertible Senior Notes due Nov 2030: unchanged at $2,500,000 (net carrying $2,447,915)
- 3.0% Green Convertible Senior Notes due June 2029: $26,971 ($26,697 net) — down from $75,125 at Dec 31, 2025; **~$48.2M of this tranche converted to equity during H1 2026**
- 3.0% Green Convertible Senior Notes due June 2028: $787 ($778 net) — down from $99,655 at Dec 31, 2025; **~$98.9M of this tranche converted to equity during H1 2026** (of which the Q1 notes already reported $18,163K/976,992 shares converted in Q1 alone, implying roughly a further ~$80.7M converted specifically in Q2)
- Total recourse debt: $2,470,704K net carrying (LT); total debt incl. non-recourse Korean JV term loan: $2,477,973K
- All covenants in compliance both period-ends: *"We and all of our subsidiaries were in compliance with all financial covenants as of June 30, 2026, and December 31, 2025."*

### Cash Flow ($ thousands)

| Line | 6mo 2025 | 6mo 2026 |
|---|--:|--:|
| Net cash from operating activities | (323,793) | 300,042 |
| Net cash used in investing activities | (21,428) | (100,492) |
| Net cash from (used in) financing activities | (1,929) | 8,283 |
| End of period cash + restricted cash | — (not separately re-derived here) | 2,688,508 |

CFO Simon Edwards on the call: *"Cash flow from operations was $226 million... Free cash flow was $175 million, and we ended the quarter with $2.7 billion of cash."* The $226M figure matches the derived discrete-Q2 operating cash flow above almost exactly (derived: $226,432K). "$2.7 billion of cash" matches the balance sheet's $2,666,859K cash + $21,649K restricted cash ≈ $2.69B. Free cash flow ($175M) is a non-GAAP measure (presumably OCF minus capex) not separately re-derived in this pull — capex line item for Q2 alone was not isolated from the 10-Q's investing-activities section in this research pass.

### Share Count / Dilution

- Weighted-avg basic shares (Q2 2026): 287,288K; diluted: 323,331K
- Shares outstanding June 30, 2026: 293,354,001 (all Class A; zero Class B; zero preferred; 20,000,000 preferred shares authorized but none issued/outstanding)

### Customer / Credit Concentration (Note 1) — flagged internal inconsistency, see note below

> **Customer Risk** — "During the three months ended June 30, 2026, revenue from two customers, the second of which is our related party ..., accounted for approximately 44% and 21% of our total revenue. During the six months ended June 30, 2026, revenue from one customer, which is not our related party, accounted for approximately 73% of our total revenue."
>
> "During the three months ended June 30, 2025, four customers, none of which are related parties, represented approximately 30%, 18%, 15%, and 11% of our total revenue. During the six months ended June 30, 2025, two customers, neither of which are related parties, represented approximately 33% and 23% of our total revenue."

**⚑ Apparent internal inconsistency, presented as found rather than resolved:** the same 10-Q's revenue footnote (Note 3, footnote 1) states *"related party revenue of $2.8 million and $376.1 million for the three and six months ended June 30, 2026, respectively"*. If the "related party" customer named in the concentration note is 21% of Q2 revenue (≈$223.7M on $1,065.4M total), that is very hard to reconcile with a $2.8M related-party revenue figure for the same three months. (The six-month figure does reconcile cleanly against the Q1 notes' own finding: Q1 2026 related-party revenue was reported as $373.3M there, and $373.3M + $2.8M = $376.1M — exactly the six-month figure disclosed here — so the $2.8M / $376.1M pair is internally self-consistent and consistent with the Q1 report. The tension is specifically between "$2.8M related-party revenue in Q2" and "a related-party customer at 21% of Q2 revenue" appearing in the same note.) Two customers or classifications may be getting conflated here (e.g., the concentration note's "related party" designation may reflect a *contractual counterparty* relationship — see the note's own asterisked definition below — that is broader than the strict related-party-transaction revenue recognized under the Bolt JV equity-method relationship). This is flagged rather than resolved because untangling it would require data not present in the excerpts reviewed; whoever drafts the analysis should treat both figures as filed-but-apparently-in-tension rather than pick one.

> *Definition of "customer."* For purposes of the concentration of risk disclosure, "customer" refers to the contractual counterparty to which we sell our products and fulfil installation obligations, which in certain transactions may be a project-finance affiliate rather than the ultimate end user of the products. See Note 7 for the Brookfield-affiliated financing framework structure.

**Credit Risk (AR):** As of June 30, 2026, three customers — the third being the related party — accounted for approximately 36%, 34%, and 17% of accounts receivable (vs. 41%/17%/15% at Dec 31, 2025, first being the related party). *"To date, we have not experienced any material credit losses from these customers."*

**Geographic:** U.S. revenue was 90% of total for both the three and six months ended June 30, 2026 (vs. 59%/58% for the same 2025 periods) — an even higher concentration than Q1's already-elevated 91%. No dollar-level segment breakout of "data center" vs. "C&I" revenue exists anywhere in the 10-Q, unchanged from Q1.

**Backlog / remaining performance obligation (RPO):** **Not disclosed.** The 10-Q explicitly elects the ASC 606 practical expedient: *"We do not disclose the value of the unsatisfied performance obligations for (i) contracts with an original expected length of one year or less and (ii) contracts for which we recognize revenue at the amount to which we have the right to invoice for services performed."* No dollar or GW backlog figure appears anywhere in the 10-Q. Searched exhaustively for "backlog," "remaining performance obligation," and "RPO" — zero quantified figures found in the filing.

### Notable MD&A / Note Items Not Emphasized on the Call

- **Specific warranty reserve for "identified product issues" nearly tripled quarter-over-quarter.** Q1 FY2026 notes flagged a new $19.7 million specific warranty reserve. As of June 30, 2026, the 10-Q discloses: *"Includes a specific warranty reserve of $58.3 million, which is accounted for as an assurance-type warranty and recognized within cost of product revenue."* That implies roughly a **$38.6 million sequential increase** in this specific reserve during Q2 alone (58.3 − 19.7), though the warranty rollforward table in the 10-Q is presented on a 6-month YTD basis only (opening balance $20.0M → additions of $71.6M → expenditures of $(13.8)M → closing $77.8M total product-warranty-plus-performance balance), so a clean Q2-only isolation of the specific-reserve component specifically (as opposed to the whole warranty/performance liability) is not separately broken out. Not mentioned on the call in the portions reviewed. No detail is given anywhere in the 10-Q on what the "identified product issues" actually are, mirroring Q1.
- **New risk factor added in Q2 — not boilerplate "no material changes."** Unlike Q1's blanket "There were no material changes in risk factors as disclosed in our 2025 Form 10-K," the Q2 10-Q reads: *"There were no material changes in risk factors as disclosed in our 2025 Form 10-K, except as set forth below,"* followed by an entirely new risk factor on short-seller activity: *"Short sellers have published reports containing allegations regarding us and our business, and we have been, and may in the future be, the subject of such activities... we treat information regarding our suppliers and sourcing arrangements as confidential and proprietary."* This is a materially new disclosure item versus Q1 and versus the FY2025 10-K baseline, and was not mentioned on the earnings call in the portions reviewed.
- **CEO adopted his first active 10b5-1 sale plan, disclosed via 10-Q Item 5 (not yet a Form 4 as of this data pull).** *"On May 27, 2026, Dr. KR Sridhar, our Chief Executive Officer and Chairman of the Board, adopted a Rule 10b5-1 trading arrangement with an expiration date of September 1, 2027... for the sale of up to 200,000 shares of common stock."* The Q1 FY2026 notes explicitly found the CEO sold zero shares on the open market through the Q1 census window — this is a new and notable governance/insider signal. See Section 2 for why no corresponding Form 4 exists yet.
- Also via Item 5: Jeffrey Immelt (director) adopted a new 10b5-1 plan on 5/1/2026 for up to 60,000 shares; Shawn Soderberg (CLO) modified her existing plan on 5/22/2026 to add option-exercise-and-sale and additional RSU-vesting-sale provisions.
- Performance guarantee payments: $5.4M (Q2 2026) vs. $3.0M (Q2 2025) for the quarter; $13.8M vs. $14.6M for the six months — roughly flat YoY on a six-month basis.
- Oracle warrant mechanics fully played out this quarter: the warrant (3,531,073 shares, $113.28 strike) was issued April 9, 2026 at an estimated fair value of $251.6 million (Note 3), then Oracle executed a cashless exercise on May 1, 2026, resulting in 1,905,433 shares issued plus an additional 248,798 inducement shares (aggregate fair value of shares issued: $324.4 million). $17.9 million has been recognized cumulatively as a reduction of revenue related to the warrant as of June 30, 2026 (of which $5.0M + $1.9M = $6.9M was recognized in the three and six months ended June 30, 2026 combined, per the note's own breakout — precise figures: $5.0 million (three months) and $1.9 million... [the note's wording is ambiguous as to whether these two figures are three-month vs six-month or two components of the same period; presented as written in the filing rather than resolved]).

---

## 2. Financing Partnership Disclosures — Brookfield / IDF corroboration check (per framework's explicit ask)

The call's headline financing claims (K.R. Sridhar, prepared remarks): *"We formed the partnership last fall at $5 billion. Nine months later, in June, Brookfield expanded its commitment fivefold to $25 billion,"* and: *"Industrial Development Funding, who has previously funded Bloom deployments, partnered with Oaktree, MUFG Bank, and Morgan Stanley to fund Bloom deployments, cumulatively bringing their total commitment to $2.6 billion."*

**Baseline confirmed in the FY2025 10-K** (filed 2026-02-09, before the June expansion described on the call): *"a $5.0 billion financing framework with Brookfield Asset Management."* This $5.0B figure is real and filed.

**The Q2 FY2026 10-Q does not corroborate the expansion figures.** Searched exhaustively across the full 10-Q text (all ~4,800 lines of the converted document) for "$20 billion," "$25 billion," "fivefold," "Nebius," "IDF," "Industrial Development Funding," "Oaktree," and "MUFG" — **none of these strings appear anywhere in the 10-Q.** "Morgan Stanley" appears only as an analyst-firm affiliation in the (separately-sourced) call transcript, not in the 10-Q. The 10-Q's Item 2 MD&A financing-partners discussion is a single cross-reference sentence pointing back to the FY2025 10-K's "Financing Partners" section, with no updated dollar figures.

**What Note 7 (Investments in Unconsolidated Affiliates) actually discloses instead** is Bloom's own minority equity stakes in the underlying Fund JVs: Bolt US Class A JVCo LLC (9.9%), Bolt US JVCo LLC (9.9%), and a newly-added Other JV, ORC HoldCo LLC (15.0%). Bloom's own dollar exposure is small: *"Our maximum exposure to loss from the involvement with the Fund JVs as of June 30, 2026 is $68.8 million"* and *"Our total capital commitment to the Fund JVs as of June 30, 2026 is $77.3 million."*

**This is not necessarily a contradiction** — the JVs' total financing capacity (funded by Brookfield/IDF/Oaktree/MUFG/Morgan Stanley as external capital providers) is structurally distinct from Bloom's own equity-method investment in those JVs, which stays on Bloom's balance sheet only at Bloom's much smaller ownership-percentage stake; a JV can be capitalized with billions of third-party project-finance dollars while Bloom's own consolidated exposure remains in the tens of millions. But **the $25 billion and $2.6 billion figures themselves are not independently verifiable against any number filed by Bloom in this 10-Q** — they exist, in the materials reviewed for this research, only in the earnings call transcript. Treat them as call-only claims, per the same caveat the Q1 notes applied to the "2.45 GW Project Jupiter" figures.

**Nebius named on the call, not in the filing.** The CFO's prepared remarks name Nebius explicitly as the end-customer behind the IDF financing arrangement (*"Nebius signed the offtake, and IDF... is purchasing the energy service on cash terms"*) and the CEO's remarks separately claim *"Nebius did [cancel combustion turbine orders and choose Bloom] this quarter."* "Nebius" does not appear anywhere in the 10-Q. Like Oracle in the Q1 report (where Oracle was at least named in a footnote, just not with the GW figures), Nebius does not appear at all — not even in a footnote — in this quarter's 10-Q.

---

## 3. Supply Chain & Manufacturing Disclosure

**No new 10-K has been filed since the Q1 FY2026 cross-reference notes.** The FY2025 10-K (accession `0001628280-26-006516`, filed 2026-02-09) remains the most recent annual report, and its manufacturing-footprint and supplier disclosures are unchanged from what the Q1 notes already documented in detail (Fremont/Newark manufacturing footprint, Korea JV, no individually-named suppliers, no supplier-concentration percentages, and the same internally-inconsistent China-exposure language between Item 1 Business and Item 1A Risk Factors). Re-summarizing in full here would duplicate the Q1 file without new information; see `research/be-analysis-2026q1/BE Q1 FY2026 SEC Cross-Reference Notes.md`, Section 3, for the full writeup. The Q2 10-Q itself contains no incremental supply-chain disclosure beyond the general risk-factor cross-reference to the 2025 10-K.

The one notable supply-chain-adjacent addition this quarter is qualitative and appears only on the call, not in any filing: K.R. Sridhar's claim that Bloom has *"multiple qualified suppliers across multiple countries for every critical input... No single supplier and no single country determines our destiny."* This is consistent in direction with, but goes further than, the 10-K's more hedged and internally-inconsistent supplier language — and, like the financing figures above, is not independently corroborated by any filed document reviewed in this research.

---

## 4. Insider Transactions (Form 4) — 2026-06-01 to 2026-08-05 window, as filed through 2026-07-29 (data-pull date)

**Coverage gap, stated explicitly:** this research was performed 2026-07-29, one day after the July 28, 2026 earnings call. The most recent Form 4 on file for Bloom Energy as of the pull is dated **2026-07-02** (filed) / transaction date 2026-07-01. **No Form 4 has yet been filed covering any transaction in the roughly four weeks around the earnings call itself (mid-July through the print).** Section 16 filers have up to two business days after a transaction to file, so trades from approximately July 24 onward would not necessarily be visible yet regardless of whether they occurred — this is a timing limitation of the data pull, not a finding that no such trades exist. If this research is revisited even a few days later, re-checking `https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001664703&type=4` for filings dated after 2026-07-02 is recommended, especially given the CEO's new 10b5-1 plan noted below.

**8 filings found in window**, covering 7 distinct reporting events:

| Date | Insider | Title | Transaction | Shares | Price | Owned after |
|---|---|---|---|--:|--:|--:|
| 2026-05-28 (filed 2026-06-03) | Chambers John T | Director | S — 10b5-1 sale, plan adopted 2/26/2026, via JC2 Investments LLC | 55,000 | $297.69 (wtd avg) | 238,333 |
| 2026-06-16 | Soderberg Shawn Marie | CLO ("See Remarks") | S — tax withholding on RSU vesting, via Shawn M. Soderberg 2005 Trust | 2,842 | $288.63 (wtd avg) | 132,265 |
| 2026-06-16 | Chitoori Satish | COO | S — tax withholding, 10b5-1 plan adopted 11/28/2025 | 2,837 | $289.11 (wtd avg) | 207,417 |
| 2026-06-16 | Kurzymski Maciej | CAO | S — tax withholding on RSU vesting | 2,259 | $288.62 (wtd avg) | 79,686 |
| 2026-06-16 | Joshi Aman | CCO | S — tax withholding, 10b5-1 plan adopted 11/26/2025 | 3,558 | $289.14 (wtd avg) | 172,150 |
| 2026-06-30 | Immelt Jeffrey R | Director | A — DSU grant, 2021 Deferred Compensation Plan | 85 | $302.70 | 231,243 |
| 2026-06-30 | Warner Cynthia J | Director | A — DSU grant, 2021 Deferred Compensation Plan | 76 | $302.70 | 34,895 |
| 2026-07-01 | Joshi Aman | CCO | S — 10b5-1 sale, plan adopted 11/26/2025 | 8,343 | $300.37 (wtd avg) | 163,807 |

**Signal read:** every disposition in this window is either (a) a pre-existing, pre-print-adopted 10b5-1 mechanical sale, or (b) a routine tax-withholding sale tied to RSU vesting — no discretionary open-market selling (Form 4 transaction code "P" for purchases does not appear at all; all dispositions are code "S," and all are footnoted as either 10b5-1-plan-driven or tax-withholding-driven). This continues the Q1 report's pattern exactly: no CEO or CFO Form 4 activity at all in this window (Sridhar and Edwards have zero Form 4 filings between 2026-06-01 and the 2026-07-02 cutoff of available data), and all named-executive activity among the other officers is mechanical.

**The one genuinely new and notable data point is not a Form 4 at all**, but the Item 5 disclosure inside the Q2 10-Q itself (see Section 1 above): **KR Sridhar adopted his first-ever active 10b5-1 sale plan on May 27, 2026**, for up to 200,000 shares, expiring September 1, 2027. As of this data pull, no Form 4 has yet recorded any sale under this plan — worth monitoring going forward, since the Q1 report's entire insider-signal thesis rested partly on the CEO having sold zero shares to date.

Chambers' 5/28/2026 sale (55,000 shares, $297.69) is the same transaction already reported in the Q1 FY2026 notes' census (which covered through 2026-07-23) — flagged here rather than silently duplicated, since this quarter's window (6/1–8/5) necessarily overlaps the prior quarter's window by construction.

---

## 5. Discrepancies Between the Earnings Call and the SEC Filings (ranked by materiality)

**(1) Brookfield's "$5B → $25B" expansion and the "$2.6B" IDF/Oaktree/MUFG/Morgan Stanley financing commitment are call-only claims, not independently corroborated by any dollar figure in the Q2 10-Q.** See Section 2 above for the full corroboration check. The original $5.0B Brookfield framework is confirmed in the FY2025 10-K; the stated expansion to $25B and the new $2.6B figure appear nowhere in the Q2 10-Q. Bloom's own consolidated balance-sheet exposure to the underlying JVs is two orders of magnitude smaller ($68.8M maximum loss exposure, $77.3M total capital commitment) — not necessarily contradictory given the equity-method accounting structure, but not a reconciliation either.

**(2) Nebius is named as a specific customer on the call and does not appear anywhere in the 10-Q**, mirroring the Q1 report's finding on Oracle/"Project Jupiter" (though Oracle at least appeared in a warrant footnote last quarter; Nebius appears in zero SEC filings reviewed this quarter).

**(3) A revenue-concentration internal inconsistency exists within the 10-Q itself** (Section 1 above): the customer-concentration note implies a related-party customer at ~21% of Q2 revenue (~$223.7M), while the revenue footnote states related-party revenue was only $2.8 million for the same three months. Not a call-vs-filing discrepancy, but a filing-vs-filing one, worth flagging to whoever drafts the report since it affects how "customer concentration" gets characterized.

**(4) The specific warranty reserve for unspecified "identified product issues" nearly tripled quarter-over-quarter ($19.7M → $58.3M) and was not mentioned on the call** in the portions reviewed, continuing the exact pattern the Q1 report flagged.

**(5) A materially new risk factor (short-seller activity) was added this quarter and was not mentioned on the call.** Notably, the new risk-factor language itself states the company treats "information regarding our suppliers and sourcing arrangements as confidential and proprietary" — a limitation on how much it can rebut short-seller claims specifically about supply chain, which is worth keeping in mind given this framework's own supply-chain diligence angle.

**(6) GAAP vs. non-GAAP gap direction reversed in character from Q1.** In Q1, the GAAP/non-GAAP EPS gap was +91% (mostly optical, since GAAP net income itself was small and freshly positive). In Q2, GAAP net income is now large in absolute terms ($196.3M) and the same-size SBC add-back narrows the gap to +26%. This is a mechanical function of the operating-leverage story the CFO described on the call ("revenue grew 166%, operating expenses grew just 48%") and is not itself a red flag — noted here only because a reader comparing the two quarters' "gap %" without normalizing for this would draw the wrong conclusion (shrinking gap ≠ shrinking SBC; SBC in dollar terms was essentially flat, $57.0M → $56.4M).

**Cleanly corroborated (no discrepancy):** revenue $1,065.4M / +165.6% YoY / +42% QoQ (10-Q figures match the call's "up 166% year-over-year and 42% sequentially" essentially exactly, small rounding aside); operating cash flow ~$226M (derived $226,432K vs. call's "$226 million"); "$2.7 billion of cash" (10-Q shows $2,666,859K cash + $21,649K restricted = ~$2.69B); gross margin 34.3% non-GAAP (matches 8-K exactly); product gross margin trajectory (37.2%, up sequentially and YoY per the call, directionally consistent with the specific-warranty-reserve headwind being more than offset by mix/volume, though the exact ex-reserve product margin was not independently re-derived this quarter given the 6-month-only warranty rollforward presentation); no backlog dollar/GW figure given on the call either, consistent with the 10-Q's explicit non-disclosure election.

---

僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
