# Applied Optoelectronics, Inc. (NASDAQ: AAOI) — 6-Quarter Financial Trend

**CIK:** 0001158114 (verified via SEC EDGAR full-text company search; confirmed name "APPLIED OPTOELECTRONICS, INC.", Sugar Land, TX, SIC 3674 Semiconductors & Related Devices, exchange Nasdaq, ticker AAOI)

**Data source:** SEC EDGAR XBRL `companyconcept` API (`data.sec.gov/api/xbrl/companyconcept/CIK0001158114/us-gaap/{Concept}.json`), cross-checked against the primary financial statements in each filing's main 10-Q/10-K document. All dollar figures in USD thousands except per-share and margin figures. AAOI's fiscal year is calendar-year (FYE 12/31).

**Filings used (6 discrete calendar quarters, Q1 2025 – Q2 2026):**

| Quarter | Form | Period end | Filed | Accession No. |
|---|---|---|---|---|
| Q1 2025 | 10-Q | 2025-03-31 | 2025-05-08 | 0001437749-25-015518 |
| Q2 2025 | 10-Q | 2025-06-30 | 2025-08-07 | 0001437749-25-025450 |
| Q3 2025 | 10-Q | 2025-09-30 | 2025-11-06 | 0001437749-25-033627 |
| Q4 2025 | *(derived: FY2025 10-K minus 9-month 10-Q YTD)* | 2025-12-31 | 2026-02-26 (10-K) | 0001437749-26-005875 |
| Q1 2026 | 10-Q | 2026-03-31 | 2026-05-07 | 0001437749-26-015620 |
| Q2 2026 | 10-Q | 2026-06-30 | 2026-08-06 | 0001437749-26-026278 |

Note: **the Q2 2026 10-Q was already filed (2026-08-06, same day as the earnings call) as of this research date (2026-08-09)**, so this table uses actual audited/reviewed 10-Q figures for Q2 2026, not press-release estimates. Q4 2025 has no standalone 10-Q (AAOI does not file a Q4 10-Q); it is derived by subtracting the Q1–Q3 2025 nine-month YTD 10-Q figure from the FY2025 10-K annual figure. This derivation is arithmetically exact for balance-sheet/flow items but **not exact for diluted EPS**, which is not additive across quarters because weighted-average diluted share counts differ quarter to quarter — flagged below.

---

## Income Statement Trend

| Concept | Q1 2025 | Q2 2025 | Q3 2025 | Q4 2025 (derived) | Q1 2026 | Q2 2026 |
|---|---|---|---|---|---|---|
| **Revenue, net** (`RevenueFromContractWithCustomerExcludingAssessedTax`) | $99,859 | $102,952 | $118,630 | $134,274 | $151,144 | $191,922 |
| **Gross Profit** (`GrossProfit`) | $30,544 | $31,162 | $33,263 | $41,944 | $43,916 | $53,207 |
| Gross Margin % *(calculated)* | 30.58% | 30.27% | 28.04% | 31.24% | 29.06% | 27.72% |
| **Operating Income (Loss)** (`OperatingIncomeLoss`) | $(8,937) | $(15,976) | $(18,187) | $(11,501) | $(12,991) | $(24,727) |
| Operating Margin % *(calculated)* | -8.95% | -15.52% | -15.33% | -8.57% | -8.60% | -12.88% |
| **Net Income (Loss)** (`NetIncomeLoss`) | $(9,172) | $(9,098) | $(17,936) | $(2,022) | $(14,281) | $(22,781) |
| Net Margin % *(calculated)* | -9.18% | -8.84% | -15.12% | -1.51% | -9.45% | -11.87% |
| **EPS, Diluted** (`EarningsPerShareDiluted`) | $(0.18) | $(0.16) | $(0.28) | ~$0.00 (derived, unreliable — see note) | $(0.19) | $(0.28) |

**Notes:**
- ⚑ *Revenue concept*: AAOI's `us-gaap:Revenues` tag only has data through 2018; from 2019 onward the company tags revenue as `RevenueFromContractWithCustomerExcludingAssessedTax` (ASC 606 tag). All revenue figures above use that tag.
- ⚑ *Q4 2025 derivation*: FY2025 revenue $455,715K (10-K) minus 9-month 2025 YTD revenue $321,441K (Q3 2025 10-Q) = $134,274K. Same subtraction method applied to Gross Profit, Operating Income, Net Income. Internal consistency check: Q1+Q2+Q3 2025 revenue = $99,859+$102,952+$118,630 = $321,441K, exactly matching the 9-month YTD figure independently reported in the Q3 2025 10-Q — confirms no restatement occurred between quarters.
- ⚑ *EPS Q4 2025 is NOT independently reported anywhere* — AAOI's FY2025 10-K diluted EPS is $(0.64) and the 9-month (Q1–Q3) diluted EPS is also $(0.64) as separately reported, which arithmetically implies ~$0.00 for Q4, but this is an artifact of diluted EPS being computed off each period's own weighted-average share count (which changed materially across the year due to ATM equity issuance) rather than a true additive quantity. Treat the Q4 2025 EPS cell as unreliable; the Q4 2025 net loss dollar figure above ($(2,022)K) is the reliable number for that quarter.
- ⚑ *Net loss narrowed sharply in Q4 2025* relative to Q1–Q3 2025 (from ~$(18)M/quarter to ~$(2)M), then loss widened again in Q1–Q2 2026 as opex (R&D, G&A) scaled ahead of the capacity build-out described in the Q2 2026 10-Q MD&A (see `AAOI 10-Q Key Data.md`).
- Gross margin has been trending down over the last two quarters (29.06% → 27.72%) even as revenue nearly doubled YoY, consistent with the Q2 2026 10-Q's disclosure of a major mix shift toward Data Center revenue (56.1% of Q2 2026 revenue vs. 43.5% in Q2 2025) and elevated ramp costs.

---

## Balance Sheet Trend (period-end)

| Concept | Mar 31 2025 | Jun 30 2025 | Sep 30 2025 | Dec 31 2025 | Mar 31 2026 | Jun 30 2026 |
|---|---|---|---|---|---|---|
| **Cash and Cash Equivalents** (`CashAndCashEquivalentsAtCarryingValue`) | $51,144 | $64,699 | $136,961 | $206,140 | $439,705 | $499,737 |
| **Inventory, Net** (`InventoryNet`) | $102,313 | $138,867 | $170,214 | $183,105 | $206,246 | $278,791 |
| **Long-Term Debt, Noncurrent** (`LongTermDebtNoncurrent`) | $1,811 | $0 | $0 | $0 | $0 | $1,657 |

**Notes:**
- Source for each column: same-quarter 10-Q (or the FY2025 10-K for the 12/31/2025 column; that figure is also re-confirmed as a comparative in both the Q1 2026 and Q2 2026 10-Qs, all three agree exactly at $206,140K cash / $183,105K inventory / $0 LT debt).
- ⚑ `LongTermDebtNoncurrent` here is a narrow XBRL line (excludes the $129.1M convertible senior notes, which AAOI tags separately as `ConvertibleNotesPayableNoncurrent` / reported as "Convertible senior notes" on the face of the balance sheet — see `AAOI 10-Q Key Data.md` for the full balance sheet including that line).
- ⚑ **Cash roughly doubled quarter-over-quarter from Dec 31 2025 ($206.1M) to Mar 31 2026 ($439.7M) and again to Jun 30 2026 ($499.7M)**, driven by a large at-the-market (ATM) common stock offering — see CapEx/financing notes below and in `AAOI 10-Q Key Data.md`. Shares outstanding rose from 74,998K (12/31/2025) to 84,386K (6/30/2026) per the Q2 2026 10-Q balance sheet.

---

## Capital Expenditure Trend (Purchases of Property, Plant & Equipment, cash-flow-statement basis)

| Concept | Q1 2025 | Q2 2025 | Q3 2025 | Q4 2025 (derived) | Q1 2026 | Q2 2026 |
|---|---|---|---|---|---|---|
| **CapEx** (`PaymentsToAcquirePropertyPlantAndEquipment` through Q1 2026; AAOI switched the XBRL tag to `PaymentsToAcquireOtherPropertyPlantAndEquipment` starting with the Q2 2026 10-Q) | $28,389 | $25,479 | $50,245 | $75,035 | $58,225 | $276,907 |

**Notes:**
- ⚑ AAOI reports CapEx as a cumulative year-to-date cash-flow-statement line, not a discrete-quarter one; discrete-quarter figures above (Q2, Q3 2025; Q4 2025; Q2 2026) are derived by subtracting the prior YTD cumulative figure from the current YTD cumulative figure. Q1 figures each year equal their own YTD figure (first quarter of the fiscal year).
- ⚑ **XBRL tag change**: the Q2 2026 10-Q retagged the "Purchase of property, plant and equipment" cash-flow line from `us-gaap:PaymentsToAcquirePropertyPlantAndEquipment` to `us-gaap:PaymentsToAcquireOtherPropertyPlantAndEquipment`. This is confirmed as the same line item because the 6-month-2025 comparative figure under the new tag ($53,868K) exactly matches the 6-month-2025 figure previously reported under the old tag. Six-month-2026 CapEx under the new tag is $335,132K; Q1 2026 CapEx (old tag) was $58,225K, implying Q2 2026 discrete CapEx of $276,907K.
- ⚑ **This is a step-change, not noise**: per the Q2 2026 10-Q MD&A (Liquidity and Capital Resources), the $335.1M six-month 2026 CapEx breaks down geographically as $169.0M US, $62.8M Taiwan, $103.3M China, driven by "facility expansion and equipment purchases to support increased production capacity for the Company's internet data center and broadband product lines, including investments related to Quantum Bandwidth products and the continued expansion of manufacturing operations for 400G, 800G, and 1.6T transceiver products." Management guided that "2026 CapEx will continue to be elevated above past year levels" and that capacity investment will continue "through at least the end of 2027."
- This capex ramp was funded primarily by a **$1.03 billion net-proceeds ATM common stock offering** completed in H1 2026 (per the Q2 2026 10-Q cash flow statement, "Proceeds from common stock offering, net" = $1,028,207K for the six months ended June 30, 2026, vs. $195,770K in the same period of 2025) — see `AAOI 10-Q Key Data.md` for the full cash flow statement.

---

*This file is intermediate source-data for a research report — not investment advice.*
