# Nebius Group N.V. (NBIS) Q1 2026 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK 0001513845), User-Agent identified per SEC fair-access policy. Company: Nebius Group N.V. (Dutch N.V., Schiphol, Netherlands), formerly Yandex N.V. (renamed 2024-08-16). Foreign private issuer — files Form 20-F (annual) and Form 6-K (interim/current reports), not 10-K/10-Q.

- **20-F** filed 2026-04-30, FY2025 (period ended 2025-12-31) — accession 0001104659-26-052948 (`nbis-20251231x20f.htm`). A **20-F/A** amendment was also filed 2026-05-22 (accession 0001104659-26-065681, 372KB) — not reviewed in this pass; likely a Part III/governance-type supplement, common shortly after an original 20-F.
- **6-K (earnings release cover)** filed 2026-05-13 (same day as the call) — accession 0001104659-26-059872 (`tm2614392d1_6k.htm`). Cover-shell only; actual content is Exhibit 99.1 (press release, `tm2614392d1_ex99-1.htm`) and Exhibit 99.2 (Letter to Shareholders, not fetched).
- **6-K (formal interim financials)** filed 2026-05-20 — accession 0001104659-26-064092 (`nbis-20260331x6k.htm`). Also a cover-shell; content is Exhibit 99.1 (`nbis-20260331xex99d1.htm`, "Operating and Financial Review and Prospects" / MD&A) and Exhibit 99.2 (`nbis-20260331xex99d2.htm`, full Unaudited Condensed Consolidated Financial Statements). **This is the authoritative source for every number below** — it postdates and supersedes the rounded figures spoken on the call and in the 05-13 press release.
- **Form 4 / Form 3**: Confirmed these DO exist on EDGAR (NBIS is not fully exempt from Section 16 filing in practice — a batch of Form 3s was filed 2026-03-18, presumably newly-appointed/newly-Section-16 officers or directors, followed by regular Form 4 activity). Filings bracketing the 2026-05-13 call: two Form 4s filed 2026-05-15 (accessions 0001513845-26-000061 and -000057, both reporting transaction date 2026-05-13 — the call date itself) plus a Form 4/A filed the same day (accession -000062, amending a transaction dated 2026-05-13), and another Form 4 filed 2026-05-19 (transaction date 2026-05-15). **Limitation: the XML detail of these specific filings (filer identity, transaction code, share count, price) could not be fetched in this research pass** — EDGAR access became intermittently unavailable partway through this session. Do not present specific insider names/amounts for the May 2026 window without verifying directly; report the existence and timing of the filings only, and flag this as an open follow-up rather than fabricating detail.

Call date: **2026-05-13, 8:00 AM EDT**. All figures below are the **Q1 2026 6-K financial statements (Exhibit 99.2) as filed 2026-05-20**, which are more precise than the rounded numbers spoken on the call.

---

## Income Statement (Three months ended Mar 31; $M)

| Line | Q1 2025* | Q1 2026 | Δ |
|---|--:|--:|--:|
| Revenue | 50.9 | 399.0 | +684% |
| Cost of revenues | 24.7 | 103.8 | +320% |
| Product development | 36.5 | 67.4 | +85% |
| Sales, general & administrative | 60.9 | 143.8 | +136% |
| Depreciation & amortization | 49.1 | 212.0 | +332% |
| Total operating costs & expenses | 171.2 | 527.0 | +208% |
| **Loss from operations** | **(120.3)** | **(128.0)** | worse in $ terms |
| Interest income | 8.5 | 14.2 | — |
| Interest expense | 0.0 | (63.7) | new (convert notes) |
| Gain from revaluation of investment in equity securities | 0.0 | **780.6** | non-cash |
| Income/(loss) from equity method investments | 0.1 | (7.6) | — |
| Other income, net | 8.3 | 19.9 | — |
| Income/(loss) before tax | (103.4) | 615.4 | — |
| Income tax expense/(benefit) | 0.9 | (5.8) | tax benefit |
| Net income/(loss) from continuing ops | (104.3) | 621.2 | — |
| Net loss from discontinued ops (Toloka) | (9.2) | 0.0 | Toloka deconsolidated Q2 2025 |
| **Net income/(loss)** | **(113.5)** | **621.2** | matches call's "$621M" exactly |
| Diluted EPS | (0.48) | 2.11 | — |
| Diluted shares | 237,916,047 | 308,971,701 | +30% share count |

*Q1 2025 restated for discontinued-operations presentation of Toloka.

**⚑ Quality-of-earnings flag — this is the single most important number in the filing and no analyst asked about it on the call.** GAAP **loss from operations actually widened** in dollar terms, from -$120.3M (Q1 2025) to -$128.0M (Q1 2026), despite revenue growing 684%. The touted "$621M net income" is **more than 100% explained by a single non-cash line**: a $780.6M "gain from revaluation of investment in equity securities," which the MD&A explicitly states is "primarily the remeasurement of our investment in ClickHouse, following a third-party investment in that company" (ClickHouse's January 2026 Series D at a ~$15B valuation — see Supply Chain notes). Strip that one non-cash mark-up and the company is still solidly loss-making at the operating level; the net-income headline and the "Group-adjusted EBITDA $130M / 13% margin, Nebius AI-adjusted EBITDA $45M/45% margin" figures cited on the call are **non-GAAP adjusted metrics that exclude the same $212.0M of D&A that is the main reason the GAAP operating loss doesn't close** (D&A alone is larger than the entire adjusted-EBITDA figure). Both framings are legitimate and disclosed, but the call emphasized only the adjusted/non-cash-gain side. See Analysis → Risk Matrix and Financials → Income Statement tab.

**⚑ Depreciation risk overlay:** the $212.0M D&A line (4.3x YoY) is the account directly named in the 20-F's disclosed **material weakness** over "controls related to fixed assets... depreciation start dates" (see below) — i.e., the one figure most responsible for the GAAP-vs-adjusted gap is also the one figure management has formally told the SEC it cannot yet reliably control.

## Segment / business-line detail (Item: MD&A "Operating Segments", Exhibit 99.1)

| Segment | Q1 2025 ($M) | Q1 2026 ($M) | YoY | Q1 2026 % of total |
|---|--:|--:|--:|--:|
| Nebius AI cloud | 41.4 | 389.7 | +841% | 97.6% |
| Avride | 0.2 | 0.9 | +350% | 0.2% |
| TripleTen | 10.5 | 11.6 | +10% | 2.9% |
| Total segment revenues | 52.1 | 402.2 | +672% | — |
| Eliminations (intercompany) | (1.2) | (3.2) | — | — |
| **Total revenues** | **50.9** | **399.0** | **+684%** | — |

MD&A verbatim: "The Avride business had only a limited contribution to the total revenue for the Group." TripleTen's own growth (+10% YoY, "driven primarily by approximately 5,000 new student enrollments") is dramatically slower than the core business — a year ago TripleTen was ~20.6% of group revenue; one year later it is ~2.9%, pure dilution rather than deceleration in absolute terms. Nebius AI cloud's Annualized Run-Rate Revenue reached **$1.9B at end of March 2026** (up from $1.25B end of Q4 2025), per the call — not independently re-stated as a discrete line in the 6-K exhibits reviewed.

## Customer concentration & geography (source: FY2025 20-F, Notes to Financial Statements — "Significant Customers" / "Supplier Concentration"; NOT re-disclosed in the Q1 6-K, so these are FY-end 2025 figures, three-and-a-half months stale relative to the call but the only quantified concentration data available)

- **Revenue concentration, customers ≥10% of revenue, by year:**

| Customer | FY2023 | FY2024 | FY2025 |
|---|--:|--:|--:|
| Customer A | <10% | 27% | 25% |
| Customer B | <10% | <10% | 15% |
| Customer C | <10% | 11% | <10% |
| Customer D | <10% | <10% | <10% |

- **Accounts-receivable concentration (credit-risk exposure, year-end gross AR):** $6.6M (59% of AR) attributable to Customer A at 2024-12-31; **$597.0M (83% of AR) attributable to Customer D at 2025-12-31**; no significant concentration at 2023-12-31.
- **Supplier concentration:** "attributable to one supplier in 2024 and to two suppliers in 2025" (measured on gross carrying amount of capex incurred that year) — suppliers not named, no dollar/percent figure given (unlike the customer side).
- 20-F does not name Customer A/B/C/D. Given Customer D was **under the 10% revenue-recognition threshold in every year shown yet is 83% of the AR balance at 2025 year-end**, and the Q1 2026 balance sheet shows AR roughly doubling further (720.3→1,479.2) with deferred revenue tripling in the same quarter, the most parsimonious read is a very large, very recently-signed contract where billings/receivables ran far ahead of revenue recognition — consistent in timing with Nebius's **original "Cloud Infrastructure Services Agreement" with Meta dated November 1, 2025** (Exhibit 4.2 to the 20-F), which would sit right at 2025 year-end as a freshly-signed, not-yet-substantially-delivered contract. **This is inference, not a filing fact — the 20-F does not name Customer D as Meta.** Flag prominently as inferred, not confirmed.
- No revenue-side customer concentration figures for Q1 2026 itself were found in the reviewed exhibits (the 20-F's table is FY-annual only; the 6-K MD&A does not restate it quarterly).

## Balance sheet & capital structure ($M)

| Line | Dec 31, 2025 | Mar 31, 2026 | Δ |
|---|--:|--:|--:|
| Cash and cash equivalents | 3,678.1 | 9,298.2 | +5,620.1 |
| Accounts receivable, net | 720.3 | 1,479.2 | +758.9 |
| Total current assets | 4,711.4 | 11,238.3 | — |
| Property and equipment | 5,553.3 | 7,131.7 | +1,578.4 |
| Goodwill | 0.0 | 163.3 | new (Tavily) |
| Investments in non-marketable equity securities | 836.6 | 1,614.1 | +777.5 (ClickHouse markup) |
| **Total assets** | **12,430.6** | **22,303.3** | +79% in one quarter |
| Deferred revenue, current | 275.5 | 685.6 | +410.1 |
| Debt, non-current | 4,103.2 | 8,432.0 | +4,328.8 (≈ new convert notes) |
| Deferred revenue, non-current | 1,302.0 | 4,092.5 | **+2,790.5** |
| Total liabilities | 7,836.6 | 15,061.4 | — |
| Additional paid-in capital | 2,360.9 | 4,386.2 | +2,025.3 (≈ NVIDIA warrant + SBC) |
| Retained earnings | 3,300.5 | 3,921.7 | +621.2 (= exactly net income) |
| **Total shareholders' equity** | **4,594.0** | **7,241.9** | — |

- **Total deferred revenue grew ~$3.2B in one quarter** (410.1 current + 2,790.5 non-current) — this is the balance-sheet face of "prepayments reached a new quarterly record" from the call, now quantified.
- **Convertible senior notes** (Subsequent Events note, 20-F): priced 2026-03-18, closed 2026-03-20 — $2.25B of 1.250% notes due 2031 (over-allotment of $337.5M exercised, upsizing this tranche to ~$2.5875B) + $1.75B of 2.625% notes due 2033 (its $262.5M over-allotment was NOT exercised) = **~$4.34B total**. Call rounded the 2033 coupon to "2.60%"; filed rate is 2.625%. A separate, earlier "2029 Notes" tranche exists from a June 2025 6-K (Exhibit 99.4, not itself reviewed here) — this is NBIS's second convertible-notes issuance, not its first.
- **NVIDIA investment structure**: NOT a plain equity purchase — a **pre-funded warrant** (agreement dated 2026-03-11) for 21,065,396 Class A shares at a $0.0001 strike, for ~$2.0B gross proceeds; exercisable by NVIDIA any time on/after 2026-03-11. Cash-flow statement shows this as "Proceeds from issuance of pre-funded warrants sale — $2,000.0M" under financing activities, separate from the convertible notes line.
- Share count: 220,406,311 Class A + 33,491,883 Class B outstanding as of 2026-03-31. Class A = 1 vote/share, Class B = 10 votes/share (super-voting structure; 20-F risk factors flag "concentration of voting power with our founding shareholder" limiting minority shareholder influence — founder % not extracted in this pass).

## Cash flow (Three months ended Mar 31; $M)

| Line | Q1 2025 | Q1 2026 |
|---|--:|--:|
| Net income/(loss) from continuing ops | (104.3) | 621.2 |
| + Depreciation of PP&E | 48.6 | 208.8 |
| − Gain from revaluation of equity securities | 0.0 | (780.6) |
| + Change in deferred revenue | 2.4 | +3,198.0 |
| − Change in accounts receivable | (9.5) | (758.9) |
| **Net cash from operating activities** | **(197.5)** | **2,258.0** |
| Purchases of PP&E and intangibles (CapEx) | (543.9) | (2,472.9) |
| Acquisitions of businesses, net of cash | 0.0 | (170.2) (Tavily) |
| **Net cash used in investing activities** | (544.0) | (2,643.1) |
| Proceeds from convertible notes, net of costs | 0.0 | +4,293.7 |
| Proceeds from NVIDIA pre-funded warrant | 0.0 | +2,000.0 |
| **Net cash from financing activities** | (181.5) | 6,295.5 |
| Net change in cash (+ restricted) | (922.7) | +5,905.3 |
| Cash + restricted, end of period | 1,527.6 | 9,626.9 |

**⚑ Free cash flow was still negative this quarter — a number the call never stated directly.** Operating cash flow of $2,258.0M ($2.3B, as cited on the call) minus CapEx of $2,472.9M = **FCF of approximately −$214.9M**. The call's framing ("operating cash flow of $2.3 billion... driven by upfront payments from our customers") is accurate but one-sided: it does not mention that CapEx consumed slightly more than that operating inflow in the same quarter. The entire quarter's ~$5.9B cash build came from financing (convertible notes + NVIDIA warrant), not from the business self-funding its own buildout — consistent with management's own framing that the incremental $20–25B 2026 CapEx guide "will be funded through additional financing," but worth stating in FCF terms since the call leaned on the OCF figure alone.

## Manufacturing / GPU supply chain

See the companion file `NBIS 20-F Compute Supply Chain Notes.md` for full detail and sourcing. Headline: 20-F risk factors state plainly, "We currently rely on Nvidia for the GPU chips we use and on a limited number of other suppliers for other key components in our infrastructure."

## Form 4 / insider signal window

**Not completed to standard in this pass.** Confirmed multiple Form 3/Form 4/Form 4A filings exist on EDGAR bracketing the 2026-05-13 call (see filing list above), which is itself informative — despite FPI/Section-16 exemption generally being available to Dutch issuers, Nebius's insiders are filing anyway (voluntarily, or the company has opted in — the specific legal basis was not verified). But the actual filer identities, transaction codes (open-market sale vs RSU-vesting withholding vs option exercise), share counts and prices could not be retrieved before EDGAR access became intermittently unavailable during this research session. **Do not state or imply any specific insider bought or sold shares, or in what size, without fetching and reading the actual Form 4 XML first.** This is flagged as an open item for the report to either skip (state "Form 4 filings exist; detail not verified this pass") or complete in a follow-up fetch.

## Stock price reference points

Not independently verified in this pass (would need a live quote source). Do not fabricate a share price or market cap for the valuation tab without a verified source — if none is available, state implied multiples only where they don't require a share price (e.g., ARR growth, backlog-to-revenue ratios), or clearly mark any price-dependent figure as unverified/omit it.

---
僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
