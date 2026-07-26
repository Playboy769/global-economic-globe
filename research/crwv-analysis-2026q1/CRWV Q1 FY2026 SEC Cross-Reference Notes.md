# CoreWeave, Inc. (CRWV) Q1 FY2026 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK 0001769628), User-Agent identified per SEC policy.

- **10-Q** filed 2026-05-08, period ended 2026-03-31 — accession 0001769628-26-000222 (`crwv-20260331.htm`)
- **10-K** filed 2026-03-02, FY2025 (period ended 2025-12-31) — accession 0001769628-26-000104 (`crwv-20251231.htm`, not yet read in full — 10-Q Item 1A states no material changes from 10-K risk factors except as noted)
- **10-Q (Q1 FY2025 comparable)** filed 2025-05-15, accession 0001769628-25-000014 — used only via the comparative columns embedded in the Q1 FY2026 10-Q, not fetched separately
- **Form 4** insider filings: 96 filings in the 2026-04-01 to 2026-05-25 window alone (CoreWeave has an unusually large number of Section-16 filers with high-frequency 10b5-1 programs); a representative sample of 8 filings was read in full (see Insider Transactions below), not an exhaustive census

Call date: **2026-05-07, 5:00pm ET** (10-Q filed the next day, 2026-05-08). All P&L/balance-sheet/cash-flow figures below are the **10-Q as filed**, which supersede the rounded/non-GAAP numbers spoken on the call.

---

## Income Statement (Three months ended Mar 31; $M except per-share)

| Line | Q1 2025 | Q1 2026 | YoY |
|---|--:|--:|--:|
| Revenue | 982 | **2,078** | +111.6% |
| Cost of revenue | 262 | 716 | +173.3% |
| Technology and infrastructure | 561 | 1,273 | +126.9% |
| Sales and marketing | 11 | 69 | +527% |
| General and administrative | 175 | 164 | −6.3% |
| **Total operating expenses** | 1,009 | **2,222** | +120.2% |
| **Operating loss** | (27) | **(144)** | — |
| Gain (loss) on fair value adjustments | 27 | 0 | — |
| Interest expense, net | (264) | (536) | +103.0% |
| Other income (expense), net | (5) | 24 | — |
| Loss before income taxes | (269) | (656) | — |
| Provision for income taxes | 46 | 84 | — |
| **Net loss** | **(315)** | **(740)** | — |
| Diluted EPS | $(1.49) | $(1.40) | — |
| Weighted-avg diluted shares (M) | 249 | 527 | +112% |

**⚑ Transcript-discrepancy resolved:** CFO Nitin Agrawal stated on the call "Net loss for Q1 was $740 million compared to a net loss of $315 million in **Q1 of 2022**" — CoreWeave was not a comparable public reporting entity in Q1 2022 (IPO was March 2025). The 10-Q confirms the real comparable figure is **Q1 2025 net loss of $315M** — an exact match once "2022" is read as a transcription error for "2025." This is a clean, fully-resolved flag; no other figures are affected.

**Cost structure note:** Cost of revenue grew **173%** YoY (34% of revenue vs 27% a year ago) — faster than revenue's 112% growth, consistent with the gross-margin decline the CFO characterized as "timing-based, not economic" (see Analysis.html Q&A Tab · Tal Liani). Technology & infrastructure (+127%) also outgrew revenue. G&A **fell 6.3%** YoY even as the company scaled — a genuine efficiency point not mentioned on the call.

**Stock-based compensation** (10-Q Note): $153M total in Q1 2026 (Cost of revenue $8M / Tech&Infra $55M / S&M $13M / G&A $77M) vs $184M in Q1 2025 (Cost of revenue $3M / Tech&Infra $55M / S&M $3M / G&A $123M) — total SBC actually **declined** YoY despite headcount/revenue growth, driven almost entirely by a G&A SBC drop ($123M→$77M).

## Remaining Performance Obligations ("RPO") — the GAAP backlog figure

- As of March 31, 2026: **$98.8 billion** of unsatisfied RPO (10-Q Note, "Remaining Performance Obligations") — the call's spoken "$99.4 billion revenue backlog" is a **different, non-GAAP metric**; the two are close but not identical, and management never explained the ~$0.6B gap or which figure is the "correct" one to model from.
- Recognition timing: **36%** over the initial 24 months (through Mar 31, 2028), **39%** between months 25–48, remainder (25%) in months 49–84. This exactly reconciles with the call's "36% next 24 months, 75% in next 4 years" claim (36+39=75) — a rare case where the call's rounded framing checks out precisely against the filed figure.

## Significant Customers — the single most important finding of this research

10-Q **Note (Significant Customers)** and **Item 1A Risk Factors** both disclose customer concentration directly, in unusually specific terms for a company of this type:

| Customer | % of revenue, Q1 2025 | % of revenue, Q1 2026 |
|---|--:|--:|
| Customer A | 72% | 45% |
| Customer B | * (under 10%) | 20% |
| Customer C | * | * |
| Customer D | * | * |

*Customer did not represent 10%+ of revenue in that period. "The customer references of A through D may represent different customers than those reported in a previous period" (10-Q verbatim caveat — ordering is not guaranteed stable quarter to quarter).*

**Accounts receivable concentration:** Customer A 39%, B 17%, C 22% of AR at Mar 31, 2026 (Customer A 68%, D 11% at Dec 31, 2025) — note Customer C is material in AR (22%) despite never crossing the 10%-of-revenue disclosure threshold.

**Risk Factors (Item 1A) states verbatim:** *"We recognized an aggregate of approximately 65% of our revenue from our top two customers for the three months ended March 31, 2026... We recognized an aggregate of approximately 72% of our revenue from our top customer for the three months ended March 31, 2025."* This is the exact **65%** figure — Customer A (45%) + Customer B (20%) = 65%, confirming Customers A and B are the "top two."

**Named contracts, same paragraph:**
- **OpenAI**: master services agreement entered **May 2025**; order form entered **September 2025** under which OpenAI "has committed to pay us up to approximately **$6.5 billion through May 31, 2031**."
- **Meta**: order form entered **March 2026** (i.e., **within Q1**, not a subsequent event) under an existing MSA, initially committed to pay up to approximately **$21 billion** (inclusive of new capacity access through Dec 20, 2032, and exercise of an existing option through Apr 10, 2032).
- ⚑ **A second, inconsistent OpenAI figure appears elsewhere in the same 10-Q** (a different risk-factor paragraph, on customer credit risk): *"in March 2025, we entered into a master services agreement with OpenAI... pursuant to which OpenAI has committed to pay us up to approximately **$11.9 billion through October 2030**."* This is a different MSA date (March 2025 vs. May 2025), different dollar figure ($11.9B vs $6.5B), and different end date (Oct 2030 vs May 2031) than the Significant-Customers-note paragraph. Both appear verbatim in the filed 10-Q; this research could not determine whether these represent two separate, additive order forms, a drafting inconsistency between sections, or one superseding the other. **Flag as unresolved — do not silently pick one figure.**
- **"Other significant customers include Microsoft and Meta"** (10-Q verbatim, credit-risk risk factor) — the first explicit 10-Q confirmation that **Microsoft** is a named significant customer (consistent with Customer A's 72%→45% trajectory being CoreWeave's historically-dominant single customer relationship, though the 10-Q does not explicitly map "Customer A" to "Microsoft" by name).
- ⚑ **Anthropic, Cohere, Jane Street, and Hudson River Trading — all named on the earnings call as customers — appear nowhere in the 10-Q.** This is either because they are each individually immaterial (each under whatever threshold triggers named disclosure) or too new to be captured in the quarter's disclosure controls. The call's customer-diversification narrative (financial services, physical AI verticals) is not corroborated or contradicted by the 10-Q — it simply isn't addressed, which given the 65%-concentration finding directly above, is itself informative.

## Balance Sheet (Mar 31, 2026 vs Dec 31, 2025; $M)

| Line | Dec 31, 2025 | Mar 31, 2026 | Δ |
|---|--:|--:|--:|
| Cash and cash equivalents | 3,127 | 2,244 | −883 |
| Restricted cash, current | 819 | 777 | −42 |
| Marketable securities | 34 | 22 | −12 |
| Accounts receivable, net | 3,169 | 2,120 | −1,049 |
| Total current assets | 7,488 | 5,609 | −1,879 |
| Property and equipment, net | 30,557 | **36,424** | **+5,867** |
| Operating lease ROU assets | 8,231 | 10,182 | +1,951 |
| Total assets | 49,302 | **55,573** | **+6,271** |
| Debt, current | 6,708 | 7,547 | +839 |
| Debt, non-current | 14,665 | **17,312** | **+2,647** |
| Deferred revenue, non-current | 6,476 | 5,393 | −1,083 |
| Total liabilities | 45,967 | **50,814** | +4,847 |
| **Total stockholders' equity** | ~3,335 | **4,759** | +1,424 |

Total debt (current + non-current): **$24,859M** (Mar 31, 2026) vs $21,373M (Dec 31, 2025), +16.3% in one quarter.

**⚑ Accounts receivable fell $1,049M (−33%) quarter-over-quarter** even as revenue grew 32% QoQ — worth noting alongside the customer-concentration finding above, though the 10-Q does not explain the AR decline and it was not addressed on the call.

**Property and equipment breakdown** (10-Q Note): Technology equipment $26,627M (Dec 31: $20,903M), Data center equipment & leasehold improvements $3,878M ($2,842M), Software $827M, plus construction-in-progress (not separately itemized in the extract captured here).

## Debt — full stack (10-Q Note 10 and MD&A)

- **Delayed draw term loan facilities, aggregate outstanding: $11.8 billion** as of Mar 31, 2026 — collateralized by contributed contracts/pledged cash flows, "generally from investment grade counterparties."
- **DDTL 4.0 Facility** (the one described on the call): entered into in **March 2026** by subsidiary **CoreWeave Compute Acquisition Co. VIII, LLC ("CCAC VIII")**, lenders led by **MUFG Bank, Ltd.** as administrative agent.
  - **$8.5 billion** total: ~$4.5B floating-rate (SOFR + 2.25%, or ABR + 1.25%) + ~$4.0B fixed-rate (2.00% + blended UST rate at draw).
  - Available in draws through **June 30, 2027** (commitment termination date); **matures March 2032**.
  - **Non-recourse**, secured by CCAC VIII equity/assets only, "except for limited guarantees related to customary non-recourse carve-out obligations."
  - Undrawn fee 0.50% p.a.; **$142M** deferred financing costs capitalized.
  - Requires interest-rate hedges on ≥95% of floating-rate exposure, plus power-cost hedging; company confirms compliance as of Mar 31, 2026.
  - ⚑ **The 10-Q text describing DDTL 4.0 does not mention Moody's, Fitch, DBRS, or any "A-equivalent" investment-grade rating claim** — the credit-rating detail cited on the call appears to be an earnings-release/IR claim, not part of the 10-Q's own legal description of the facility. This is common (10-Qs often don't repeat rating-agency marketing language) but means the "A- equivalent, first-ever investment grade" framing is not independently corroborated in the filed document reviewed here.
- **Senior Notes / Convertible Notes: $6.4 billion aggregate outstanding** — $2.0B 2030 Senior Notes, $1.8B 2031 Senior Notes, $2.6B 2031 Convertible Senior Notes.
- **Revolving Credit Facility**: $2.5B total, $686M available (i.e., ~$1.8B drawn) as of Mar 31, 2026.
- OEM and software-license vendor financing arrangements also disclosed, not separately quantified in the extract reviewed.
- **DCSP Note Receivable** (Oct 2024, CoreWeave as *lender*): up to $305M delayed-draw term loan **funding provided by CoreWeave** to a data-center-services-provider counterparty, 7-year term, 13.00% p.a. — i.e., CoreWeave is simultaneously a large borrower and, in this one instance, itself an infrastructure lender.
- **Liquidity**: Total liquidity (cash + marketable securities + facility availability) = **$11,091M** at Mar 31, 2026 (up from $6,862M at Dec 31, 2025), of which $8,825M is undrawn facility capacity.

## NVIDIA — equity, supply, and collaboration (multiple 10-Q sections)

- **Equity purchase confirmed**: *"In January 2026, we entered into a securities purchase agreement with NVIDIA Corporation for a private placement of approximately 23 million shares of our Class A common stock at a purchase price of $87.20 per share, for aggregate gross proceeds of $2.0 billion."* — this **directly resolves** Analysis.html's "尚待驗證" flag on Special Event B: **NVIDIA is confirmed as the direct counterparty/purchaser**, not a coincidentally-timed public offering.
- **Collaboration framework**: *"in January 2026, we announced that we entered into a collaboration framework with NVIDIA Corporation to expand our long-standing complementary relationship to advance AI adoption at global scale"* — both the equity purchase and the collaboration framework are dated **January 2026, within Q1**.
- **Supply dependency (Item 1A)**: *"as a result of our obligations in our current customer contracts, all of the GPUs used in our infrastructure today are NVIDIA GPUs."* Also: *"our current customers have contractually specified our use of NVIDIA GPUs"* — an unusual framing where the sole-sourcing is customer-contract-driven, not (only) a CoreWeave procurement choice.
- **Supplier concentration trend** (Item 1A, total purchases by year): FY2023 top-3 suppliers 57%/22%/11%; FY2024 46%/16%/14%; FY2025 23%/20%/17%. Steady deconcentration of the top supplier (57%→46%→23%) mirrors the customer-side deconcentration trend (72%→45% for the top customer) — both sides of the business show the same "still concentrated but visibly diversifying" pattern.
- **Upstream chain, verbatim**: *"our suppliers themselves rely on a complex network of third-party suppliers... For example, NVIDIA relies on suppliers such as **Taiwan Semiconductor Manufacturing Company** for semiconductor fabrication and other manufacturers for compute and networking components. Any disruption... whether due to... geopolitical factors such as the growing potential for military conflict between China and Taiwan..."* — this is the 10-Q's own explicit TSMC/Taiwan-geopolitical-risk language, one supply-chain tier removed from CoreWeave itself.
- **5GW option**: no mention of "5 GW," "gigawatt," or any specific NVIDIA capacity-option figure found anywhere in the 10-Q text searched. This appears to be an earnings-call/press-only disclosure with no corresponding filed-document quantification — strengthens the existing Analysis.html flag that the 5GW mechanism, pricing, and terms are undisclosed in SEC filings.
- **NVIDIA Rubin platform**: 10-Q risk factors reference "our expected deployment of the NVIDIA Rubin platform in the second half of 2026" (what the call called "Vera Rubin").

## Cash Flow (Three months ended Mar 31; $M)

| Line | Q1 2025 | Q1 2026 |
|---|--:|--:|
| Net cash provided by operating activities | 61 | **2,984** |
| Purchase of property and equipment (incl. capitalized internal-use software) | (1,407) | **(7,695)** |
| Other investing (net) | (26) | (13) |
| Net cash used in investing activities | (1,433) | (7,708) |
| Proceeds from issuance of debt, net | 785 | 3,290 |
| Repayments of debt | (271) | (1,335) |
| Issuance of common stock in private placement, net | 0 | **1,985** |
| Proceeds from IPO, net | 1,423 | 0 |
| Net cash provided by financing activities | 1,854 | 3,914 |

**⚑ CapEx discrepancy (unexplained on the call):** The 10-Q's actual cash "purchase of property and equipment" is **$7,695M**, materially higher than the **"$6.8 billion" CapEx figure Nitin Agrawal stated on the call**. The gap (~$895M, ~13%) was not reconciled by management — possibilities include the call figure excluding capitalized internal-use software, being stated on an accrual/committed rather than cash-paid basis, or excluding some financed-via-debt equipment, but none of this was explained. Treat the 10-Q's $7,695M as the GAAP-accurate figure for any modeling.

**Financing activities detail**: the $1,985M "issuance of common stock in a private placement, net" line matches the NVIDIA $2.0B equity purchase (net of ~$15M issuance costs) almost exactly, confirming the cash actually landed in Q1.

**⚑ Operating cash flow ($2,984M) vastly exceeds Adjusted EBITDA ($1.2B per the call)** — a ~$1.8B gap driven primarily by deferred-revenue and working-capital timing (RPO/backlog converting to cash-in-advance faster than it converts to recognized revenue). This is a genuinely positive cash-conversion signal not highlighted on the call, though it should be read alongside the AR-decline note above rather than in isolation.

## Insider Transactions (Form 4) — representative sample, 2026-04-01 to 2026-05-22 window

CoreWeave has an unusually large number of Section-16 filers with high-frequency, pre-scheduled 10b5-1 sales (96 total Form 4 filings in this ~8-week window — far more than a typical large-cap issuer). The following 8 filings were read in full; this is a representative sample, not an exhaustive census of all 96.

| Date | Insider | Role | Transaction |
|---|---|---|---|
| 05/05/2026 (filed 05/07, day of the print) | **Michael N. Intrator** | CEO & President, Director, 10% Owner | 10b5-1 sale (plan adopted 11/20/2025): ~200,000 shares direct + 107,693 shares via **Omnadora Capital LLC** (his LLC) = **~307,693 shares** sold in seven weighted-average price tranches from **$122.71 to $129.49**. |
| 05/05/2026 (filed 05/07) | Goldberg Chen | EVP, Product & Engineering | RSU vesting 37,500 shares, then sold 19,222 shares @ $125 for tax withholding (standard, non-discretionary). |
| 05/06/2026 (filed 05/08) | Sachin Jain | Chief Operating Officer | 10b5-1 sale (plan from 9/12/2025, modified 11/20/2025): 7,335 shares @ $131.13. |
| 05/06/2026 (filed 05/08) | **Brian M. Venturo** | Chief Strategy Officer, Director (co-founder) | 10b5-1 sale (plan 11/13/2025) via **West Clay Capital LLC** + Venturo Family GST Exempt Trust: **76,924 shares** across tranches from **$131.13 to $138.19**. |
| 05/18/2026 (filed 05/20) | **Brian M. Venturo** | Chief Strategy Officer, Director | Same 10b5-1 plan, much larger tranche: **375,000 shares** (300,000 via West Clay Capital LLC + 75,000 via GST Trust, both fully liquidated) across tranches from **$98.75 to $104.39**. |

**⚑ Stock price trajectory implied by Form 4 transaction prices — a major, call-contradicting finding:**
- **05/05/2026** (2 days pre-print): weighted-average tranches **$122.71 – $129.49**
- **05/06/2026** (1 day pre-print): weighted-average tranches **$131.13 – $138.19** (still rising into the print)
- **05/18/2026** (~2 weeks post-print): weighted-average tranches **$98.75 – $104.39**

This implies the stock **fell roughly 25–30%** from its pre-print level to two weeks after the earnings call — a significant post-earnings selloff that directly contradicts the surface "record quarter, reaffirmed guidance" framing of the call. This is consistent with (and considerably amplifies) Keith Weiss's opening comment that CapEx/component pricing was "weighing on the stock a little bit after hours" — "a little" understates what the Form 4 price trail shows over the following two weeks. **This should be treated as one of the report's most important findings** and surfaced prominently in both Analysis.html (risk matrix) and Financials.html (valuation tab), since none of the qualitative Q&A concerns (gross margin trend, component-cost pass-through vagueness, customer concentration) were, on their own, evident from the call transcript as being severe enough to justify a decline of this magnitude — the market reaction reveals investors weighted these concerns more heavily than management's framing suggested they should.

**Insider signal interpretation**: All sales identified were executed under **pre-existing 10b5-1 plans adopted well before the print** (Sept–Nov 2025) — mechanically scheduled, not opportunistically timed. No evidence of unusual acceleration or a break from the pre-set plans was found in this sample. Venturo's much larger 05/18 tranche (375,000 vs 76,924 two weeks earlier) is consistent with a scheduled plan hitting a larger vesting/release tranche, not a discretionary reaction to the post-print price decline — but the sample reviewed here (8 of 96 filings) is not exhaustive and a CFO-specific filing (Nitin Agrawal) was not located in the filings sampled; a full census would be needed to rule out any late-window discretionary activity.

## Supply Chain Disclosure — confidence: medium-high (unusually specific for this type of company)

Unlike many software/cloud issuers, CoreWeave's Item 1A risk factors name suppliers and quantify concentration directly (see Debt/NVIDIA sections above for verbatim quotes): NVIDIA (sole GPU source, contractually mandated by customers), TSMC (named one tier upstream, as NVIDIA's own fabrication supplier, with explicit China-Taiwan geopolitical risk language), and unnamed "OEM and software license" vendors financed via separate arrangements. Top-3-supplier purchase concentration is disclosed by year (57%/22%/11% FY2023 → 23%/20%/17% FY2025), a level of quantification comparable to what a semiconductor manufacturer typically discloses, not what a generic cloud/software 10-K discloses (contrast with AMZN's supply-chain disclosure, which is generic boilerplate with no named suppliers or percentages — see `AMZN Q1 FY2026 SEC Cross-Reference Notes.md`).

---

僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
