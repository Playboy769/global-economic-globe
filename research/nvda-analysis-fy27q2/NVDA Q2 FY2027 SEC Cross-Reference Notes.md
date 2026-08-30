# NVIDIA Corporation (NVDA) Q2 FY2027 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK 0001045810), User-Agent identified per SEC policy.

- **10-Q** filed 2026-08-26, period ended 2026-07-26 — accession 0001045810-26-000075 (`nvda-20260726.htm`)
- **10-Q (Q1 FY2027)** filed 2026-05-20, period ended 2026-04-26 — accession 0001045810-26-000052, used for QoQ balance-sheet deltas not shown in the Q2 filing's own two-column balance sheet (which only compares to FY2026 year-end)
- **10-K** filed 2026-02-25, FY2026 (period ended 2026-01-25) — accession 0001045810-26-000021, used to back out Q4 FY2026 standalone figures (10-Q XBRL feeds only carry discrete Q1–Q3 durations; Q4 is derived as FY − 9-month)
- **Form 4** insider filings reviewed: full issuer feed 2026-06-15 through 2026-08-30 (20 filings)
- **XBRL company-concept API** (`data.sec.gov/api/xbrl/companyconcept`) used for `Revenues` and `GrossProfit` to build the 6-quarter trend series, cross-checked against the Q1 FY2027 10-Q's own MD&A comparison table for Q4 FY2026 ($68,127M — matches the derived figure exactly)

Call date: **2026-08-26, 5:00pm ET** (same day the 10-Q and an 8-K were filed). All P&L figures below are the **10-Q as filed**, which supersede any rounded numbers spoken on the call.

---

## Income Statement (Three months ended Jul 26; $M)

| Line | Q2 FY2026 | Q1 FY2027 | Q2 FY2027 | QoQ | YoY |
|---|--:|--:|--:|--:|--:|
| Revenue | 46,743 | 81,615 | 96,221 | +18% | +106% |
| Cost of revenue | 12,890 | 20,458 | 24,079 | — | — |
| Gross margin | 72.4% | 74.9% | 75.0% | +0.1pp | +2.6pp |
| R&D | 4,291 | 6,321 | 7,054 | — | +64% |
| SG&A | 1,122 | 1,300 | 1,354 | — | +21% |
| Operating expenses | 5,413 | 7,621 | 8,408 | +10% | +55% |
| **Operating income** | 28,440 | 53,536 | **63,734** | +19% | +124% |
| Other income, net | 2,766 | 15,929* | 7,773 | — | — |
| Income before tax | 31,206 | 69,903* | 71,507 | — | — |
| Income tax expense (eff. rate) | 4,784 (15.3%) | 11,582 (16.6%)* | 11,819 (16.5%) | — | — |
| **Net income** | 26,422 | 58,321 | **59,688** | +2% | +126% |
| Diluted EPS | $1.08 | $2.39 | **$2.46** | +3% | +128% |
| Diluted shares | 24,532 | 24,391 | 24,285 | — | — |

*Q1 FY2027 figures per the Q1 SEC Cross-Reference Notes.

**⚑ Quality-of-earnings flag (second consecutive quarter):** Net income growth (+126% YoY) again outruns operating income growth (+124% YoY) — smaller gap than Q1's blowout, but the *composition* is what matters: pretax "Other income, net" was **$7,773M in Q2**, of which **$7,771M was unrealized/realized gains on equity securities** (10-Q Note 7, "Other Income, Net"). Tax-effecting operating income alone at the quarter's 16.5% effective rate implies **"core" net income of ≈$53.2B and core diluted EPS of ≈$2.19** — vs. the reported $59.7B / $2.46. The gap (~$0.27/share, ~11% of reported EPS) is smaller than Q1's (~$0.56/share on a much larger $16.1B other-income swing), suggesting the equity-gains contribution to headline EPS is moderating sequentially even as the dollar amount ($7.77B) is still large in absolute terms. See Valuation tab for the full walk and the H1 cumulative gain ($23.7B).

## Revenue by Market Platform ($M)

| Line | Q2 FY2026 | Q1 FY2027 | Q2 FY2027 | QoQ | YoY |
|---|--:|--:|--:|--:|--:|
| Data Center | 41,096 | 75,246 | **89,023** | +18% | +117% |
| — Hyperscale | 24,168 | 43,050 | 48,710 | +13% | +102% |
| — AI Clouds, Industrial & Enterprise (ACIE) | 16,928 | 32,196 | 40,313 | +25% | +138% |
| Edge Computing | 5,647 | 6,369 | 7,198 | +13% | +27% |
| **Total revenue** | 46,743 | 81,615 | **96,221** | +18% | +106% |

⚑ Footnote in the 10-Q: "During the second quarter of fiscal year 2027, we reclassified a company from ACIE to Hyperscale due to a change in their business model and recast the prior period revenue associated with this company" — the Q1 FY2027 ACIE/Hyperscale split shown above is the **recast** figure, not what was originally reported in May. The unnamed company's business-model change (moving from an ACIE-style neocloud/enterprise classification to Hyperscale) is itself a data point worth flagging — it mechanically flatters Hyperscale's YoY comp and is exactly the kind of reclassification the Q&A tab's Stacy Rasgon exchange predicted would matter once neocloud/hyperscaler lines blur.

Reportable segments (accounting basis, different cut): Compute & Networking $88,299M (+114% YoY, op income $62,696M, +121% YoY) / Graphics $7,922M (+46% YoY, op income $3,899M, +74% YoY).

Shipments of Data Center Hopper (H200) products to China during Q2 FY2027 were **less than 1% of Data Center revenue** — consistent with the call. A **$0.4 billion charge** was incurred in H1 FY2027 for H200 excess inventory/purchase obligations as demand diminished; any future H200 shipments carry a **25% import tariff** NVIDIA says it cannot pass through to customers.

## Customer concentration & geography

- **Direct customers:** Q2 FY27 alone — **one** direct customer at 16% of total revenue (down from *two* customers at 23%/16% in Q2 FY26). H1 FY27 (cumulative) — three direct customers at 16%, 15%, 13%.
- **Accounts receivable concentration — sharply up:** **five** direct customers = 22%, 14%, 13%, 11%, 10% of the AR balance at Jul 26, 2026 (**70% of AR in five names**) — vs. three customers at 25%/18%/13% (56%) at Jan 25, 2026, and vs. Q1 FY27's own three-customer, 30/18/16% (64%) reading. Concentration in AR is now running well ahead of concentration in revenue (16% top customer) — consistent with extended payment terms on a handful of very large multi-quarter deals.
- **Indirect customers:** "one AI research and deployment company contributed a meaningful amount of revenue by purchasing cloud services from our customers" (10-Q, unnamed) — identical boilerplate to Q1, still unnamed.
- **Geography (by customer HQ):** Ex-US revenue rose to **38%** of total in Q2 FY27 (was 22% in Q1 FY27, 30% in Q2 FY26) — Taiwan alone jumped to $26,985M (28% of revenue) from $12,006M in Q1, consistent with Blackwell Ultra system build concentration at Taiwanese ODMs/OSATs, not a change in end-demand geography.

## Balance sheet & liquidity

- Total assets $320.3B (Jan 25 2026: $206.8B; Apr 26 2026: $259.5B).
- **Accounts receivable, net: $63,059M** (Apr 26, 2026: $40,710M; Jan 25, 2026: $38,466M) — **+55% QoQ**, far outpacing revenue's +18% QoQ.
- **Days Sales Outstanding: ≈60 days in Q2** (AR ÷ revenue × 91) vs. **≈45 days in Q1** — matches Colette Kress's own call comment ("DSO increased to 60 days... extended payment terms for large purchases by certain investment-grade customers to be shipped over multiple quarters") almost exactly. This is the single cleanest management-claim-to-filing cross-check available this quarter.
- Inventories $31.6B (raw materials alone $11.3B, up from $3.8B at FY start) — building ahead of Vera Rubin ramp.
- Cash, cash equivalents & marketable debt securities $56.6B; marketable equity securities $42.8B.
- **Operating cash flow, Q2 standalone (H1 $74,421M − Q1 $50,344M) ≈ $24.1B** — well below Q2 net income ($59.7B), almost entirely explained by the $22.3B AR build in the quarter. Free cash flow (Q2 standalone, CapEx ≈$2.68B) ≈ **$21.4B**, down from Q1's ≈$48.6B FCF despite higher revenue and earnings — a genuine, filings-confirmed instance of "record income, weak cash conversion," driven by the DSO story above rather than by margin or profitability deterioration.
- **Supply and capacity commitments jumped from $119B (last quarter) to $279B** as of Jul 26, 2026 (Note 10) — by far the largest single QoQ change in the filing. Full commitment stack by year: Supply & capacity $279B, Cloud service agreements $29B, Data center leases not yet commenced (own use) $25B, Equity investments $25B, Capital expenditures $8B = **$366B total**. A second table adds AI-cloud-partner commitments ($36B) and third-party data center leases not yet commenced ($20B) = **$56B** more.
- **Guarantees:** Land/power/shell guarantees for AI clouds ($3.5B max exposure, unchanged from Q1) **plus new SB Energy Corp. guarantees capped at $105.0B** (August 2026, PORTS Technology Campus, Pike County OH, ~4.25GW for OpenAI, option for +3.8GW) = **$108.5B total guarantee exposure**, a number that did not exist in the Q1 filing at all.
- **Customer prepayments (contract liabilities/deferred revenue):** balance grew from $2,572M (start of H1) to **$6,412M** (Jul 26, 2026), driven by $17.7B of additions in H1 including **$15.6B of customer advances**; $13.9B was recognized as revenue in the same period, including $13.0B of customer advances. Also disclosed: $3.2B of remaining performance obligations from contracts >1 year, of which ~39% recognizes over the next twelve months.
- **Interest coverage:** Operating income $63,734M ÷ interest expense $227M (Q2) ≈ **280x** — trivial leverage risk despite the new $25B senior notes issuance in June 2026; net debt is negative given $56.6B+ of cash/marketable debt securities against $33.4B of gross debt.
- **Groq, Inc. accrued purchase consideration fell from $3,921M (Jan 25, 2026) to $986M (Jul 26, 2026)** — the financing-activities cash flow statement shows a **$2,944M "Groq, Inc." outflow** in H1 FY2027, i.e., NVIDIA paid down most of the IP-license liability flagged in the Q1 report. Ties directly back to the Q1 Q&A tab's Arcuri/LPX cross-reference.
- **Public company warrants** received in Q2 FY2027 (fair value $824M as of Jul 26, 2026) and a new $1,000M notional equity forward contract — both new this quarter, sourced from otherwise-undisclosed counterparties; flagged in Special Events tab as a minor, currently unexplained item.

## Manufacturing / supply chain

No change from Q1: fabless model, same named suppliers (TSMC, Samsung Electronics; SK Hynix, Micron for memory; Foxconn, Wistron, Fabrinet for assembly/test; CoWoS advanced packaging via TSMC; Groq non-exclusive IP license). Supply/capacity commitments more than doubled QoQ ($119B→$279B) — see Balance Sheet above.

## Form 4 / insider signal window (2026-06-15 through 2026-08-30)

- **CEO (Jensen Huang) & CFO (Colette Kress):** last filed **2026-06-23** (post-Q1-print trading-window resumption). **No new Form 4 from either since** — i.e., through the Aug 26 print and four calendar days after, consistent with the standard pre-earnings blackout extending a few trading days past the print; Aug 31/Sept 1 is a US holiday weekend (Labor Day), so a post-print resumption filing may simply not have posted yet as of this report's compilation date (Aug 30, 2026).
- **Directors — routine cluster:** Stevens, Shah, Seawell, Neal, Lora, Jones, Hudson, Dabiri, Coxe all filed 2026-06-23/06-29 (post-Q1 blackout resumption, same pattern as Q1). **Tench Coxe filed twice more** (2026-07-06, 2026-08-07), both **Code "G" (bona fide gift)** — charitable share transfers, not open-market sales, no pricing signal.
- **Suzanne Nora Johnson** filed 2026-08-12 with two **Code "A" (award/grant)** transactions — consistent with a routine director equity grant, not a market transaction.
- **⚑ Notable — Section 16 officer departure, 2 days before the print:** **Ajay K. Puri, EVP Worldwide Field Operations**, filed a Form 4 on **2026-08-24** disclosing he **"retired from his role... effective August 24, 2026 and is no longer subject to Section 16."** No share transaction accompanies the filing (both `nonDerivativeTable` and `derivativeTable` are empty) — this is a pure departure notice. Puri ran NVIDIA's global sales organization; his exit was **not mentioned anywhere on the earnings call** by either management or any of the eight analysts, and no 8-K in the reviewed window discusses a successor or the reason for departure. Flagged as a governance/succession item in the Risk Matrix — timing (two trading days pre-print) invites scrutiny even though a field-ops leadership retirement carries no automatic negative inference on its own.
- **Net read:** aside from the Puri departure (a governance/disclosure gap, not a trading signal), insider activity this window is procedural — blackout, then scheduled resumption, gifts, and a routine grant. No accelerated selling by the CEO, CFO, or any director.

## Stock price reference points (not in filings — public market data, used only to illustrate multiples; sourced via web search with cross-source variance noted)

- 2026-08-28 close (2 trading days after the print): **$217.89**
- Peer forward P/E and revenue-growth figures (AMD, AVGO, TSM) pulled from multiple finance-data aggregators showed material cross-source disagreement (e.g., AVGO forward P/E cited anywhere from ~19x to ~24x, TSM from ~19x to ~27x depending on the estimate basis/date) — treated as illustrative ranges only in the Valuation tab, not as filings-grade precision.

---
僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
