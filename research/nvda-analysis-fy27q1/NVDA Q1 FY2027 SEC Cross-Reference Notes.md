# NVIDIA Corporation (NVDA) Q1 FY2027 — SEC Cross-Reference Notes

Sources pulled from SEC EDGAR (CIK 0001045810), User-Agent identified per SEC policy.

- **10-Q** filed 2026-05-20, period ended 2026-04-26 — accession 0001045810-26-000052 (`nvda-20260426.htm`)
- **10-K** filed 2026-02-25, FY2026 (period ended 2026-01-25) — accession 0001045810-26-000021 (`nvda-20260125.htm`)
- **Form 4** insider filings reviewed: full 2026 issuer feed (46 filings, Jan–Jul 2026), with focus on the Mar–Jun window bracketing the call

Call date: **2026-05-20, 5:00pm ET** (same day the 10-Q was filed). All P&L figures below are the **10-Q as filed**, which supersede any rounded numbers spoken on the call.

---

## Income Statement (Three months ended Apr 26; $M)

| Line | Q1 FY2026 | Q1 FY2027 | YoY |
|---|--:|--:|--:|
| Revenue | 44,062 | 81,615 | +85.3% |
| Cost of revenue | 17,394 | 20,458 | +17.6% |
| Gross margin | 60.5% | 74.9% | +14.4pp |
| R&D | 3,989 | 6,321 | +58.5% |
| SG&A | 1,041 | 1,300 | +24.9% |
| **Operating income** | 21,638 | **53,536** | +147.4% |
| Interest income | 515 | 540 | — |
| Interest expense | (63) | (102) | — |
| Other income (expense), net | (180) | **15,929** | — |
| Income before tax | 21,910 | 69,903 | +219% |
| Income tax expense (eff. rate) | 3,135 (14.3%) | 11,582 (16.6%) | — |
| **Net income** | 18,775 | **58,321** | +210.7% |
| Diluted EPS | $0.76 | **$2.39** | +214.5% |
| Diluted shares | 24,611 | 24,391 | — |

**⚑ Quality-of-earnings flag:** Net income +211% outruns operating income +147% (itself already a blowout, genuinely revenue-driven — this is not a GOOG-style "op income normal, net income fake" pattern). The gap is a **$16.1B swing in Other income**, driven by unrealized gains of **$13.4B on publicly-held equity securities** and **$2.6B on non-marketable equity securities** (10-Q MD&A, "Total Other Income, Net"). Non-marketable securities balance jumped $22.3B→$42.3B in the quarter ($17.9B net additions + $2.6B unrealized gains); FY2026 10-K attributes part of NVIDIA's investment-gain history to its previously-disclosed stake in **Intel common stock**, though the Q1 FY27 filing does not itemize which holdings drove this quarter's $13.4B public-equity gain. Tax-effecting operating income alone at the quarter's 16.6% rate implies "core" net income of **~$44.6B** and core diluted EPS of **~$1.83** — vs the reported $58.3B / $2.39. See Financials → Income Statement and Valuation tabs for the full walk.

## Revenue by Market Platform ($M) — new reporting framework this quarter

| Line | Q1 FY2026 | Q4 FY2026 | Q1 FY2027 | YoY | QoQ |
|---|--:|--:|--:|--:|--:|
| Data Center | 39,112 | 62,314 | **75,246** | +92.4% | +20.8% |
| — Hyperscale | 17,599 | 33,814 | 37,869 | +115.2% | +12.0% |
| — AI Clouds, Industrial & Enterprise (ACIE) | 21,513 | 28,500 | 37,377 | +73.7% | +31.2% |
| Edge Computing | 4,950 | 5,813 | 6,369 | +28.7% | +9.6% |
| **Total revenue** | 44,062 | 68,127 | **81,615** | +85.3% | +19.8% |

Reportable segments (accounting basis, different cut): Compute & Networking $74,550M (+88% YoY, op income $53,335M) / Graphics $7,065M (+58% YoY, op income $2,941M).

No Data Center Hopper shipments to China this quarter, vs $4.6B in Q1 FY2026. H200 licenses granted from Feb 2026 but **zero revenue recognized to date**; any future shipment carries a 25% import tariff.

## Customer concentration & geography

- **Direct customers:** three represented **21%, 17%, 16%** of Q1 FY27 revenue (all Compute & Networking) — up from two at 16%/14% in Q1 FY26. FY2026 (full year, per 10-K): one at 22%, one at 14%.
- **Indirect customers:** "one AI research and deployment company contributed to a meaningful amount of revenue by purchasing cloud services from our customers" (10-Q Note 13) — unnamed.
- **Accounts receivable concentration:** three direct customers = 30%, 18%, 16% of AR balance at Apr 26, 2026 (25%/18%/13% at Jan 25, 2026) — concentration is rising faster in AR than in revenue.
- **Geography (by customer HQ):** US $63,769M (78%) / Taiwan $12,006M (15%) / China incl. HK $4,550M (6%) / Other $1,290M. Ex-US revenue fell to 22% of total from 42% a year ago — mechanically due to the Hopper-China falloff, not a demand shift.

## Balance sheet & capital return

- Total assets $259.5B (Jan 25 2026: $206.8B). Cash+marketable debt securities $50.3B; marketable equity securities $30.2B (Jan 25: $12.9B); non-marketable securities $42.3B (Jan 25: $22.3B).
- Inventories $25.8B, up from $21.4B — raw materials nearly doubled ($3.8B→$6.6B).
- Share repurchases: 108M shares for $20.2B in Q1. **May 18, 2026: Board approved additional $80.0B buyback authorization** (on top of $38.5B remaining) and **raised the quarterly dividend from $0.01 to $0.25/share** (record date Jun 4, paid Jun 26).
- Commitments (Note 10, as of Apr 26 2026): manufacturing/supply/capacity commitments **$119B** ($95B due within FY2027); multi-year cloud service agreement commitments **$30B**; other vendor commitments **$6B**. (These, plus $25.8B on-hand inventory, are the components behind the CFO's "$145B total supply" call comment — not a single filed line item.)

## Manufacturing / supply chain (10-K Item 1, "Manufacturing")

Fabless model. Named suppliers:
- **Foundries:** Taiwan Semiconductor Manufacturing Co. (TSMC), Samsung Electronics
- **Memory:** SK Hynix, Micron Technology, Samsung
- **Advanced packaging:** CoWoS (TSMC process)
- **Assembly/test/packaging subcontractors:** Hon Hai Precision Industry (Foxconn), Wistron Corporation, Fabrinet

10-K: "supply chain is mainly concentrated in Asia, we are expanding into the U.S. and Latin America." No individual supplier $ allocation disclosed — the Sankey diagram in Analysis.html's supply-chain tab uses illustrative relative weights for the upstream side (flagged as such); the NVIDIA→market-platform side uses the real Note 13 revenue split above.

**Groq IP license (10-K risk factors + 10-Q Note 7):** NVIDIA "entered into an intellectual property license arrangement with Groq, Inc., that required significant, nonrefundable payments." 10-Q Note 7 shows **$3,957M "accrued purchase consideration... related to the Groq, Inc. non-exclusive license agreement"** sitting in accrued liabilities. Groq is known publicly for SRAM-based, low-latency inference chips (LPUs) — the same category Jensen described as "LPX" on the call (low-latency, high-token-rate, SRAM-based, "niche… for some time to come"). The call never mentions Groq by name. See Analysis.html Q&A (Arcuri question) and Risk Matrix.

## Form 4 / insider signal window

- **Blackout confirmed:** zero Form 4 filings between 2026-03-24 and 2026-05-29 — a clean ~9-week window bracketing the May 20 print, consistent with standard trading-window policy rather than a discretionary pause.
- **CEO (Jensen Huang):** filed 03/02, 03/18 (pre-blackout) and resumed 06/18, 06/23 (post-print). The 06/23 filing (transaction 06/17) is a **Code "F"** transaction (shares withheld for tax on vesting, 45,723 shares @ $207.41) — not a discretionary open-market sale. Huang's direct + trust/LLC holdings run into the hundreds of millions of shares; his Form 4 activity this window shows no acceleration vs prior quarters.
- **CFO (Colette Kress):** recurring Rule 10b5-1 plan (adopted March 4, 2025); filed 01/13, 02/04, 03/02, 03/18, resumed 06/17 — steady cadence, no step-change in size.
- **Directors:** a cluster of 8 independent directors each filed small routine 10b5-1 sales on 2026-06-29 (transaction date 06/25) — e.g. John Dabiri sold 625 shares @ $214 on 05/27 under a plan adopted Dec 10, 2025. Standard quarterly pattern, immaterial size.
- **Net read:** no red flags and no bullish tell either — insider activity this window is procedural (blackout → scheduled 10b5-1 resumption), not sentiment-driven. See Risk Matrix (Low severity item).

## Stock price reference points (not in filings — public market data, used only to illustrate multiples)

- Fri 2026-05-19 (day before the print) close: **$222.32**, market cap ~$5.40T
- 2026-07-17 (most recent close pre-report): **$202.81**, 52-week range $164.07–$236.54

---
僅供資訊參考，非投資建議 / For informational purposes only — not investment advice.
