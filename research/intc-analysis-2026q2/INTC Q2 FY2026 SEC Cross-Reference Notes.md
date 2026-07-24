# INTC Q2 FY2026 — SEC EDGAR Cross-Reference Notes

Fetched directly from SEC EDGAR (data.sec.gov / www.sec.gov) with identifying User-Agent, per repo's earnings-call framework. CIK 0000050863.

## Filing availability as of this report (2026-07-24)
- **Q2 FY2026 10-Q (period ended Jun 27, 2026): NOT YET FILED.** Most recent 10-Q on EDGAR is for the period ended Mar 28, 2026 (filed 2026-04-24, accession 0000050863-26-000079). Earnings were only released 2026-07-23; the 10-Q typically follows the earnings release by 1–4 weeks, so it is too fresh to exist yet.
- **Most recent 10-K: FY2025** (period ended Dec 27, 2025), filed 2026-01-23, accession 0000050863-26-000011, doc `intc-20251227.htm`. Used as the structural source for supply-chain and risk-factor disclosures below, per the same convention this repo used for GOOG Q2 FY2026 (carry forward the latest annual filing when the quarter's own 10-Q isn't indexed yet).
- Most recent 8-K: filed 2026-07-23 (accession 0000050863-26-000155) — the earnings release itself.

## Supply chain / risk factor extracts (from FY2025 10-K, Item 1 & Item 1A)
- **Sole-source dependency**: ASML Holding N.V. is the sole supplier of EUV lithography tools used in Intel 4/3/18A and planned future nodes (Item 1A, "We rely upon a complex global supply chain").
- **External foundry dependency**: TSMC exclusively manufactures some of Intel's most advanced current/future products and supplies critical compute-die components for others; leading-edge external foundry capacity is geographically concentrated in Asia.
- **Manufacturing footprint (2025)**: production fabs in Oregon (18A ramp), Arizona (Intel 7 + 18A ramp), Ireland (Intel 4/3), Israel (Intel 7); assembly & test in China, New Mexico (key advanced-packaging site), Vietnam, Malaysia (new advanced-packaging facility under construction); Costa Rica assembly/test being consolidated into Vietnam/Malaysia by end-2026; Ohio fab construction pace slowed; Germany fab and Poland assembly/test expansions discontinued.
- **Customer concentration** (Note 3, Operating Segments): three largest customers (unnamed — "Customer A/B/C") accounted for 19% / 12% / 12% = 43% of FY2025 net revenue combined, essentially unchanged from FY2024 (45%) and up from FY2023 (40%).
- **Substrate/component arrangements**: Intel has entered long-term purchase commitments and/or large prepayments with suppliers to secure future substrate and third-party foundry capacity — carries risk of excess/obsolete inventory or unrecoverable prepayments if demand disappoints.
- **US government as significant stockholder**: flagged as a distinct risk factor (subjects Intel to added regulatory scrutiny in non-US markets); this ties directly to the CHIPS Act "Secure Enclave" Warrant and Common Stock Agreement referenced in the Q2 earnings release footnote on Escrowed Shares.

## Form 4 insider-trading review (window: filings dated 2026-05-04 through 2026-06-02, the most recent batch before the 2026-07-23 call; no Form 4s have been filed yet reflecting post-earnings activity)
Parsed directly from each filing's rendered Form 4 HTML (`xslF345X06/form4.xml` under the filing's EDGAR archive folder):

| Filer | Title | Date | Transaction | Detail |
|---|---|---|---|---|
| David Zinsner | EVP, CFO | 2026-06-01 | Code M (option exercise) | +37,015 shares acquired |
| David Zinsner | EVP, CFO | 2026-06-01 | Code F (tax withholding on vesting) | −18,353 shares disposed @ $109.82 |
| Nagasubramaniyan Chandrasekaran | EVP, Chief Technology & Operations Officer, GM Intel Foundry | 2026-05-29 | **Code S (open-market sale)** | −21,024 shares @ $118.279, leaving 205,852 shares held |
| Aliyar Katouzian | (newly announced) Lead, CCPG | 2026-05-30 | Code A (RSU award) | +32,729 RSUs — new-hire/promotion grant, not a market transaction |
| Aparna Bawa | (newly announced) Chief Legal/Ethics/Compliance/People/Culture Officer | 2026-05-30 | Code A (RSU award) | +27,274 RSUs — new-hire/promotion grant, not a market transaction |

**Read**: The only discretionary open-market sale in this window is Naga Chandrasekaran's — notable because he is the executive most directly responsible for the 18A/14A manufacturing ramp central to this quarter's bullish narrative. It predates the earnings beat and CapEx raise by nearly two months and was at $118.279, a price this report cannot yet compare to post-earnings levels (no trading-price feed was fetched). Absent a Form 4 disclosing a Rule 10b5-1 plan flag (the filings' plan checkbox was not parsed in this pass), this cannot be distinguished from routine, pre-scheduled diversification — but it is the one non-routine, non-grant, non-withholding transaction among Intel's newly elevated leadership team in the pre-earnings window, and is carried into the Risk Matrix tab as a low-to-medium-confidence data point, not a red flag.

## Explicitly not verified this cycle
- Conversion mechanics, milestone schedule, and remaining balance of the CHIPS Act Escrowed Shares arrangement — the earnings release footnote is the only source; the 10-Q (once filed) should carry the full contractual description.
- The counterparty and structure behind the $12.2B net partner-distribution outflow in Q2 cash flow (labeled "Partner contributions"/"Partner distributions" in the release) — plausibly Intel's semiconductor co-investment (SCIP-style) partnerships such as the Arizona fab joint ventures, but this report does not have primary-source confirmation and states so explicitly in the Cash Flow tab.
