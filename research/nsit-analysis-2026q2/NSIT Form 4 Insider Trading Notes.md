# NSIT Form 4 Insider Trading Notes — 2026-06-01 through 2026-09-14

Source: SEC EDGAR submissions for CIK 0000932696, https://data.sec.gov/submissions/CIK0000932696.json (form=4 filings only, filed 2026-06-01 through 2026-09-14)

Earnings call: **2026-08-06** (before market). 10-Q filed same day, **2026-08-06**.

## All Form 4 filings in the window

| Filed | Period/Txn date | Filer | Role | Code | Shares | Price | Accession |
|---|---|---|---|---|---|---|---|
| 2026-06-02 | 2026-05-29 | Karim Adatia | General Counsel (officer) | S (sale) | 559 | $107.40 (wtd. avg.) | 0001193125-26-253540 |
| 2026-08-17 | 2026-08-13 | Anthony Ibarguen | Director | S (sale) | 4,000 | $154.83 (wtd. avg., range $154.50–$155.28) | 0001628280-26-057365 |
| 2026-09-01 | 2026-08-30 | Thomas Reichert | Director | M/A (RSU vest) | 217 | $0.00 (RSU settlement) | 0001628280-26-059787 |
| 2026-09-01 | 2026-08-30 | Janet Foutty | Director | M/A (RSU vest) | 217 | $0.00 (RSU settlement) | 0001628280-26-059798 |
| 2026-09-01 | 2026-08-28 | Karim Adatia | General Counsel (officer) | S (sale) | 248 | $157.40 | 0001628280-26-059836 |

**No Form 4 filings at all between 2026-06-03 and 2026-08-12** — a genuine ~10-week gap covering the entire pre-earnings quiet period and the earnings call itself (2026-08-06). This is itself a mildly notable data point: **no insider (director or officer) traded in the days immediately before or on the earnings call date**, and the first post-call transaction (Ibarguen's sale) didn't occur until a full week after the print (2026-08-13), filed four days later (2026-08-17) — comfortably outside any window that would suggest trading on advance knowledge of the beat-and-raise results.

## Detail on transactions nearest the print

- **Anthony Ibarguen (Director), 2026-08-13 sale, 4,000 shares, $154.50–$155.28/share (~$619K total)** — the only open-market sale in the window that falls reasonably close to the earnings call, occurring 5 trading days after the 2026-08-06 report and 7 days after the 10-Q filing. Footnote explicitly identifies this as **multiple transactions at a weighted-average price**, standard block-sale mechanics, not a single print. Retains 18,188 shares after the sale (a majority of a presumably larger prior holding — sold roughly 18% of the resulting post-sale position). **`aff10b5One` field = 0 (false)** — explicitly **not** made pursuant to a Rule 10b5-1 trading plan.
- **Karim Adatia (General Counsel), two small sales**: 559 shares on 2026-05-29 (filed late, 2026-06-02 — before the window's Q2 activity, included for completeness) at $107.40, and 248 shares on 2026-08-28 at $157.40 (**after** shares owned following transaction = 0, i.e., this fully liquidated his direct holding as of that filing). Both explicitly **not** 10b5-1 plan transactions (`aff10b5One` = false/0 in both filings).
- **Reichert and Foutty (both Directors), 2026-08-30**: identical mechanical RSU vesting/settlement transactions (217 shares each, Code M/A at $0), from RSU grants dated **August 30, 2024** vesting in three equal annual installments beginning **August 30, 2025** — i.e., this is the second of three scheduled annual vesting tranches, pure compensation mechanics with zero discretionary trading signal. Not 10b5-1-flagged (not applicable to a vesting event).

## 10b5-1 plan mentions

**None of the five Form 4s in the window are marked as executed under a Rule 10b5-1 trading plan** (`aff10b5One` = false/0 in all five XML filings). This means every open-market sale in the window (Adatia ×2, Ibarguen ×1) was a discretionary sale, not a pre-scheduled plan sale — worth noting for the risk matrix, though the dollar amounts involved are small (largest single sale ≈$619K by a director) relative to Insight's ~$3.2B market cap context and don't constitute a material insider-selling signal.

## Read for the risk matrix / Q&A tab

- **No cluster of Section 16 officer or director selling immediately before or after the beat-and-raise print** — the complete absence of any Form 4 in the ~10-week window bracketing the quiet period and the call itself (2026-06-03 to 2026-08-12) is a mildly reassuring (if weak, since absence of Form 4s during a quiet period is also just standard company insider-trading-policy compliance, not necessarily bullish signal) data point.
- The one meaningful open-market sale (Ibarguen, director, ~$619K, one week post-call) and the two small GC sales (Adatia, totaling ~$100K combined) are all modest in size and **not** 10b5-1 scheduled — but also not large enough, or clustered enough, to read as a coordinated "sell the guidance raise" signal. No CEO (Jack Azagury, new CEO effective Apr 13, 2026) or CFO (James Morgado) Form 4 activity appears anywhere in the window — notably, neither the new CEO nor the CFO sold or bought shares around this print, which given Azagury's very recent CEO transition (still inside his first ~4 months per his own call commentary — "my first four months, Adam, have been solely focused on our organic business") is unsurprising but worth flagging as an absence, not a positive signal either way.
- Karim Adatia's Aug 28 sale reducing his direct holding to **zero** shares owned following the transaction is worth a light flag — full liquidation of a direct stake by the General Counsel — though Form 4 doesn't disclose whether he retains indirect holdings (trusts, etc.) or unvested equity awards not yet reportable, so this should not be over-read as a bearish signal without more context.
