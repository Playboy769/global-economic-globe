# NSIT Sources — SEC EDGAR URLs and accessions used (CIK 0000932696)

## Core filing lists
- Submissions (all filings): https://data.sec.gov/submissions/CIK0000932696.json
- Company facts (XBRL, all periods): https://data.sec.gov/api/xbrl/companyfacts/CIK0000932696.json

## Q2 FY2026 (period ended 2026-06-30) — primary filing under analysis
- 10-Q, filed 2026-08-06, acc 0000932696-26-000070: https://www.sec.gov/Archives/edgar/data/932696/000093269626000070/nsit-20260630.htm
- 10-Q filing index: https://www.sec.gov/Archives/edgar/data/932696/000093269626000070/index.json
- 8-K, filed 2026-08-06, acc 0000932696-26-000067: https://www.sec.gov/Archives/edgar/data/932696/000093269626000067/nsit-20260806.htm
- 8-K Exhibit 99.1 (earnings press release), same accession: https://www.sec.gov/Archives/edgar/data/932696/000093269626000067/nsit-20260806xerx8kex991.htm
- 8-K filing index: https://www.sec.gov/Archives/edgar/data/932696/000093269626000067/index.json

## Prior-quarter 10-Qs and 10-K (for 6-quarter trend table)
- Q1 FY2026 10-Q, filed 2026-05-07, acc 0000932696-26-000050: https://www.sec.gov/Archives/edgar/data/932696/000093269626000050/nsit-20260331.htm
- FY2025 10-K, filed 2026-02-12, acc 0000932696-26-000007: https://www.sec.gov/Archives/edgar/data/932696/000093269626000007/nsit-20251231.htm
- Q3 FY2025 10-Q, filed 2025-10-30, acc 0000932696-25-000019: https://www.sec.gov/Archives/edgar/data/932696/000093269625000019/nsit-20250930.htm
- Q2 FY2025 10-Q, filed 2025-07-31, acc 0001628280-25-036953: https://www.sec.gov/Archives/edgar/data/932696/000162828025036953/nsit-20250630.htm (referenced via XBRL companyfacts; full text not separately downloaded — all figures cross-checked against Q2 FY26 10-Q's own comparative-period columns, which are identical)
- Q1 FY2025 10-Q, filed 2025-05-01, acc 0001628280-25-021416: https://www.sec.gov/Archives/edgar/data/932696/000162828025021416/nsit-20250331.htm

## Form 4 filings (2026-06-01 through 2026-09-14)
- Karim Adatia, filed 2026-06-02 (period 2026-05-29), acc 0001193125-26-253540: https://www.sec.gov/Archives/edgar/data/932696/000119312526253540/ownership.xml
- Anthony Ibarguen, filed 2026-08-17 (period 2026-08-13), acc 0001628280-26-057365: https://www.sec.gov/Archives/edgar/data/932696/000162828026057365/wk-form4_1786997199.xml
- Thomas Reichert, filed 2026-09-01 (period 2026-08-30), acc 0001628280-26-059787: https://www.sec.gov/Archives/edgar/data/932696/000162828026059787/wk-form4_1788293178.xml
- Janet Foutty, filed 2026-09-01 (period 2026-08-30), acc 0001628280-26-059798: https://www.sec.gov/Archives/edgar/data/932696/000162828026059798/wk-form4_1788293359.xml
- Karim Adatia, filed 2026-09-01 (period 2026-08-28), acc 0001628280-26-059836: https://www.sec.gov/Archives/edgar/data/932696/000162828026059836/wk-form4_1788293781.xml

## Not used / explicitly checked and found not applicable
- No Form 4 filings exist between 2026-06-03 and 2026-08-12 (confirmed via submissions.json full filter, form=4, date range 2026-01-01 to 2026-09-14).
- No 8-K Exhibit with a separate investor slide deck (PDF/PPT) was found attached to the 2026-08-06 8-K on EDGAR — the earnings call slide presentation referenced in the transcript ("the accompanying slide presentation that was posted this morning") is hosted on investor.insight.com and was not filed as a distinct SEC exhibit; the $171M cloud GP / $95M core services GP figures cited on the call could not be verified against any SEC-filed document.
- Google and Microsoft partner-program-change specifics are not named in either the 10-Q or 8-K text (0 hits on "Google", "Microsoft" in the Q2 FY26 10-Q) — only in the earnings-call transcript.

## Local scratchpad files (raw downloads + cleaned text, for reproducibility)
All under `C:/Users/ryan9/AppData/Local/Temp/claude/C--Users-ryan9-OneDrive----Claudecode/a5d223a8-96d5-4d92-a1a0-5fffce56b570/scratchpad/nsit/`:
- `submissions.json`, `companyfacts.json` — raw SEC API responses
- `q2fy26_10q.htm` / `.txt`, `q2fy26_8kex991.htm` / `.txt` — Q2 FY26 10-Q and 8-K Ex-99.1
- `q1fy26_10q.htm` / `.txt`, `q3fy25_10q.htm` / `.txt`, `q1fy25_10q.htm` / `.txt` — comparative 10-Qs
- `fy25_10k.htm` / `.txt` — FY2025 10-K
- `form4/*.xml` — all five Form 4 XML filings in the window
- `extract.py`, `us-gaap-tags.txt` — XBRL extraction script and full tag inventory
