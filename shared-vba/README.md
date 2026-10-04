# shared-vba/

VBA modules shared **verbatim** between two Excel workbooks:

| Workbook | Path | How it uses these |
|---|---|---|
| SEC-Filing-Fetcher (standalone) | `SEC-Filing-Fetcher/SECFilingFetcher (NN).xlsm` | `modHttp` / `modJsonUtil` / `modPrices` are its own HTTP/JSON/price layer. `modSECData` is **new** — sits alongside the full `modSEC` (see "Convergence" below). |
| RR4 portfolio | `Compound RR4 Portfolio 2026 H2.xlsm` | The "Earnings" page (`RR4/modEarnings.bas`, nav code E; until 2026-09-13 the lower half of "Company research") calls `modSECData.GetXbrlFinancials` / `modMOPSData.GetTwFinancials` / `GetTwMonthlyRevenue` and `modPrices`. RR4 does **not** get `modSEC` / `modMOPS` / `modCharts` / `modTheme`. |

Background: `RR4/company-research-sec-merge-plan.md` + `RR4/company-research-p0-dependency-audit.md`.

## Files

| File | Origin | Public surface |
|---|---|---|
| `modHttp.bas` | byte-identical copy of `SEC-Filing-Fetcher/modHttp.bas` | `HttpGet(url)`, `Const SEC_USER_AGENT` |
| `modJsonUtil.bas` | byte-identical copy | `ExtractJsonValueRaw`, `ExtractJsonString`, `JsonUnescape`, `SplitJsonArrayElements`, `FindKeyValueStart` |
| `modPrices.bas` | byte-identical copy | `FetchDailyPrices`, `LatestPrice`, `NearestPriceOnOrBefore` |
| `modSECData.bas` | **new**; its Private helpers are verbatim from `modSEC.bas` (P0 snapshot) | `GetXbrlFinancials(tickerOrCik, wantAnnual, wantQuarter)`, `XbrlValue(maps, key, accn, reportDate, form)`, `XbrlFyFp(allMaps, accn, reportDate, form)` |

`modSECData` depends only on `modHttp` + `modJsonUtil`. It has **no** dependency on any
worksheet, on `modCharts`, or on `modTheme`. Every name it shares with `modSEC` is `Private`;
its two public readers are renamed (`XbrlValue` / `XbrlFyFp`, not `LookupConceptValue` /
`LookupFyFp`) so it can be imported into the same VBAProject as the full `modSEC` with **zero
"Ambiguous name detected"**.

## Sync rule — programmatic injection ONLY, never manual VBE "Import File"

`modHttp.bas` contains Chinese error strings; `modJsonUtil.bas` has a UTF-8 BOM. Manual
**File > Import File** in the VBE decodes `.bas` with the system ANSI codepage and turns both
into mojibake (this repo has hit that before — see the global memory `vba-bas-ascii-only` and
the mangled `RR4/Sanner.bas` banner string).

So the sync script (`scripts/sync-shared-vba.ps1`, to be added) MUST inject via the VBIDE
object model, the same mechanism `SEC-Filing-Fetcher/build.ps1` already uses:

```
# for each workbook, for each shared module:
$vbproj.VBComponents.Remove($existing)      # if present
$vbproj.VBComponents.Import($absolutePathToBas)
```

`Import` reads the file as UTF-8 and preserves the `Attribute VB_Name` line. Requires the
one-time Excel setting **Trust access to the VBA project object model** (File > Options >
Trust Center > Macro Settings) on the machine running the script.

`modSECData.bas` itself is pure ASCII — keep it that way (no Chinese literals), so that even a
manual import can't corrupt it.

## Verifying a change (per CLAUDE.md "SEC-Filing-Fetcher 慣例")

`build.ps1` printing SUCCESS does **not** mean the VBA runs. After injecting, verify with Excel
COM (`$excel.AutomationSecurity = 1`):

```
# smoke test the headless reader directly, no sheets involved:
$res = $excel.Run("modSECData.GetXbrlFinancials", "NVDA", 4, 8)
$res.Item("ok")            # -> True
$res.Item("entityName")    # -> "NVIDIA CORP"
$res.Item("filings10Q").Count
# then in the standalone workbook, re-trigger Input!A1 = "NVDA" and diff the
# Filings sheet against a pre-change capture to confirm no regression.
```

## Status / next steps

- [x] `modHttp` / `modJsonUtil` / `modPrices` copied here verbatim (diff-verified identical).
- [x] `modSECData.bas` — US headless XBRL reader, ASCII-clean, self-contained. 10 Private
      helpers verified line-identical to `modSEC.bas`. **Not yet compiled/run in Excel.**
- [x] `modMOPSData.bas` — TW (MOPS) headless reader: `GetTwFinancials` + `GetTwMonthlyRevenue`.
      21 Private helpers verified line-identical to `modMOPS.bas` (CJK comments Anglicised).
      **Not yet compiled/run in Excel.**
- [x] `scripts/sync-shared-vba.ps1` — injects the 5 modules into both workbooks via
      `CodeModule.AddFromString` (UTF-8, never VBE Import). Has `-DryRun`; lints the two new
      modules for non-ASCII.
- [x] RR4 side: `RR4/CompanyResearchSEC.bas` renderer + `RunDeepDive` entry point;
      `RR4/SheetCompanyResearch_Code.txt` (`Worksheet_Change` on B34/D34). **Not yet run.**
- [x] `CLAUDE.md` directory map + `RR4/company-research-sec-merge-plan.md` §0 updated.
- [x] **Smoke tested live (2026-09-03, in-memory, nothing saved):**
      - `modSECData.GetXbrlFinancials("NVDA",4,8)` -> ok, NVIDIA CORP, 8 quarters, correct
        FY/FP labels, Revenue/NI/EPS internally consistent.
      - `modMOPSData.GetTwFinancials("2395",6)` -> ok, Advantech, GM 37.7% / OM 19.6% match.
      - `modMOPSData.GetTwMonthlyRevenue("2395",15)` -> ok, recompute-from-cumulative + YoY OK.
      - `CompanyResearchSEC.RunDeepDive` -> renders both US and TW tables + derived block +
        monthly-revenue block into the "Company research" sheet (tested against a THROWAWAY
        COPY of the RR4 workbook; the live file was untouched).
      - Fixed one bug found in testing: `DsoRow`'s non-short-circuit `And` hit `CDbl("")` on
        TW Q4 columns (blank Revenue). Fixed.
      - Known / pre-existing (matches the standalone tools, NOT introduced here):
        NVDA CapEx blank (single-tag `PaymentsToAcquirePropertyPlantAndEquipment` misses
        NVDA's filings); NVDA FY2025Q1 EPS 5.98 is the pre-split as-reported value;
        TW mid-year-quarter CFO/CapEx blank (MOPS tags cash-flow only on the YTD context).
- [ ] **YOU: put it into the real RR4 workbook.** Either (a) close RR4 and run
      `scripts/sync-shared-vba.ps1` + import `CompanyResearchSEC.bas` + paste
      `SheetCompanyResearch_Code.txt` into the sheet, then `Debug > Compile`; or (b) let me
      inject into your open instance and you `Ctrl+S`. Then type `NVDA` / `2395` into B34.
- [ ] **Convergence** — `CONVERGENCE.md`. Collapse `modSEC.bas` / `modMOPS.bas` onto the
      *Data modules so there is one source of truth. Deferred: it mutates the working
      standalone tool and needs a COM regression check. The current duplicated state is
      safe (no name clash — the *Data modules' shared names are all Private).
