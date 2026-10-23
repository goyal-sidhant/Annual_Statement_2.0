### Category 9: Power Query refresh

This category is Step 4's first half: for each ticked client, the program makes a copy of the ITC and Sales report (the "refreshed copy"), opens it in Excel, makes Excel re-run the report's Power Query (Excel's built-in data-import tool; "refresh" means re-running its queries so the report picks up the client's latest files) and decides whether the refresh worked. Both projects do this through COM (Windows' way for one program to drive another; here Python drives Excel), which needs pywin32 (the add-on Python package that provides COM).

**How the examples in categories 9 to 16 were run.** Every "RAN" below means: a copy of each project's code at its snapshot was run on the same made-up inputs (invented clients `ClientA-DL`, `ClientB-MH`, invented figures and paths) with `python -B` in `/root/merge-scratch/run/C5b/` (scripts `scripts/p9.py`, `p11a.py`, `p11b.py`, `p12.py`, `p13.py`, `p14.py`, `p15.sh`, `p16.py`; outputs in `out/`). This Linux machine has no Excel and no pywin32, so where Excel is involved a stand-in was used: `C5b/fakewin/` holds fake `pythoncom`, `win32com`, `win32gui`, `win32api`, `win32con` modules whose "Excel" reads and writes the made-up workbook with openpyxl (a Python library that reads and writes `.xlsx` files without Excel) and records every call the program makes. The stand-in proves the ORDER of calls, the waits and what the program does with the answers; it cannot prove what real Excel does, so everything Excel itself does stays READ. Windows screens were built off-screen (Tk under `xvfb-run`; Qt with `QT_QPA_PLATFORM=offscreen`).

#### 9.1 Rule by rule

| Rule | A's version | B's version | Relation |
|---|---|---|---|
| Step 4 without pywin32 | A-R64: pywin32 imported inside a guard; each report fails with a clear message, the rest of the app runs. `A:power_query_extractor/core/report_processor.py:20-36, 66-77` | B-R124 (category 15): imported without a guard; the whole window cannot start. `B:power_query_extractor/core/report_processor.py:11-15` | **different** (9.2) |
| Which reports are processed | A-R65: one row per client; the ITC and Sales ticks start TICKED where the report exists; unticked → "Not selected for processing". `A:power_query_extractor/gui/extractor_window.py:693-694, 849-861`; `A:power_query_extractor/core/report_processor.py:86-104` | B-R61 (part) and B-R136 (ticks): one row per Version folder; every tick starts UNTICKED; unticked → the same "Not selected for processing". `B:power_query_extractor/core/report_processor.py:50-68`; `B:gui/views/extraction_view.py:291-390` | **different** (9.3) |
| Which file is the "original" report | A-R66: first file the folder listing returns for `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`. `A:power_query_extractor/core/report_processor.py:111-125` | B-R61: the same search and "first file" rule. `B:power_query_extractor/core/report_processor.py:72-89` | **same** - [GST RULE] in B's registry only (9.8) |
| Name of the refreshed copy | A-R67: `<stem><suffix>.xlsx`, `{timestamp}` → `DDMMYY_HHMM`; default `_Refreshed_{timestamp}` typed in 3 places. `A:power_query_extractor/core/report_processor.py:127-130, 162-164` | B-R62: identical code; default typed in 5 places. `B:power_query_extractor/core/report_processor.py:91-94, 118-120` | **same** behaviour (9.9) |
| Recognising a refreshed copy | A-R68: derived from the CONFIGURED suffix, plus the old marker `_refreshed_`. `A:power_query_extractor/core/refresh_naming.py:14-66` | B-R63 (Skip) and B-R136 (status): fixed texts `<stem>_Refreshed_*.xlsx` and `*_Refreshed*`. `B:power_query_extractor/core/report_processor.py:100-101`; `B:gui/pyqt_main_window.py:466-479` | **different** (9.4; status text in 16.12) |
| Skip refresh | A-R69: newest recognised copy reused as success; the tick writes the shared option at once and is never saved. `A:power_query_extractor/core/report_processor.py:132-159`; `A:power_query_extractor/gui/extractor_window.py:483-492` | B-R63: the same reuse logic with B's fixed search; the tick is written into the shared option at each Start and saved in settings. `B:power_query_extractor/core/report_processor.py:96-115`; `B:gui/controllers/extraction_controller.py:53-55` | **different** (9.4) - [GST RULE] by this comparison |
| `auto_refresh_if_missing` | A-R70: never read. `A:power_query_extractor/config/cell_mappings.py:250` | B-R64: never read (same file). | **same** (dead) |
| How the refresh is triggered | A-R71: `Dispatch`, visible maximised Excel, `AppActivate`, keys Alt+A, R, A ⚠, always save. `A:power_query_extractor/core/report_processor.py:209-415` | B-R66, B-R67, B-R69: identical steps. `B:power_query_extractor/core/report_processor.py:165-366` | **same** (9.5) |
| Waits, retries and the wait box | A-R72: same fixed waits; box 5-60, typed values not checked. `A:power_query_extractor/gui/extractor_window.py:269-287, 864` | B-R68 (waits), B-R65 (box): same waits; box 5-120, enforced by Qt; value saved in settings. `B:gui/views/extraction_view.py:126-128` | **different** (box only) (9.6) |
| When a refresh counts as successful | A-R73: no error dialog AND `Info!BB2` ⚠ filled, not 0, not an error text. `A:power_query_extractor/core/report_processor.py:638-685` | B-R70, B-R71: identical test. `B:power_query_extractor/core/report_processor.py:589-636` | **same** - [GST RULE] ⚠ (9.7) |
| "Could not open - continue without refresh" branch | Same unreachable branch in A, `A:power_query_extractor/core/report_processor.py:167-181` (A's registry lists it as dead code, gap A-G74) | B-R72: unreachable. `B:power_query_extractor/core/report_processor.py:122-137` | **same** (dead) |
| Step 4 options (wait, suffix, Skip) | A-R72 (box), A-R119 (category 14: never stored) | B-R65: box 5-120, values read at Start and stored in `settings.json` | **different** (9.6 and 14.6) |

#### 9.2 Step 4 on a PC without pywin32 (A-R64 against B-R124)

- **A.** The five pywin32 modules are imported inside a guard (`A:power_query_extractor/core/report_processor.py:20-36`). Without them, a warning is logged and every report gets the status "pywin32 is not installed for this Python, so Power Query cannot be refreshed. Install it with: pip install pywin32" (`:66-77`); Steps 1 to 3 still work.
- **B.** The same five modules are imported without a guard (`B:power_query_extractor/core/report_processor.py:11-15`), and that module is loaded when the main window is built (`B:gui/controllers/extraction_controller.py:10`), so nothing at all opens.
- **Example (RAN, `p9.py nostub` and `p15.sh` case 2B).** The same made-up Version folder with an ITC and a Sales report, on this PC without pywin32:

| | A | B |
|---|---|---|
| Loading the refresh module | works; `WIN32_AVAILABLE = False`; log "pywin32 not available (No module named 'pythoncom'). Power Query refresh is disabled…" | `ModuleNotFoundError: No module named 'pythoncom'` |
| Processing the client | ITC and Sales both `success False`, error "pywin32 is not installed for this Python, so Power Query cannot be refreshed…"; the reports on disk untouched | not reachable |
| Starting the program (`main.py`) | starts (after the install offer, 15.3) | console: "Missing module: No module named 'pythoncom'" / "Please install required packages:" / "pip install PyQt5 openpyxl"; exit code 1 |

- **What it changes in a client's annual working.** On an office Windows PC without pywin32 neither program gets as far as Step 4: A's start-up check refuses to start (after offering to install it, 15.2 and 15.3), B cannot open its window and its advice ("pip install PyQt5 openpyxl") does not fix the cause. A's guard matters where A does start without a working pywin32: on a non-Windows PC, or on Windows when pywin32 is present on disk but cannot be loaded (A's check only looks for the module file, `A:utils/requirements_check.py:64-70`; A audit RAN XL E2, simulated). There A still sorts the season's GST files and builds the ITC/Sales reports with openpyxl (their Power Query is lost, 12.2) and Step 4 shows every client as failed with a clear message. Technical point only.
- **When it came in.** A's guard: `c463256`, 26-07-2026, "Fix: app failed to start when pywin32 was missing". B's unguarded import is the shared code of `edac6a3` (31-08-2025); it became a start-up requirement in `c258047` (07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"), when the main window began to import the extractor.

#### 9.3 Which reports are sent to the refresh (A-R65 against B-R61 / B-R136)

- **Same.** An unticked report gets `{success: False, error: 'Not selected for processing'}` and no data in both (RAN `p9.py` (a): identical results).
- **Different.** A lists one row per client folder (only its newest Version folder, A-R42 in category 6) with the ticks already ON where the report exists. B lists every Version folder that holds a report, with every tick OFF, and sends one request per ticked Version folder named `<client folder> / <version folder>`.
- **Example (RAN, `p11a.py` / `p11b.py`).** Made-up run folder `Annual Statement-311225 1000` with `ClientA-DL` (two Version folders: `Version-301225 0900` with an ITC report only; `Version-311225 1000` with ITC and Sales) and `ClientB-MH` (one Version folder, ITC report plus an earlier refreshed ITC copy):

| | A (after Scan) | B (after the page fills) |
|---|---|---|
| Rows | `ClientA-DL` (Version-311225 1000), `ClientB-MH` | `ClientA-DL` → `Version-311225 1000`, `Version-301225 0900`; `ClientB-MH` → `Version-311225 1000` |
| Ticks at start | ClientA: ITC ✓ Sales ✓; ClientB: ITC ✓ (Sales cannot be ticked) | all off; "Start" with nothing ticked would be refused |
| Sent after the user's minimum action | A: press Start → 2 requests (ClientB ITC only; ClientA ITC + Sales) | B: press Select All, then Start → 3 requests (`ClientA-DL / Version-311225 1000` ITC + Sales; `ClientA-DL / Version-301225 0900` ITC; `ClientB-MH / Version-311225 1000` ITC) |

- **What it changes.** In B the older Version folder of ClientA is refreshed too if the user presses Select All, which costs another Excel session (about 39 s at the default wait, 9.6) and puts a second ClientA row into every sheet of the consolidated workbook (11.3). In A only the newest Version folder can be refreshed. Technical/presentation point; which Version folder's figures should feed the season's working is the author's call (the choice of folder itself is category 6, A-R42 / B-R39).

#### 9.4 Recognising refreshed copies and "Skip refresh" (A-R68, A-R69 against B-R63) · [GST RULE] (treated as one here)

Neither registry tags these rules. This comparison tags Skip refresh **[GST RULE]**, "when in doubt": it decides which file the figures in the consolidated workbook are read from - a fresh refresh, an older refreshed copy, or the unrefreshed original. Below is only what each project does.

- **A.** The literal pieces of the CONFIGURED suffix around `{timestamp}` are the markers (`_PQ_{timestamp}` → `_pq_`); a file counts as a refreshed copy if, in lower case, it starts with the original's name, ends `.xlsx` and contains all markers, OR contains the old marker `_refreshed_` (`A:power_query_extractor/core/refresh_naming.py:23-66`). Skip refresh reuses the newest such copy (`A:power_query_extractor/core/report_processor.py:132-159`).
- **B.** Skip refresh looks only for `<original name>_Refreshed_*.xlsx`, whatever the suffix box says (`B:power_query_extractor/core/report_processor.py:100-101`).
- **Both.** The reused copy is reported as success "Skipped - Using existing refreshed file" without any check that its own refresh had worked (shared gap A-G2 / B-G7); with no copy found, the report is refreshed normally.
- **Example (RAN, `p9.py` (d2)).** ITC report `ITC_Report_ClientA_DL_311225_1000.xlsx`, Skip refresh ticked, refresh itself replaced by a recorder. The probe sorted the folder listing by name, as a Windows NTFS disk lists it (on this Linux machine's own order the copy came first in S2 and S5 and both projects then took the copy as the "original" - the shared copy-of-a-copy defect A-G15 / B-G8, section 7):

| Case | Version folder holds (made-up) | Suffix box | A | B |
|---|---|---|---|---|
| S1 | original + `…_1000_Refreshed_311225_1100.xlsx` | `_Refreshed_{timestamp}` (default) | reuses the `_Refreshed_` copy, no Excel | same |
| S2 | original + `…_1000_PQ_311225_1100.xlsx` | `_PQ_{timestamp}` | reuses the `_PQ_` copy, no Excel | does not see it; refreshes again into `…_1000_PQ_<now>.xlsx` |
| S3 | original only | `_R{timestamp}` | treats the ORIGINAL report itself as a refreshed copy (`_r` is inside `_report_`), reuses it and reports success; nothing is refreshed | refreshes into `…_1000_R<now>.xlsx` |
| S4 | original + `…_1000_Refreshed_311225_1100.xlsx` | `_PQ_{timestamp}` | reuses the old `_Refreshed_` copy (old marker) | reuses the same copy (fixed search) |
| S5 | original + `…_1000_Refreshed.xlsx` | `_Refreshed` (no `{timestamp}`) | reuses it | does not see it (no `_` after "Refreshed"); refreshes again and overwrites `…_1000_Refreshed.xlsx` |

- **What it changes in a client's annual working.** Suppose the staff refreshed 40 clients' ITC reports last evening with the suffix `_PQ_{timestamp}` and today tick Skip refresh only to rebuild the consolidated workbook. A reuses yesterday's copies in seconds; B refreshes all 40 again (about 40 × 39 s ≈ 26 minutes of Excel at the default wait) and the figures then come from today's refresh. With the suffix `_R{timestamp}`, A would put the figures of the never-refreshed Step 3 reports under a ✓ in the consolidated workbook; B would refresh. Which behaviour is wanted is the author's decision; the over-match itself (`_R`, `_` or a client name containing "Refreshed") is recorded as part of gap A-G16.
- **The tick itself.** A writes the tick into the shared option the moment it is clicked (RAN `p11a.py`: option `True` right after the tick; `A:power_query_extractor/gui/extractor_window.py:483-492`) and never saves it (A-R119). B writes it at every Start (`B:gui/controllers/extraction_controller.py:53-55`) and saves it as `pq_skip_refresh` (B-R116).
- **When each version came in.** Shared: Skip refresh with the fixed `<stem>_Refreshed_*.xlsx` search, `bd176bf` (30-10-2025, "Enhance report processing UI with configurable wait time and file suffix options; add ITC/Sales report selection functionality"); B kept it unchanged. A: `2aaa2b1` (25-07-2026, "Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one"); its message gives the reason: with a customised suffix "'skip refresh' found nothing and re-refreshed every report instead of reusing the existing files - the opposite of what the option is for", and adds that files made under the old default are still accepted "so changing the setting does not make the whole existing history read as 'Never refreshed'". B's write of the option at each Start: `c258047` (07-02-2026).

#### 9.5 How the refresh is done (A-R71 against B-R66, B-R67, B-R69) - same

The two refresh functions differ only by an unreachable `return` that B deleted and by the kind of error a few `except` lines catch (`git diff A/main B/main -- power_query_extractor/core/report_processor.py`, READ). RAN with the fake Excel (`p9.py` (e)): the recorded call sequence is identical in both - `CoInitialize`, `Dispatch("Excel.Application")`, `Visible = True`, `DisplayAlerts = False`, `WindowState = -4137` (maximised), open the copy, `Dispatch("WScript.Shell")`, `AppActivate`, keys `%a`, `r`, `a` (Alt+A, R, A = Data → Refresh All in an English Excel ⚠: ribbon key tips depend on the Excel version and display language), `Save`, `Close(SaveChanges=True)`, `Quit`, `CoUninitialize`; then a second, hidden Excel for reading. With a made-up error dialog on screen, both clicked its "OK" button and both still saved the workbook. Nothing in either project uses `RefreshAll`, `BackgroundQuery = False` or `CalculateUntilAsyncQueriesDone` (READ). Shared since `edac6a3`.

#### 9.6 Waits and the "wait time" box (A-R72 against B-R68, B-R65)

- **Same waits.** RAN with the fake Excel and a counting `sleep` (`p9.py` (e)): 39 s per refreshed report at the default wait of 10 s and 139 s at 60 s, in both (= 19 + w + max(10, w) seconds, before Excel's own start-up and opening time).
- **Different box.**

| | A | B |
|---|---|---|
| Range of the arrows | 5 to 60, hint "seconds (Min: 5, Max: 60)" (`A:power_query_extractor/gui/extractor_window.py:270-287`) | 5 to 120 (`B:gui/views/extraction_view.py:126-128`) |
| Typed `0` | accepted: 0 (RAN `p11a.py`) | Qt keeps 5 (RAN `p11b.py`, `setValue(0)` → 5) |
| Typed `200` | accepted: 200 | kept at 120 |
| Typed `12.7` | read as 12 | not possible (whole numbers only) |
| Typed `abc` or nothing | `TclError` when Start reads it - the click does nothing (A audit RAN PQ gui_probe) | not possible |
| Remembered next time | no, back to 10 (A-R119) | yes, `pq_wait_time` (B-R116) |

- **What it changes.** By the formula above: in A a typed 0 makes each report wait 29 s in all, a typed 200 makes it 419 s; B stays between 34 s (5) and 259 s (120). For a firm refreshing 2 reports × 150 clients at the end of the annual return season, the default 39 s per report is about 3¼ hours of Excel driving in either project. Technical point. History: the 5-60 box is the shared `bd176bf` (30-10-2025); B's 5-120 is `c258047` (07-02-2026).

#### 9.7 When a refresh counts as successful (A-R73 against B-R70, B-R71) · [GST RULE] · ⚠ - same

- **What both do.** Success = no error dialog was found AND a sheet named `info` (any letter case) exists AND its cell `BB2` ⚠ (a template cell position that can change with the template) is not empty, not blank, not `0` / `0.0`, and not one of `#REF!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#ERROR!`. The function is identical in both (`A:power_query_extractor/core/report_processor.py:638-685` = `B:power_query_extractor/core/report_processor.py:589-636`).
- **Worked example through both (RAN, `p9.py` (f), identical results):**

| `Info!BB2` holds (made-up) | Result in A and B |
|---|---|
| empty | fail - "Cell BB2 is empty" |
| `''` or spaces | fail - "Cell BB2 is blank" |
| `0`, `0.0`, `'0'`, `'0.0'` | fail - "Cell BB2 contains 0" |
| `'00'` | pass - "Validated - Client Name: 00" |
| `#N/A`, `#n/a`, `#VALUE!` | fail - "Cell BB2 contains error: …" |
| `ClientA` (the client being processed) | pass |
| `ClientB` (another client's name) | pass - BB2 is never compared with the client |
| `12.5`, `-2146826246`, `True`, `False` | pass |
| no sheet named info | fail - "Info sheet not found" |

  In both, a failed test still saves the workbook and still reads every figure (RAN `p9.py` (e): BB2 empty → status failed, J15 and AA1 read), and the consolidated workbook then shows "Missing" for every figure of that report (11.4).
- **When it came in.** Same code in both since `edac6a3` (31-08-2025, "Initial commit"); neither project changed it. Nothing in either repository writes BB2; the code's comment and log call it the client name (READ). Whether `Info!BB2` is still the cell the current templates fill is for the author to confirm ⚠.

#### 9.8 The registry difference on "which file is the original" (A-R66 / B-R61)

Same code in both (RAN `p9.py` (c): same listing, same original taken). B's registry tags it [GST RULE] ("when in doubt: which file the figures come from"); A's registry does not. Following the "when in doubt" rule, this comparison treats it as a [GST RULE]; there is nothing to choose between the projects. Shared since `edac6a3` (31-08-2025, "Initial commit"); the ITC/Sales choice since `bd176bf` (30-10-2025). On this Linux machine the search listed an older refreshed copy before the original in some folders, and both projects then refreshed that copy (shared gap A-G15 / B-G8; on a Windows NTFS disk the original normally lists first, MEDIUM).

#### 9.9 Copy name (A-R67 / B-R62) - same

RAN `p9.py` (b), identical in both: `_Refreshed_{timestamp}` → `…_Refreshed_<DDMMYY_HHMM>.xlsx`; `_PQ_{timestamp}` → `…_PQ_<…>.xlsx`; `{timestamp}` → the stamp glued to the name; `_Refreshed` → the same name every run; `_Refreshed_{TIMESTAMP}` (capitals) → the text `{TIMESTAMP}` stays in the name; an empty suffix → "… are the same file" and the report fails. Only the number of places where the default text is typed differs (A 3: `A:power_query_extractor/core/report_processor.py:51`, `A:power_query_extractor/gui/extractor_window.py:297`, `A:power_query_extractor/core/refresh_naming.py:14`; B 5: `B:power_query_extractor/core/report_processor.py:28`, `B:models/app_state.py:22`, `B:gui/controllers/extraction_controller.py:24, 139`, `B:gui/views/extraction_view.py:136`) - a code-tidiness point, no behaviour difference.

### Category 10: Extraction cell mappings

After the refresh, a second, hidden Excel opens the refreshed copy and reads fixed cells. The list of cells is the extraction mapping (a table in the code saying "take this cell of this sheet of this report and put it in this column of this sheet of the consolidated workbook"). It lives in `power_query_extractor/config/cell_mappings.py`, which is **byte-for-byte the same file in both projects** (RAN: `git rev-parse` gives blob `598495f` for both). Every row is a **[GST RULE]**, and every return name, table number (GSTR-9 Tables 6J, 8D), sheet name and cell address is **⚠** (it follows the template version of the time).

#### 10.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| "GSTR-3B Difference" ← ITC `Diff GSTR-3B` ⚠ J15, K15, L15 | A-R74 | B-R73 | **same** [GST RULE] ⚠ |
| "GSTR-1 Difference" ← Sales `Diff - GSTR-1` ⚠ N1, O1, P1, Q1 | A-R75 | B-R74 | **same** [GST RULE] ⚠ |
| "Turnover GSTR-1" ← Sales `Diff - GSTR-1` ⚠ I1 | A-R76 | B-R75 | **same** [GST RULE] ⚠ |
| "Reco Diff" ← ITC `Reco` ⚠ C37, D37, E37 | A-R77 | B-R76 | **same** [GST RULE] ⚠ |
| "Table 6J Match" ← ITC `T6 of R9` ⚠ K29 | A-R78 | B-R77 | **same** [GST RULE] ⚠ |
| "Table 8D Limits" ← ITC `T8 of R9` ⚠ G6 | A-R79 | B-R78 | **same** [GST RULE] ⚠ |
| "IMS Reco" ← ITC `IMS_Reco` ⚠ AA1, AD1, AG1 (shows J15/K15/L15 instead, 10.3) | A-R80 | B-R79 | **same** [GST RULE] ⚠ |
| "Extra In Purchase" ← ITC `ExtraInPurchase` ⚠ O1, P1, Q1 (same mismatch) | A-R81 | B-R80 | **same** [GST RULE] ⚠ |
| "Extra In GSTR2B" ← ITC `ExtraInGSTR2B` ⚠ S1, T1, U1 (same mismatch) | A-R82 | B-R81 | **same** [GST RULE] ⚠ |
| How mappings become cell reads (exact, case-sensitive sheet names; each value stored twice) | A-R83 (`A:power_query_extractor/core/report_processor.py:527-636`) | B-R82 (`B:power_query_extractor/core/report_processor.py:478-568`) | **same** [GST RULE] (only the kind of error an `except` catches differs) |
| Consolidated look-up: first field CONTAINING the column title, else the exact field | A-R88 (category 11, `A:power_query_extractor/core/data_consolidator.py:137-150`) | B-R83 (`B:power_query_extractor/core/data_consolidator.py:125-138`) | **same** [GST RULE] (10.3) |
| `FIELD_DISPLAY_NAMES` read only by dead code | A-R84 | B-R85 | **same** (dead) |
| Calculated columns (subtract / add / multiply / divide; 'Error', 'Div/0', 'Unknown'), unused | A-R85 | B-R84 | **same** [GST RULE] (dormant); B narrowed one bare `except:` to `except (TypeError, ValueError)` in `c258047`, same results |

There is no difference in this category, so no choice between the projects arises. The decisions the author still has to take are about the mapping itself (which cell is right), not about A or B.

#### 10.2 Worked example: the 22 cells through both (RAN, `p9.py` (g))

A made-up ITC report with a different invented number in every mapped cell (J15 = 101, K15 = 102, L15 = 103, C37 = 201 …, AA1 = 501, AD1 = 502, AG1 = 503, O1 = 601 …, S1 = 701 …, K29 = "Yes", G6 = 401.5) and a made-up Sales report (N1 = 801 … Q1 = 804, I1 = 900000.25) were read by each project's own reading code through the fake Excel. The 22 values read, the stored keys (each value twice: under the field name, e.g. `IMS_Reco_IGST`, and under the bare cell, e.g. `AA1`) and the whole consolidated workbook came out **identical** in A and B. With the sheet renamed `diff gstr-3b` (lower case), both read `None` for J15/K15/L15 (sheet names are matched exactly) while AA1 was still read as 501; and a result with J15 = `None` put AA1 = 501 into the "IMS Reco" IGST column in both (10.3, row "None").

#### 10.3 The 9-column mismatch, in both (A-R80 to A-R82, A-R88 / B-R79 to B-R81, B-R83) · [GST RULE] · ⚠

Technical description only. In both projects the consolidated workbook fills the IGST, CGST and SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" from `Diff GSTR-3B` J15, K15, L15 whenever those cells were read with a value, because the one-word titles "IGST", "CGST", "SGST" are found inside the field names `GSTR-3B_Difference_Diff_IGST` etc. first. RAN (same results in both):

| `Diff GSTR-3B`!J15 read as | "GSTR-3B Difference" Diff IGST shows | "IMS Reco" IGST shows | "Extra In Purchase" IGST shows |
|---|---|---|---|
| 101 | 101 | **101** (configured AA1 = 501 not shown) | **101** (configured O1 = 601 not shown) |
| `None` (empty cell) | Missing | 501 (AA1) | 601 (O1) |
| `''` (empty text) | Missing | **Missing** | **Missing** |
| 0 | 0 | **0** | **0** |

For a client's GSTR-9 working this means: the three IMS Reco tax-head figures and the six "extra" figures in the consolidated workbook are, today, the GSTR-3B difference figures in both projects. Which cell each column should hold is the author's call. History: the look-up is from `edac6a3` (31-08-2025); the affected one-word columns came with `47a1492` (30-10-2025, "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets"), the last shared commit, so the mismatch is in A and B alike (gap A-G1 / B-G1).

**When each mapping came in.** "GSTR-3B Difference", "GSTR-1 Difference" and "Turnover GSTR-1": `edac6a3` (31-08-2025). "Reco Diff", "Table 6J Match", "Table 8D Limits", "IMS Reco", "Extra In Purchase", "Extra In GSTR2B": `47a1492` (30-10-2025). Neither project changed the file after the split.

### Category 11: Consolidated output layout

The consolidated workbook (`PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, one workbook per Step 4 run with a "Summary" sheet and one sheet per output group) is built by `DataConsolidator.create_report`, which is the same in both projects apart from B's deletion of a leftover header write that was overwritten anyway (`git diff … -- power_query_extractor/core/data_consolidator.py`, READ; RAN `p9.py` (g): identical workbooks from identical results). The differences come from what each project SENDS to it and what it does around it. B's registry also files the Step 3 summary workbook (`GST_Processing_Summary_…xlsx`, B-R94, B-R95) here; A files it under category 8 (A-R61), so it is compared here as well.

#### 11.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| File name, folder, when written, errors | A-R86: always written; into the run folder chosen by the LAST Step 4 scan; an error loses the results and shows an error box. `A:power_query_extractor/gui/extractor_window.py:941-966` | B-R86: written only if at least one result; into the run folder shown when Start was pressed; an error becomes one log line and the Extraction_Log is still written. `B:gui/controllers/extraction_controller.py:86-101` | **different** (11.2) |
| Sheet layout; column A label | A-R87: one row per client, column A = client folder name | B-R87: one row per Version folder, column A = `<client folder> / <version folder>` | **different** (11.3) |
| Which extracted value goes into a column | A-R88 | B-R83 (category 10) | **same** [GST RULE] (10.3) |
| Failed refresh → every figure "Missing" | A-R89 | B-R88 | **same** [GST RULE] (11.4) |
| `#,##0.00` for non-zero numbers (display only) | A-R90 | B-R89 | **same** [GST RULE] (11.4) |
| Summary sheet | A-R91: "Total Clients" = clients sent | B-R90: "Total Clients" = Version folders sent | **same code, different count** (11.3) |
| Description lines (dormant); leftover header write | A-R92 (leftover still there, overwritten) | B-R91 (leftover deleted in `c258047`) | **same** output |
| `show_timestamp` never read | A-R93 | B-R92 | **same** (dead) |
| `_write_section`, `_create_details_sheet` never called | A-R94, A-R95 | B-R93 | **same** (dead) |
| Step 3 summary workbook: when written, failure feedback | A-R61 (category 8): when at least one client was processed; failure → log line "❌ Failed to create summary report". `A:gui/handlers/processing_handler.py:455-475` | B-R94: when at least one client succeeded; failure → nothing on screen. `B:gui/controllers/processing_controller.py:183-193` | **different** (11.5) |
| Step 3 summary workbook: content | A-R61: five sheets; "Variations" lists the unrecognised files | B-R95: same five sheets; "Variations" always "No file variations found" | **different** (11.5) |

#### 11.2 Where and when the consolidated workbook is written (A-R86 against B-R86)

- **Example (RAN, `p11a.py` / `p11b.py`, refresh replaced by the same canned results with invented figures):**

| Situation | A | B |
|---|---|---|
| Normal run | `target1/Annual Statement-311225 1000/PQ_Extraction_Report_<now>.xlsx`, then the question "Processing complete! Open the consolidated report?" | the same workbook in the same folder, plus `Extraction_Log_<now>.txt` beside it; no question |
| The user browses Step 4 to a second target and scans it while the run is going | the workbook goes to the SECOND target's run folder (`target2/Annual Statement-010126 0900/…`) although it holds target1's clients | the workbook stays in `target1/Annual Statement-311225 1000` (the page shows the new folder, the run keeps its own) |
| A value the workbook writer cannot store (here a made-up two-cell result) | error box "Processing failed: Cannot convert (1, 2) to Excel"; no workbook; the figures exist only in the log | log line "Report error: Cannot convert (1, 2) to Excel"; no workbook; `Extraction_Log_<now>.txt` written |
| Stop pressed before the first Version folder | (A's Step 4 has no Stop) | "Extraction stopped by user", "Extraction complete: 0/3 clients processed"; no workbook; Extraction_Log written |

- **What it changes.** During the annual return season a staff member who checks another client batch on Step 4 while A is refreshing would find the consolidated workbook filed under the other batch's run folder; in B it stays with its own batch. When one cell cannot be written, B at least leaves a text log of every figure in the run folder; A leaves the figures only in the program log. Technical points.
- **When it came in.** A: the folder of the last scan since `edac6a3` (`47a1492:power_query_extractor/gui/extractor_window.py:808`, same line in the shared Tk extractor). B: `c258047` (07-02-2026) wrote it into the target folder itself and only "if results"; `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking") moved it back into the run folder and added the Extraction_Log.

#### 11.3 One row per client or per Version folder; "Total Clients" (A-R87, A-R91 against B-R87, B-R90)

- **Example (RAN, same tree as 9.3, ticks as each project starts plus Select All in B):**

| | A | B |
|---|---|---|
| Rows of "Turnover GSTR-1" | `ClientB-MH` ✗ Missing; `ClientA-DL` ✓ 250000 | `ClientA-DL / Version-311225 1000` ✓ 250000; `ClientA-DL / Version-301225 0900` ✗ Missing; `ClientB-MH / Version-311225 1000` ✗ Missing |
| Summary "Total Clients" | 2 | 3 |
| "ITC Reports Processed" / "Sales Reports Processed" | 2/2 and 1/2 | 3/3 and 1/3 |
| Row order | as the file system lists the client folders (no sorting; ClientB came first here) | client folders by name; Version folders newest name first |

- **What it changes.** A client's annual working from B's workbook can contain two or more rows for the same client (one per Version folder ticked), labelled with the folder names; A's has one row per client, labelled with the client folder name (e.g. `ClientA-DL`). B's "Total Clients" counts Version folders, not clients. Layout point for the author; the cells and figures in each row are the same rules (category 10).
- **When it came in.** The consolidator is shared (`edac6a3`). A's per-client list is the shared Tk extractor's; B's per-version tree and labels: `b0e7d0e` (09-02-2026).

#### 11.4 "Missing" and number display (A-R89, A-R90 against B-R88, B-R89) · [GST RULE] - same

RAN `p9.py` (g), identical in both: a report whose refresh status is not success shows "Missing" in every column although its figures were read (made-up ClientB-MH row); an empty value shows "Missing"; 0 stays 0; non-zero numbers get the display format `#,##0.00` (two decimals, thousands separators) while the stored value is not rounded; text stays text. B's audit also RAN that a text starting with `=` is stored as a live formula; the same line is in A (READ, identical code). Shared since `edac6a3` (31-08-2025, "Initial commit"); neither project changed these lines after the split.

#### 11.5 The Step 3 summary workbook (A-R61 against B-R94, B-R95)

- **Same.** Five sheets (Summary, Client Status, File Mapping, Errors, Variations), same layout; the writing code is identical (MAP: `ExcelHandler.create_summary_report` IDENTICAL).
- **Different.** What is passed in and what happens on failure. A passes the scan's real list of problem files (`A:gui/handlers/processing_handler.py:224`); B passes an empty list (`B:gui/controllers/processing_controller.py:73`).
- **Example (RAN, `p16.py`: a real Step 3 run on the same made-up source folder, openpyxl path, with one unrecognised file `Purchase register ClientB.xlsx`):**

| Sheet "Variations" | A | B |
|---|---|---|
| Row 2 | `Purchase register ClientB.xlsx` · "No matching pattern found" · "Check file naming convention" · 4859 · "Rename file" | "No file variations found" |

  The "Summary" sheets were identical (Total Clients 2, Successful 2, Total Files 8, Reports Generated 4, Processing Mode fresh).
- **What it changes.** After organising a season's files, A's summary workbook tells the staff which input file was not picked up for which reason; B's says there were none although the scan found one (B shows only a count on screen). The hint text in that sheet is itself a [GST RULE] (A-R12 / B-R12, category 1; MAP 2.1 item 2), so this difference also decides whether the hint is ever seen.
- **When it came in.** The workbook: `edac6a3`. B's empty list: `c258047` (07-02-2026); at the split the real list was passed (`47a1492:gui/handlers/processing_handler.py:161`, READ). A summary that fails: A logs "❌ Failed to create summary report"; B shows nothing because `create_summary_report` catches its own errors and B's worker skips only the success line (B's audit RAN, XL E4).

### Category 12: Excel handling and safety

This category covers how each project starts, uses and closes Excel, and what happens to files when something goes wrong. The author's Standard 12 ("Excel through COM", `rebuild_pipeline_FINAL.md:844-878`) is the yardstick in Section 8; here only A and B are set side by side. Both projects share almost all of this code; the one real difference is where Step 3 builds a report while Excel works on it.

#### 12.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Excel or openpyxl for Step 3 reports | A-R96: Excel only on Windows with pywin32, else openpyxl (Power Query dropped); reachable on any PC without pywin32. `A:core/excel_handler.py:19-28, 89-102` | B-R97: the same condition and fallback, but unreachable through the app (no window without pywin32). `B:core/excel_handler.py:19-28, 83-93` | **different** (reachability) (12.2) |
| Step 3 Excel session | A-R97: `GetActiveObject` (the user's Excel) else `Dispatch`; hide; save twice; always `Quit`. `A:core/excel_handler.py:104-260` | B-R96, B-R110: identical. `B:core/excel_handler.py:95-217` | **same** (12.4) |
| Where the report is built | A-R98: on a local copy in `%TEMP%\gst_excel_temp`, copied to the Version folder once at the end. `A:core/excel_handler.py:135-155, 212-215, 254-260` | B-R52 (category 8), B-R105: template copied straight to the report's final name and edited there; `gst_excel_temp` created and never used. `B:core/excel_handler.py:52-53, 122-129` | **different** (12.3) |
| A report counts as created if the file exists | A-R99 | B-R59 (category 8) | **same** |
| "Already open" check by full path | A-R100 (shared helper `A:utils/excel_com.py:19-29` since `ba2a753`) | B-R98, B-R99 (the same loop typed twice) | **same** behaviour |
| Step 3 open with one 8.3 short-path retry | A-R101 | B-R98 | **same** |
| Step 4 open with long-path prefix and fallbacks | A-R102 | B-R99 | **same** |
| Step 4 Excel sessions (two `Dispatch` per report; hidden reading session) | A-R103 | B-R101, B-R66, B-R110 | **same** (RAN 9.5) |
| 218 characters treated as Excel's path limit ⚠ | A-R104 | B-R100 | **same** |
| Error-dialog search and clicking ⚠ (window words and classes assume an English Windows/Excel) | A-R105 | B-R102, B-R103 | **same** (RAN 9.5: "OK" clicked in both) |
| openpyxl load with `keep_links=True`, macros not kept | A-R106 | B-R97 | **same** |
| `EXCEL_SAFETY` stored, never read | A-R107 | B-R106, B-R107, B-R108, B-R109 | **same** (dead) |
| All Python warnings switched off | A-R108 | B-R104 | **same** (RAN `p12.py`: first warning filter `('ignore', None, Warning, None, 0)` in both) |
| `CoInitialize` / `CoUninitialize` inside each Excel function | part of A-R97 and A-R103 | B-R110 | **same** |

#### 12.2 The openpyxl fallback: reachable in A, not in B (A-R96 against B-R97)

- RAN `p12.py`: on this PC `USE_WIN32COM` is `False` in both, and the same openpyxl path builds the report in both. In A that path is what a user actually gets on a PC where pywin32 cannot be loaded (Windows refuses to start A without pywin32, 15.2, so in normal office use it is a non-Windows or broken-pywin32 case). In B the path cannot be reached because B's window does not open without pywin32 (RAN `p15.sh` 2B).
- **A side effect found while running (category 8 rule, wording correction).** On this path a Links cell that sits INSIDE a merged range but is not its top-left cell (made-up template with `A2:B2` merged, mapped cell B2): A skips that cell silently and reports success (B2 empty, the other cells written); B fails the whole report ("Fallback method error: 'MergedCell' object attribute 'value' is read-only") and leaves the template under the report's name. When the merge STARTS at the mapped cell (`B2:C2`), both write it and succeed (RAN, openpyxl of both virtual environments). This sharpens MAP row A-R58 / B-R53 and B-R97 ("a merged target cell makes the report fail").
- **When it came in.** Fallback: `edac6a3`. B's narrower error handling that turns the merged cell into a failure: `c258047` (07-02-2026); its message lists "Fixed all 26 bare except: clauses with specific exception types".

#### 12.3 Where Step 3 builds a report (A-R98 against B-R52, B-R105)

- **A.** Copies the template to a local staging file (a temporary local copy that is worked on and then copied to its final place) `%TEMP%\gst_excel_temp\<report name>`, lets Excel work there, copies the finished file to the Version folder once and deletes the staging file.
- **B.** Copies the template straight to `…\Version-…\ITC_Report_<client>_<code>_<DDMMYY_HHMM>.xlsx` and lets Excel open and save that file in place; it also creates `gst_excel_temp` at every run and never uses it.
- **Example (RAN, `p12.py`).** Made-up ITC template; each project's own Step 3 code; Excel = the fake Excel in the COM cases:

| Case | A | B |
|---|---|---|
| Normal (fake Excel) | returns True; report at the final name with B2 = the GSTR-3B folder text; staging folder empty afterwards | returns True; the same report; `gst_excel_temp` exists, empty |
| Excel cannot open the file (made-up error "-2147352567 … Open method of Workbooks class failed" on every attempt) | returns False; **no file** at the report name | returns False; **the unfilled template sits under the report name** (B2 empty, template text in B3) |
| openpyxl path, template that is not a workbook | returns False; no file at the report name | returns False; a byte-for-byte copy of the bad template under the report name |

- **What it changes.** In B a failed report leaves a file that looks like a finished ITC report (right name, right folder) but has empty Links cells; Step 4 later lists it as a report to refresh, and its refresh would read the template's own text. In A the Version folder simply has no report for that client, which the log names as failed. On a network share A also writes each report in one sequential copy instead of Excel's many small reads and writes (A's code comment, `A:core/excel_handler.py:135-146`, which also states that the report's data sources are absolute paths in the Links sheet, so the copy is byte-identical). Technical point (file safety).
- **When it came in.** A: `12d505f` (23-07-2026, "Build Excel reports on local disk, copy to network once"); at the split the template was copied straight to the output path (`core/excel_handler.py:135@47a1492`), which is what B still does. The `gst_excel_temp` folder itself: `edac6a3`.

#### 12.4 The shared Excel session (A-R97, A-R103 against B-R96, B-R101, B-R110) - same

RAN with the fake Excel (`p12.py`): identical call sequences in both. With an Excel "already open by the user" (the fake answered `GetActiveObject`), both used it, set `Visible = False`, `DisplayAlerts = False`, `ScreenUpdating = False` on it, saved, set alerts and screen updating back on, and then called `Quit` on the USER'S Excel ("excel.Quit USER-OWN-EXCEL"); `Visible` was never set back. Without a running Excel both started one with `Dispatch` and quit it. Each report waited 1.0 s of fixed pauses (12.1 / category 13). What real Excel then does with the user's own unsaved workbooks is not testable here (MEDIUM: a hidden "save changes?" prompt or lost work). Same in both since `edac6a3`; it is a shared gap (taking over, hiding and quitting the user's own Excel: A-G54 / B-G63; clean-up that can skip `Quit`: A-G28 / B-G38; the comparison with Standard 12 is Section 8), not an A-or-B choice.

### Category 13: Performance and operational constants

#### 13.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| How the source folder is read in Step 1 | A-R109: one directory listing, its cached details reused. `A:utils/helpers.py:309-365` | B-R1, B-R3 (category 1): three wildcard searches, then a size query and a read per file. `B:utils/helpers.py:249-266`; `B:core/file_parser.py:171-215` | **different** (13.2) |
| Signature reads in parallel | A-R110: up to 16 at once from 4 files on. `A:utils/helpers.py:267-272, 346-361` | — (one after another) | **only A** (13.2) |
| Run folder worked out once per run | A-R111. `A:core/file_organizer.py:45-47, 150-161` | — (recomputed for every client, `B:core/file_organizer.py:146-179`) | **only A** (13.3) |
| Step 3 fixed pauses (0.5 s after copy, 0.5 s after close) | A-R112. `A:core/excel_handler.py:155, 210` | B-R114. `B:core/excel_handler.py:125, 180` | **same** (RAN `p12.py`: 1.0 s per report in both) |
| Setup checklist re-check 350 ms after typing | A-R113. `A:gui/tabs/setup_tab.py:506-534` | — (no checklist) | **only A** (16.5) |
| Background jobs | A-R135, A-R143: Step 1 scan on the window's own thread; Step 3 and Step 4 in daemon threads (killed instantly when the program ends); Step 4 can be started twice | B-R111: scan, processing and extraction each in its own Qt thread; a second job of the same kind refused silently | **different** (13.4) |
| Stop | A-R135: Step 3 only, between clients | B-R112: Step 3 between clients and Step 4 between Version folders | **different** (13.4) |
| Progress and time-left display | A-R141 (category 16): bar split into copying and reports, "📊 Progress: X.X%" | B-R113: bar moves between clients only; ends "Complete \| Xm Ys"; scan progress per file | **different display, same tracker** (13.5) |

#### 13.2 Reading the source folder (A-R109, A-R110 against B-R1, B-R3)

- **Example (RAN, `p13.py`).** A made-up source folder with 24 recognisable files of 4 clients. Every file-system call on that folder was counted, and each was given an invented 20 ms delay as a stand-in for an office network share (the real delay of the share is not known here):

| | A | B |
|---|---|---|
| Directory listings | 1 | 3 |
| Size/type queries (`stat`) | 2 | 26 |
| File opens (signature reads) | 24, in a pool of reader threads | 24, one after another |
| Time with 20 ms per call | 0.11 s | 1.08 s |
| Clients found | 4 | 4 |

- **What it changes.** For a firm whose GST files for the annual return season sit on a share reached over a VPN, A's scan is the faster one; A's commit message (`4f84ba9`) reports "40.78s wall clock for 2508 files" on the real share, against about 11 minutes before. B's scan runs in a background thread, so its window stays usable while it scans (13.4); A's window freezes until the scan ends. Technical point.
- **When it came in.** A: `a603c02` (23-07-2026, "Speed up network file scanning (behavior-preserving)"), `56aaad9` (25-07-2026, "Reuse directory-listing metadata: 81% fewer round-trips per file"), `4f84ba9` (25-07-2026, "Read Excel signatures concurrently: 11 min -> 41s on the real share"). B: the shared three-search scan of `edac6a3`, kept.

#### 13.3 Run folder once per run (A-R111)

- **Example (RAN, `p13.py`).** Mode "Re-run" (which reuses the newest existing `Annual Statement-…` folder), two clients; between client 1 and client 2 a newer run folder `Annual Statement-311225 1015` appears on the share (for example a colleague's run):

| | A | B |
|---|---|---|
| ClientA | `Annual Statement-311225 1000\ClientA-DL` | `Annual Statement-311225 1000\ClientA-DL` |
| ClientB | `Annual Statement-311225 1000\ClientB-DL` | **`Annual Statement-311225 1015\ClientB-DL`** |

- **What it changes.** In B one run can split its clients over two run folders, so its summary workbook (written into the first client's run folder) and Step 4 (which shows one run folder) miss the clients that went elsewhere; A keeps the whole run in one folder and lists the target folder once instead of once per client. Technical point. History: A `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor"); B: the shared per-client look-up of `edac6a3`.

#### 13.4 Background jobs and Stop (A-R135, A-R143 against B-R111, B-R112)

- RAN (C5b, Qt off-screen): two scans started at once in B → the second was refused with only the log line "Scan already in progress"; one result came back. RAN `p11b.py` run 4: Stop before the first Version folder → "Extraction stopped by user", no consolidated workbook, Extraction_Log written. A: the scan runs on the window's thread (`A:gui/handlers/file_handler.py:71-72`, READ), Step 3 and Step 4 run in daemon threads (`A:gui/handlers/processing_handler.py:129-134`; `A:power_query_extractor/gui/extractor_window.py:878-883`), Step 4 has no Stop and its Start button is switched back on by any Step 4 scan, even during a run (A audit RAN, PQ gui_probe).
- **What it changes.** In B a staff member can stop a long Step 4 run between Version folders; in A the only way is to close the program, which kills the daemon thread in the middle of a client (Excel may be left open). Technical point. History: A's Step 3 re-entrancy guard `33d0baf` (23-07-2026); B's thread pattern `c258047` (07-02-2026).

#### 13.5 Progress display (A-R141 against B-R113)

The time-left tracker and its time texts are the same code (MAP: `ProgressTracker.*` IDENTICAL; RAN `p13.py`: `format_duration` gives "0s", "59s", "1m 1s", "59m 59s", "1h 1m", "25h 1m" in both). Only the display differs: A moves the bar inside each client (first half copying, second half reports) and shows "📊 Progress: X.X%" and "Client i of n: <client> | ETA: …"; B moves the bar only between clients, shows "Client i/N: <client> | ETA: …" and ends with "Complete | Xm Ys". Presentation only.

### Category 14: Settings, cache, logs and run history

#### 14.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Where settings live, file format | A-R114: one file `~/.gst_organizer_cache.json`, one line of JSON (a plain-text settings format). `A:gui/utils/cache_manager.py:14-48` | B-R115, B-R116: folder `~/.gst_organizer/` with `settings.json` (indented), `session_state.json`, `run_history.json`. `B:models/cache_manager.py:16-62` | **different** (14.2) |
| Which keys | A-R115: `source_folder`, `itc_template`, `sales_template`, `target_folder`, `processing_mode`, `workflow_mode`, `include_client_name`, `dark_mode`, `client_name_max_length`, `recent_folders`. `A:gui/main_window.py:358-389` | B-R116: 11 keys incl. `include_client_name_in_folders`, `pq_wait_time`, `pq_suffix_pattern`, `pq_skip_refresh`. `B:models/app_state.py:10-46` | **different** (14.2) |
| When settings are saved | A-R116: after a Browse, the workflow radio, the length arrows, the global tick, dark mode - NOT a typed path, NOT on closing | B-R116: at every scan, after processing, dark mode, File > Save Settings, on closing - NOT at Start | **different** (14.3) |
| Recent folders (never read) | A-R117: `recent_folders` in the settings file | B-R117: `recent_source_folders` in `session_state.json` (plus `active_tab`, `window_size`, also never read) | **different place, both unused** (14.3) |
| One-time import of A's settings file | — | B-R118: imports `~/.gst_organizer_cache.json` once | **only B** (14.4) |
| How the standalone extractor finds its folder | A-R118: home settings file first, then three local files. `A:power_query_extractor/extractor_main.py:36-63` | B-R128 (category 15): the old entry point cannot even be imported | **different** (14.5) |
| Step 4 options stored? | A-R119: no (10 s, `_Refreshed_{timestamp}`, Skip off at every launch) | B-R65, B-R116: yes | **different** (14.6) |
| Main log file `gst_organizer.log` in the current folder | A-R120. `A:main.py:20-27` | B-R121. `B:main.py:15-23` | **same** (RAN `p15.sh`: created in the folder the program was started from, in both) |
| Extractor log file | A-R121: standalone extractor writes `pq_extractor.log` | B-R127, B-R128: B's launcher sets up no log file | **different** (14.5) |
| What is logged | A-R122: client names, Links values (first 50 characters), every extracted figure; `_log` always "info" | B-R121, B-R123: the same lines and the same eight text cut-offs (50 / last 50 / 100 / 200 characters; READ, same lines in both files) | **same** |
| Run history and extraction log | A-R123: none | B-R119, B-R120, B-R122: `run_history.json`, `Extraction_Log_<YYYYMMDD_HHMM>.txt` per Step 4 run | **only B** (14.7) |

#### 14.2 One settings file or a settings folder (A-R114, A-R115 against B-R115, B-R116)

RAN `p14.py`: A writes one line of JSON to `~/.gst_organizer_cache.json`; B writes `~/.gst_organizer/settings.json` (12 lines, indented, UTF-8) plus `session_state.json`. Both rewrite the whole file in place (a crash in the middle of a write leaves a broken file; both then start with defaults). Neither follows Standard 13's `%LOCALAPPDATA%\[App Name]\` (Section 8). History: A's file is the shared `edac6a3` design; B's folder `c258047` (07-02-2026).

#### 14.3 The same settings file through both programs (A-R114 to A-R117 against B-R116 to B-R118)

- **Example (RAN, `p14.py`, real windows, fresh scratch home).** The file A writes, with made-up paths: source `/srv/share/GST/Inputs-FY2425`, mode `resume`, workflow `full`, include-name ON, a typed max length `0`, recent folders `[…FY2425, …FY2324]`. Then the user types a new source `/srv/share/GST/Inputs-FY2526` into the box and closes the window.

| Step | A | B |
|---|---|---|
| Window opens: max length box shows | 35 (A replaces 0 or less by 35) | 15 (B's import copies 0; Qt's box keeps 15-100) |
| Files after opening | unchanged | `.gst_organizer/settings.json` written (`client_name_max_length` 0, `include_client_name_in_folders` true, the Step 4 keys at their defaults), `session_state.json` with only `…FY2425` (the old `recent_folders` list is NOT carried over), marker `.migrated` |
| After typing a new source and closing | settings file unchanged - the typed path is lost | `settings.json`: source `…FY2526`, max length **15**; `session_state.json`: `[…FY2526, …FY2425]`, `active_tab` 0, `window_size` [1400, 900] |
| A save as made by Browse (A) | `source_folder` `…FY2526`, max length 35, `recent_folders` `[…FY2526, …FY2425, …FY2324]` | — |

- **What it changes.** A staff member moving from A to B keeps the four paths and the include-name tick, loses the workflow choice and the recent list, and a typed 0 becomes 15 characters in B (35 in A) - so client folder names are cut at a different length (category 6). In A, a path typed (not browsed) before closing is forgotten at the next start. Technical points.
- **When it came in.** A's save points: shared Tk code plus `510f7b4` (26-07-2026, `workflow_mode`). B's settings, close-time save and import: `c258047` (07-02-2026).

#### 14.4 B's import of A's settings file (B-R118) - only B

Shown in 14.3. The import is the only place where B reads the file A writes today; it runs once (marker `.migrated`) and leaves the old file in place (RAN). A has nothing to import.

#### 14.5 Standalone extractor and its log (A-R118, A-R121 against B-R127, B-R128)

RAN `p15.sh`: A's `launch_extractor.py` (case 4A) found the target in `~/.gst_organizer_cache.json` ("Found target folder in cache: /srv/share/GST/Out-FY2425"), opened the extractor alone and wrote `pq_extractor.log` in the current folder. B's `launch_extractor.py` (case 4B) printed "Failed to launch: No module named 'pythoncom'" (with pywin32 it opens the full window on Step 4 and sets up no log file at all, READ `B:launch_extractor.py:1-28`). B's old `power_query_extractor/extractor_main.py` (case 5B) cannot be imported: `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'` (deleted by `a62d3cc`, 07-02-2026). A: home settings file first since `c1583d6` (23-07-2026, "Fix extractor cache path and version-folder selection"). Technical point.

#### 14.6 Step 4 options remembered or not (A-R119 against B-R65, B-R116)

RAN: B's second start (`p14.py`) showed the suffix `_PQ_{timestamp}` typed in the first session; A's panel starts at `_Refreshed_{timestamp}`, wait 10 and Skip off every time (`A:power_query_extractor/gui/extractor_window.py:270, 297, 315`; RAN `p11a.py` default 10). Combined with 9.4: in A, copies made yesterday with a custom suffix are, after a restart, recognised only if they also contain `_refreshed_` (RAN C5b `p145`: with the default suffix in the box A showed "Never refreshed" for a `_PQ_` copy). History: B `c258047` (07-02-2026).

#### 14.7 Run history and the extraction log (A-R123 against B-R119, B-R120, B-R122) - only B

- RAN (`p14.py`, B's own `RunHistory`): a processing record is kept under the scan key `ClientA-DL` (it feeds Step 2's "Last Run" column); the Step 4 result for the same client is kept under `ClientA-DL / Version-311225 1000`, a separate record that no screen reads, with `pq_refreshed: true` although the made-up refresh had FAILED; after 61 runs, 50 run entries were kept and `total_runs` was 61. RAN `p11b.py`: `Extraction_Log_<YYYYMMDD_HHMM>.txt` begins "PQ Extraction Log - …", 60 "=", then every Step 4 screen line.
- **What it changes.** B keeps a permanent record in the user profile of every extracted figure, path and BB2 text for every client (client data stored outside the client folders, relevant to confidentiality); A keeps no history, only the workbooks and the log files. Technical/confidentiality point. History: B `c258047` (07-02-2026) and `b0e7d0e` (09-02-2026, Extraction_Log).

### Category 15: Start-up and dependencies

#### 15.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Which packages are required and checked | A-R124: openpyxl essential; pywin32 essential on Windows only; checked at start with a report. `A:utils/requirements_check.py:17, 37-84`; `A:main.py:68-106` | B-R124: nothing checked; pywin32 needed on every system just to open the window | **different** (15.2) |
| Offer to install missing packages and restart | A-R125. `A:utils/dependency_installer.py:26-158` | — | **only A** (15.3) |
| `requirements.txt` ⚠ | A-R126: `openpyxl>=3.1.2`, `pywin32>=305; sys_platform == 'win32'`, `lxml>=4.9.0`, `pytest>=7.0.0` | B-R124: `PyQt5>=5.15.0`, `openpyxl>=3.1.2`, `lxml>=4.9.0`; pywin32 only as the comment `# pywin32>=306` | **different** (15.4) |
| Console handling | A-R127: console output always re-wrapped as UTF-8 (no console → crash); banner and report; waits for Enter on exit paths. `A:main.py:17, 45-73` | B-R125: re-wrapped only if a console exists; prints only on failure. `B:main.py:11-13, 32-59` | **different** (15.5) |
| Preparing the program folder | A-R128: program folder first on the search path; creates `core`, `utils`, `gui`, `logs`, `temp` and placeholder `__init__.py` files. `A:main.py:29-43` | B-R125, B-R127: search path only; nothing created | **different** (15.5) |
| Step 4 loading guard card | A-R129. `A:gui/tabs/extract_tab.py:21-30, 84-105` | — | **only A** |
| Unused version labels | A-R130: `3.0.0` in core, gui, utils; `1.0.0` in power_query_extractor | B-R129 (category 16): `4.0.0` in gui, `3.0.0` in core and utils, `1.0.0` in power_query_extractor (RAN grep) | **different** (labels only) |
| Window start-up order; Step 4 filled from disk at start | — (A builds Step 4 empty; no scan at start-up, A-R143, A-R144) | B-R126: fills Step 4 by reading the saved target folder on the window's thread. `B:gui/pyqt_main_window.py:35-68, 174-190` | **only B** (16.13) |
| Launcher | A's `launch_extractor.py` opens the extractor alone (A-R128, A-F29) | B-R127: opens the full window on Step 4; no log set-up | **different** (14.5) |
| Old extractor entry point | A's works (A-R118, A-R121) | B-R128: cannot be imported | **different** (14.5) |

#### 15.2 Requirements check (A-R124 against B-R124)

- **Example (RAN, A's own check on this Linux PC and with Windows simulated by loading the check while `sys.platform = 'win32'`; B by starting `main.py`, `p15.sh` 2B):**

| | A on Linux | A on Windows (simulated) | B |
|---|---|---|---|
| pywin32 missing | "[not found] pywin32 …" - optional, the program continues | "[MISSING] pywin32 …" - essential: error, wait for Enter, exit 1 (after the install offer) | "Missing module: No module named 'pythoncom'" / "pip install PyQt5 openpyxl", exit 1 |

- **What it changes.** On an office Windows PC without pywin32, A refuses to start and says exactly which package is missing and why ("Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them."), so no report can silently lose its Power Query; B also refuses to start, with advice that does not name pywin32. Technical point. History: A `f363946` (26-07-2026, "Check dependencies at startup and report them clearly") and `be6418c` (26-07-2026, "Require pywin32 on Windows so reports cannot lose Power Query"); B `c258047`.

#### 15.3 Offer to install (A-R125) - only A

RAN: A's command is `[<this python>, -m, pip, install, -r, <program folder>/requirements.txt]` with a 600-second limit; the WHOLE file is installed (pytest included). On this Linux PC the offer appears at every start because pywin32 stays "missing" (A audit RAN START; `p15.sh` 2A stopped at it). History: `a5ad487` (26-07-2026, "Offer to install missing dependencies, then restart").

#### 15.4 `requirements.txt` ⚠ (A-R126 against B-R124)

READ (both files). Minimum versions only in both; package versions change over time ⚠. A lists pywin32 as a real requirement on Windows (`064c634`, 23-07-2026) and pytest (`6c1e1c6`, 23-07-2026); B lists PyQt5 and leaves pywin32 as a comment (`c258047`). Standard 13 asks for exact versions (Section 8).

#### 15.5 Console and program folder (A-R127, A-R128 against B-R125, B-R127)

- **Example (RAN, `p15.sh`, fresh copies of each program, standard output closed as under `pythonw.exe`, the Windows Python that runs without a console window):**

| | A | B |
|---|---|---|
| Start with no console | dies at once: `AttributeError: 'NoneType' object has no attribute 'buffer'` (`A:main.py:17`), before any log file | passes the console step (`B:main.py:12` checks first); on this PC stops later at the missing pywin32, writing "Import Error: No module named 'pythoncom'" to `gst_organizer.log` |
| Folders created in the program folder on a normal start | `logs`, `temp` (never used) | none |

- **What it changes.** If the office later starts the tool from a desktop shortcut without a console (Standard 13's launcher uses `pythonw.exe`), A cannot start at all and B starts but shows its start-up errors nowhere. A on a read-only network program folder fails the first time `logs`/`temp` do not exist yet (A audit, MEDIUM). Technical point. History: A's unconditional wrapper is the shared `edac6a3` line; B added the check in `c258047`.

### Category 16: User-interface behaviour rules

This category holds the screens' fixed values: window, tabs, buttons, tables, messages and help texts. Two items here are GST rules: the Dry Run gate (what must be chosen before a preview of the reports, 16.4) and the file names shown to the user as the expected input names (16.11). Everything else is presentation or technical behaviour.

#### 16.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Main window | A-R131: "GST File Organizer v3.0", 1200×800, minimum 1000×700, tabs with emoji. `A:gui/main_window.py:57-77` | B-R129: "GST File Organizer v4.0", 1400×900, minimum 1200×800, File and View menus. `B:gui/pyqt_main_window.py:70-118` | **different** (16.2) |
| Automatic tab switches | A-R132 | B-R140 (with coloured tab dots) | **different** (16.3) |
| Setup checklist | A-R133. `A:gui/utils/setup_validation.py:13-115` | — | **only A** (16.5) |
| Scan and Dry Run gates | A-R134 (Dry Run uses A-R51's template gate) | B-R133 (Dry Run needs only a ticked client; also the processing and extraction gates) | **different** (16.4) - Dry Run gate [GST RULE] per MAP 2.1 item 7 |
| Run control (what is greyed out) | A-R135: Start greyed, Stop active; Step 2/3 buttons react to the mouse even when greyed, not to the keyboard | B-R138: Start/Stop swap; everything else stays active; a disabled Stop looks enabled | **different** (16.6) |
| Long client-name warnings | A-R136 | — (B's help note still promises one, B-R141) | **only A** |
| Per-client include-name toggle on Step 2 | A-R137 | — (B-R36, category 6: nothing sets it) | **only A** |
| Step 2 client table | A-R138 | B-R134 | **different** (16.7) |
| Dry-run text | A-R139 | B-R135 | **different** (16.4) |
| Step 3 log and completion | A-R140 | B-R137 (HTML logs), B-R139 (feedback) | **different** (16.8) |
| Progress and time left | A-R141 | B-R113 (category 13) | **different display** (13.5) |
| Export of ticked clients | A-R142 | — (B's "EXPORT LIST" button is connected to nothing, B-F10) | **only A** |
| Step 4 panel behaviour | A-R143: no Stop; question "Open the consolidated report?" | B-R138, B-R139: Stop; no end message, status "Extraction complete" | **different** (11.2, 13.4) |
| Step 4 folder | A-R144: follows Step 1's target as typed; no scan at start | B-R39 (category 6), B-R126: the last SAVED target, read from disk at start-up, after processing, after extraction, on Refresh List | **different** (16.13) |
| "Last refresh" status text | A-R145 | B-R136 | **different** (16.12) |
| Mouse-wheel rules | A-R146 | — (Qt's own scrolling) | **only A** |
| Tooltips (help text on hover) | A-R147 (9 places, RAN grep) | — (0 places, RAN grep `setToolTip`) | **only A** |
| Dark mode | A-R148: records and restores each widget's colours | B-R144: one style sheet; Steps 1-2 stay light | **different** (16.9) |
| GUI colour table | A-R149: used | B-R130: never read; B-R132: B's own colour table and `#info` style never used | **different use** |
| `GUI_CONFIG['fonts']` | A-R150 (dead) | B-R130 (dead) | **same** (dead) |
| `GUI_CONFIG['colors']['white']` | A-R151 (dead) | B-R130 (dead) | **same** (dead) |
| Message texts in use | A-R152: four `MESSAGES` texts used | B-R131: none used (texts written inline) | **different** (16.4) |
| Other message texts | A-R153 (dead) | B-R131 (dead) | **same** (dead) |
| Workflow choice ("organise only" / "full pipeline") | A-R154 | — | **only A** (16.10) |
| Title strip and status bar | A-R155: blue strip, Dark Mode tick, "v3.0 \| Production Ready" | B-R129: plain status bar "Ready" | **different** (16.2) |
| Version label | A-R156: "v3.0" | B-R129: "v4.0" | **different** (16.2) |
| Shortened paths in the status bar (`truncate_path`) | A-R157: used after Browse (`A:gui/handlers/file_handler.py:24, 54`) | B-R143: the same function, never called | **different use** (RAN C5b `p166`: the same made-up 112-character network path gives the same "...te Limited-DL\Version-311225 1000\Other ITC related files" in both - on Linux a Windows-style path counts as one part, so only "..." + the last 57 characters remain; only A shows it) |
| Help texts with the expected file names ⚠ | A-R16 (category 1), A-R147: the card reads the patterns' own example names | B-R141: hand-typed card | **different** - [GST RULE] ⚠ (16.11) |
| Size and time displays | A-R31 (sizes), A-R145 (`%Y-%m-%d %H:%M`) | B-R142 (`dd/mm/yy HH:MM`; history `YYYY-MM-DD HH:MM:SS`) | **different** (time formats; `format_size` identical) |
| Theme leftovers | — | B-R132 (dead) | **only B** (dead) |

#### 16.2 Window, title and version (A-R131, A-R155, A-R156 against B-R129)

RAN `p16.py`: A - title "GST File Organizer v3.0", size 1200×800, minimum 1000×700, tabs "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing", "🔄 Step 4: Extract". B - title "GST File Organizer v4.0", 1400×900, minimum 1200×800, tabs "Step 1: Setup", "Step 2: Validation", "Step 3: Processing", "Step 4: Extract && Refresh" (shown as "&"), menus File and View. Neither label tells the user which version is newer: A ("v3.0") has 34 commits of its own after the split, B ("v4.0") 5. Presentation only.

#### 16.3 Automatic tab switches (A-R132 against B-R140)

RAN `p16.py`: after a scan both switch to Step 2; Dry Run switches both to Step 3; after a real Step 3 run A (workflow "organise only") stays on Step 3 and logs "ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.", while B always moves to Step 4 (index 3) and marks Step 3 green and Step 4 orange. Presentation only.

#### 16.4 Gates, messages and the dry run (A-R134, A-R139, A-R152 against B-R133, B-R135, B-R131)

- **Example (RAN, `p16.py`, same made-up source folder: ClientA with all six files, ClientB with two):**

| Action | A | B |
|---|---|---|
| Scan with an empty source box | error box "Error": "Please select a source folder containing GST files" | warning box: "Please select a source folder first." |
| Scan with a folder that does not exist | error box: "Source folder does not exist" | warning box: "Folder not found: <path>" |
| Scan with a FILE as source | passes the gate, then error box "Scan Error": "Not a directory: <path>" | warning box: "Folder not found: <path>" |
| Dry Run, no templates chosen | error box: "Please select both ITC and Sales template files"; no preview | preview printed, ending "Would create: ITC Report" / "Would create: Sales Report" |
| Start, no templates chosen | the same error box as Dry Run (READ: the same gate function, `A:gui/handlers/processing_handler.py:78`) | warning box: "Please select a valid ITC template." |
| Dry Run text | "[HH:MM:SS]" before each line; "🎯 Mode: fresh", "👥 Selected: 2 clients"; indented lines | no time, no mode, no count; leading spaces lost (HTML log, 16.8) |
| Step 4 Start with nothing ticked (READ) | warning "No Selection": "Please select at least one report to process" (`A:power_query_extractor/gui/extractor_window.py:859-861`); (Step 4 Scan with an empty folder box: "No Folder" / "Please select a folder first", `:613-615`) | warning: "No versions selected for extraction."; no run folder shown: "No Level 1 folder available."; ticked Version folders gone: "Could not find version folders for the selected items. Make sure clients have been processed first (Step 3)." (`B:gui/pyqt_main_window.py:506-538`) |

- **[GST RULE] for the Dry Run gate only** (MAP tags A-R51 / B-R48 / B-R133 as a GST rule, because it decides what must be chosen before a preview of the reports is shown). What each does is in the table; no recommendation. History: the gate is in the shared Tk code (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`); B's PyQt Dry Run without it came with `c258047` (07-02-2026). The rest of this sub-section is presentation only.
- **Correction (A's audit, A-R152).** A's audit says the four `MESSAGES` texts are "shown as `Warning` boxes before Dry Run and Start". The code shows "Please select a source folder containing GST files" (before a SCAN), "Please select both ITC and Sales template files" and "Please select a target folder for organized files" in ERROR boxes titled "Error" (`A:gui/handlers/file_handler.py:123, 135, 139`); only "Please select at least one client to process" is a "Warning" box (`A:gui/handlers/processing_handler.py:45, 83`). RAN above.

#### 16.5 Setup checklist and its re-check (A-R133, A-R113) - only A

RAN (A's `evaluate_setup` on made-up paths): source = a FILE, ITC template = a FOLDER named `tpl_folder.xlsx`, Sales template missing, target blank → source ✓, ITC ✓, Sales ✕ "File not found", target ○ "Not selected yet"; Scan allowed, Step 3 not; pill "2 of 4 items set"; still needed for Step 3: Sales template, Target folder. (The check is existence only - a file passes as a folder and a folder as a template.) B shows a fixed 7-line instruction card and checks paths only when SCAN or START is pressed (16.4). History: `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"); its message: "Each required input reports its own state as you pick it, and the buttons enable themselves - no more clicking Scan to discover something is missing".

#### 16.6 Run control (A-R135 against B-R138)

- **A.** Greys Start and enables Stop during Step 3, but its Step 2 and Step 3 buttons are bound to a mouse press instead of a button command (`A:gui/tabs/processing_tab.py:74, 87`; `A:gui/tabs/validation_tab.py:118, 130, 142, 190`). RAN (C5b `p166`, real A window under xvfb): with Start greyed as during a run, a mouse press on it still called the start handler; with Start enabled, the Space key and the button's own `invoke()` called nothing. (A's re-entrancy guard, `33d0baf`, then ignores a second Start while a run is active.)
- **B.** Swaps Start and Stop and leaves everything else active. RAN (C5b `p166`, real B window off-screen, run state switched on): Step 3 Start disabled, Stop enabled; Step 2's START, DRY RUN, Select All, Clear All, Complete Only and EXPORT LIST all still enabled. Its red Stop style has no disabled look, so a disabled Stop looks enabled (B audit RAN MW e16; `B:gui/theme/theme_manager.py:104-111`).
- **What it changes.** In A a keyboard user cannot press Start (RAN) or the Step 2 buttons (READ: wired the same way); in B a second click on Step 2's START during a run is silently refused (13.4). Technical point.

#### 16.7 Step 2 client table (A-R138 against B-R134)

- **Example (RAN, `p16.py`):**

| | A | B |
|---|---|---|
| Columns | ✓, 🏢 Client Name, 🌏 State, 📊 Status, 📁 Files, ❌ Missing, ➕ Extra, 📁 Name | Select, Client, State, Status, Files, Missing, Extra, Folder Name, Last Run |
| ClientA row | ☐ · ClientA · Delhi · "✅ Complete" · 6 · None · 0 · ☐ No | [x] · ClientA · Delhi · Complete · 6 · 0 · 0 · ClientA-DL · (empty) |
| ClientB row | ☐ · ClientB · Maharashtra · "⚠️ Missing 4 files" · 2 · "Annual Report, GSTR-2B Reco, IMS Reco, Sales Reco" · 0 · ☐ No | [x] · ClientB · Maharashtra · Missing 4 files · 2 · **4** · 0 · ClientB-MH · (empty) |
| Rows ticked after a scan | none (Select All ticks both) | all, complete or not |

- **What it changes.** In A the staff see WHICH return files a client lacks (e.g. the IMS Reco and GSTR-2B Reco files for ClientB) and must tick clients on purpose; in B they see only the number 4 and every client - complete or not - is processed unless unticked. The status logic itself is category 5 ([GST RULE], same in both). History: A's table is the shared Tk code; B's `c258047` / `7e09e5c`.

#### 16.8 Logs and completion feedback (A-R140 against B-R137, B-R139)

- **Example (RAN).** The message `    Links B11 = <ImsReco-ClientA-Delhi-31032025> & done` written to B's Step 3 log comes out as "Links B11 = & done": B inserts each message as HTML (a coloured `<span>`), so the leading spaces and the text between `<` and `>` vanish (`B:gui/views/processing_view.py:142-157`); B's audit RAN that 2,000 lines take about 49 s to show. A writes the same text unchanged with a "[HH:MM:SS]" stamp into a Tk text box (`A:gui/utils/ui_helpers.py:16-33`). After a real Step 3 run (`p16.py`): both showed "Processing Complete" / "Successfully processed 2/2 clients … Files organized in: <target>"; A then opened the run folder (on Linux with `xdg-open "<target>/Annual Statement-<now>"`, on Windows with Explorer); B opens Explorer on Windows only and moves to Step 4.
- **What it changes.** A file name or note containing `<…>` in B's on-screen log is shown incomplete. Technical point.

#### 16.9 Dark mode (A-R148 against B-R144)

READ, with the audits' runs (A START dm_test; B MW e8): A records each widget's colours and restores them; B applies one dark style sheet, while Steps 1 and 2 keep their own light colours and some labels become dark on dark. B's light style also makes three Step 1 Browse buttons white on white (B audit RAN VW t5_pixels). Presentation only. History: A `aeb83cf`, `c89c1d4`, `725fbaf`, `169c7b5` (all 26-07-2026); B `c258047`, `7e09e5c`.

#### 16.10 Workflow choice (A-R154) - only A

A offers "organise only" (default) or "full pipeline"; "full" opens Step 4 after Step 3, fills its folder and scans it, without starting the refresh (`A:gui/handlers/processing_handler.py:136-175`; A audit RAN GUI exp1). B has no choice and behaves like A's "full" every time: after every Step 3 run it rebuilds the Step 4 list from disk (`B:gui/pyqt_main_window.py:384`) and moves to Step 4 (`:401`; RAN 16.3), and, like A, never starts the refresh by itself. History: `510f7b4` (26-07-2026, "Add a workflow choice: organise-only or full pipeline").

#### 16.11 Help texts with the expected file names (A-R16, A-R147 against B-R141) · [GST RULE] · ⚠

Described only; the patterns themselves are category 1.

| | A | B |
|---|---|---|
| Source of the text | each pattern's own example name, read from `FILE_PATTERNS` (`A:gui/tabs/setup_tab.py:477-503`) | a hand-typed list (`B:gui/views/setup_view.py:488-499`) |
| IMS Reco line ⚠ | "ImsReco-ClientName-State-DDMMYYYY.xlsx" | "ImsReco-Client-State-DDMMYYYY" |
| Mode lines | tooltips on the radio buttons (`A:gui/tabs/setup_tab.py:71-133`) | "Re-Run: Overwrite existing", "Resume: Skip already processed" (contradicting the mode buttons, B-R44) |
| Promise of a long-name warning and per-client settings | warning exists (A-R136) | note "Names >10 chars will show warning. This overrides individual client settings if checked." - neither exists in B |

Same six types in the same order in both; A's lines include `.xlsx`, B's do not; neither mentions `.xls`, the ban on hyphens inside a client or state name, or the "(N)" forms (READ). **When it came in.** A: `322717f` (25-07-2026); its code comment says the card is "read from FILE_PATTERNS so this card can never drift from the regexes the parser actually uses". B: `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections"). Which wording the staff should see is the author's call.

#### 16.12 "Last refresh" status on Step 4 (A-R145 against B-R136)

- **Example (RAN, `p11a.py`, `p11b.py` and C5b `p145`):**

| Version folder holds (made-up) | Suffix box | A shows | B shows |
|---|---|---|---|
| ClientA: ITC + Sales reports, no copy | default | "Never refreshed" | "Pending" (orange) |
| ClientB: ITC report + `…_Refreshed_311225_1100.xlsx` | default | "ITC: 2025-12-31 04:00 \| Sales: Never" | "Refreshed (31/12/25 04:00)" (green) |
| ClientA: ITC report + `…_PQ_311225_1100.xlsx` | `_Refreshed_{timestamp}` | "Never refreshed" | "Pending" |
| the same | `_PQ_{timestamp}` | "ITC: 2025-12-31 04:00 \| Sales: Never" | "Pending" |

- In both the time is the copy's date modified, so a copy left by a FAILED refresh also shows as refreshed (shared gap A-G2 / B-G7). A shows one time per report type; B one status per Version folder, where one refreshed ITC copy marks the whole folder.
- **What it changes.** A staff member checking which clients' reports are already refreshed for the GSTR-9 working sees per-report information in A and a single flag per Version folder in B; with a custom suffix that does not contain "_Refreshed", B never shows "Refreshed". Technical point. History: shared `bd176bf` (per report type, fixed `ITC_Report_*_Refreshed_*.xlsx`, `47a1492:power_query_extractor/gui/extractor_window.py:645, 652`); A `2aaa2b1` (configured suffix) and `2acffed` (25-07-2026, recomputed after a run); B `b0e7d0e` (09-02-2026, per Version folder, `*_Refreshed*`).

#### 16.13 Step 4's folder and start-up fill (A-R144 against B-R39, B-R126)

READ, with runs: A's Step 4 box follows Step 1's target as typed until the user puts another folder there, and Step 4 stays empty until Scan is pressed (`A:gui/tabs/extract_tab.py:76-131`; `bfcb24a`, 26-07-2026). B reads the last SAVED target and fills Step 4 from disk when the window opens, after processing, after extraction and on Refresh List, on the window's own thread (`B:gui/pyqt_main_window.py:171-190, 419-504`; RAN `p11b.py` through that function), so a slow or unreachable network target delays B's start-up. Technical point; which run folder is chosen is category 6.
