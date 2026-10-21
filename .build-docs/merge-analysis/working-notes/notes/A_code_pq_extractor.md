# A-R2 — Project A (AnnualStatement_oldv1 @ bf2f1dd, 27-07-2026): Power Query Extractor, its launchers, and the Excel COM helper — code-reader working notes

Reader: A-R2. Files covered (every line read with the Read tool or `cat -n`):
`power_query_extractor/__init__.py`, `power_query_extractor/config/__init__.py`,
`power_query_extractor/config/cell_mappings.py`, `power_query_extractor/core/__init__.py`,
`power_query_extractor/core/data_consolidator.py`, `power_query_extractor/core/refresh_naming.py`,
`power_query_extractor/core/report_processor.py`, `power_query_extractor/extractor_main.py`,
`power_query_extractor/gui/__init__.py`, `power_query_extractor/gui/extractor_window.py`,
`launch_extractor.py`, `utils/excel_com.py`.

Also read (for callers / context only, no full sections): `gui/tabs/extract_tab.py` (all 131 lines),
`gui/utils/cache_manager.py` (all), `gui/main_window.py:100-200, 331-397`, `gui/handlers/processing_handler.py:69-200, 285-300, 350-360, 455-470`,
`core/file_organizer.py:85-200`, `core/excel_handler.py:315-345` + grep of its Dispatch/Quit lines,
`utils/constants.py:75-95`, `utils/helpers.py:25-35, 68-125`, `main.py:15-30`,
`tests/test_extractor_panel.py`, `tests/test_extractor_standalone.py`, `tests/test_optional_pywin32.py:30-101`,
`docs/superpowers/specs/2026-07-25-merge-extractor-design.md` (all 149 lines), and the full diffs of every
A/main commit that touched my files (edac6a3, 01862a0, 4c42758, bd176bf, 47a1492, 275d9a4, c1583d6, 064c634,
ba2a753, c1f425e, 2acffed, a677670, 2aaa2b1, bfcb24a, c463256, bf2f1dd).

All paths are repository-relative; line numbers are those of snapshot bf2f1dd.

---

## 0. How to read these notes

- **READ** = I read the code. **RAN** = I executed it in my scratch copy `/root/merge-scratch/run/A-R2/A`
  (a copy of A-copy without `__pycache__`), with the shared `/root/merge-scratch/venv-A` (Python 3.14.6,
  Tk 9.0, openpyxl 3.1.5; nothing installed), on Linux **without** Excel or pywin32, using
  `python -B`, `PYTHONDONTWRITEBYTECODE=1`, `pytest -p no:cacheprovider`, `xvfb-run -a` for Tk. Probe scripts:
  `/root/merge-scratch/run/A-R2/gui_probe.py`, `scan_probe.py`, `ver_probe.py`, `getattr_probe.py` (plus inline
  `python -c` runs listed in section 5).
- **HIGH / MEDIUM / LOW** = my confidence in an inference.
- **[GST RULE]** = anything deciding a tax figure, which return/table/sheet/cell a figure belongs to, a rate,
  limit, rounding, period, GSTIN/PAN/state logic, ITC eligibility, due date. Described only — never judged.
  **TIME-VARYING** = a value that can change over time (form/table numbers, template cell addresses, etc.).
- **Exact line counts.** Several files have no newline at the end, so `wc -l` shows one line fewer than the
  real count. Real counts (RAN, Python `len(src.splitlines())`): `power_query_extractor/__init__.py` 3,
  `config/__init__.py` 1, `cell_mappings.py` 278, `core/__init__.py` 1, `data_consolidator.py` 432,
  `refresh_naming.py` 66, `report_processor.py` 804, `extractor_main.py` 82, `gui/__init__.py` 1,
  `extractor_window.py` 1063, `launch_extractor.py` 30, `utils/excel_com.py` 29.

### Plain-English glossary (terms used below)

- **COM / "Excel through COM"**: a Windows mechanism that lets Python remote-control the real Microsoft Excel
  program (start it, open a file, press its buttons, read cells, save, quit). Needs Windows, Excel and pywin32.
- **pywin32**: the Python add-on package that provides COM access; its modules are `pythoncom`, `win32com`,
  `win32api`, `win32con`, `win32gui`.
- **Dispatch vs DispatchEx**: two ways to get an Excel. `Dispatch` may attach to an Excel that is already
  running (the user's own); `DispatchEx` always starts a separate, private Excel.
- **Power Query**: Excel's built-in data-import engine. "Refresh" re-runs its queries so the workbook pulls
  in fresh data from the files it is linked to.
- **Tk / Tkinter**: Python's built-in toolkit for windows, buttons, check boxes.
- **Widget**: one on-screen element (button, label, entry box...).
- **Thread / worker thread**: a second line of work running at the same time as the window, so the window
  stays responsive. **Daemon thread**: a worker that is killed instantly, without clean-up, when the program
  exits.
- **Marshal to the UI thread**: hand a screen update to the window's own thread instead of doing it from the
  worker (Tk is not safe to touch from two threads).
- **Callback**: a function handed to other code, to be called later (e.g. "call this to write a log line").
- **glob**: wildcard file-name matching, e.g. `ITC_Report_*.xlsx`.
- **mtime**: a file's or folder's "date modified" time kept by the file system.
- **openpyxl**: Python library that writes and reads `.xlsx` files without Excel.
- **SendKeys**: simulating key presses, as if a person typed them.
- **Window handle (hwnd)**: Windows' internal ID for a window or dialog.
- **Dead code**: code that nothing calls.
- **Module-level code**: code that runs the moment a file is imported (loaded), not when a function is called.
- **Shallow copy**: copying a container but not the things inside it (both copies share the inner items).

---

## 1. Most important findings (one line each; details in the file sections)

1. **Nine columns of the consolidated workbook receive the wrong cell's value (RAN, HIGH).** In
   `PQ_Extraction_Report_*.xlsx`, the `IGST`/`CGST`/`SGST` columns of output sheets "IMS Reco",
   "Extra In Purchase" and "Extra In GSTR2B" are filled from `Diff GSTR-3B!J15/K15/L15`, not from
   `IMS_Reco!AA1/AD1/AG1`, `ExtraInPurchase!O1/P1/Q1`, `ExtraInGSTR2B!S1/T1/U1` as configured (whenever the
   GSTR-3B difference cells were read with a value; only if those came back empty does the configured cell
   win). Cause: a
   substring lookup in `data_consolidator.py:141-145`. Which cell belongs to which column is the author's
   [GST RULE]; the point here is only that the code does not follow its own configuration. Mappings date from
   shared commit 47a1492 (30-10-2025); the lookup from edac6a3. (Same finding as A-X-EXCEL item 7.)
2. **Pressing Scan (or Browse) while a run is going re-enables "Start Processing" (RAN, HIGH)**
   (`extractor_window.py:756-757` inside `display_clients`), so a second worker thread can be started next
   to the first; the re-scan also swaps `self.level1_folder`, so the running job's consolidated report is
   saved into whatever folder was scanned last (RAN: saved into the re-scanned folder).
3. **The refresh is done by key presses, and a failed refresh still leaves a "Refreshed" file that is later
   trusted (READ, HIGH).** `report_processor.py:258-293` sends Alt+A, R, A to whatever window is in front
   (it continues even when it could not bring Excel to the front, 273-274), then waits fixed times. The copy
   made by `shutil.copy2` (164) keeps the ORIGINAL's date-modified time and the "_Refreshed_" name even if
   Excel never opened it; the client list then shows it as refreshed (`extractor_window.py:45-82`) and
   "Skip Refresh" reuses it and reports success (`report_processor.py:137-155`).
4. **Excel handling differs from Standard 12 "Excel through COM" in most points (READ, HIGH on code; MEDIUM on
   real-world effect):** `Dispatch` not `DispatchEx` (218, 536); refresh Excel is made **visible and
   maximised** (225-227); no `RefreshAll` / `BackgroundQuery=False` / `CalculateUntilAsyncQueriesDone` /
   `Refreshing` check; `IgnoreReadOnlyRecommended=True` only in the last-resort open (781);
   `workbook.ReadOnly` never checked; `CoUninitialize()` is called inside the same function while Excel
   objects are still referenced (412, 633). If `Dispatch` attaches to the user's own Excel, that Excel is
   driven by key presses and then `Quit()` with `DisplayAlerts=False` (406, 628).
5. **Figures are thrown away when the refresh "failed" (RAN, HIGH).** If a report's status is not success
   (error dialog seen, Info!BB2 check failed, any exception), every column of that report shows `Missing`
   (`data_consolidator.py:128-129`) even though values were extracted (the GUI log even says "(Data
   extracted despite error)", `extractor_window.py:924-925`). Conversely, a ✓ row can show all `Missing`
   when the refresh succeeded but extraction failed (`report_processor.py:615-617` returns `{}`). The
   workbook never says WHY (the methods that would write error details, `_write_section` and
   `_create_details_sheet`, are dead code).
6. **"Refreshed File Suffix" rules are fragile (RAN):** a custom suffix such as `_R{timestamp}`,
   `_{timestamp}` or `_Report{timestamp}` makes every ORIGINAL report count as "refreshed" (its name already
   contains `_r` / `_` / `_report`), so Skip Refresh extracts from the un-refreshed original and reports
   success. An empty suffix fails the copy (same file); `{TIMESTAMP}` in capitals is not replaced by the
   writer but is understood by the reader. The suffix, wait time and Skip-Refresh choice are **not saved**
   anywhere: each launch starts again at `_Refreshed_{timestamp}`, 10 s, unticked.
7. **Which batch and which version are used is decided by "date modified" (RAN):** the newest
   `Annual Statement-*` folder and, per client, the newest `Version-*` folder by mtime
   (`extractor_window.py:631, 648`). Merely creating a file in an older folder (for example Excel's `~$` lock
   file when someone opens a report there) makes that older folder "newest" (RAN on Linux; NTFS behaves the
   same way for new entries — MEDIUM-HIGH). Pointing the folder box directly at an older
   `Annual Statement-...` folder does NOT work ("No processed folders found", RAN), contrary to the bfcb24a
   commit message and the design spec.
8. **Closing the window during a run kills the worker without clean-up (READ, HIGH).** The worker is a daemon
   thread (`extractor_window.py:882`); neither the standalone window (no close handler at all) nor the main
   window (`gui/main_window.py:344-355` checks only Step 3's flag) asks first, so Excel can be left open with a
   half-refreshed workbook. There is no Stop button in the extractor.
9. **When embedded as Step 4, the panel's status messages are invisible (READ, HIGH).** `status_var` is
   shown only by the panel's own status bar (`extractor_window.py:579-585`), which is switched off when
   embedded (`gui/tabs/extract_tab.py:78`). "Scanning folder...", "No processed folders found",
   "Error scanning folder" therefore appear only in the log on the other sub-tab ("Processing").
10. **Logging: `pq_extractor.log` is created in the current folder at import time (RAN)** and receives every
    client name, file name and extracted figure (`report_processor.py:56` logs every GUI line too). The
    design spec says this `basicConfig` "moves into the standalone entry point" — it did not; Step 4 avoids it
    by importing `extractor_window` directly (`gui/tabs/extract_tab.py:8-10, 25`).
11. **Dead code:** `_write_section` (`data_consolidator.py:250-352`), `_create_details_sheet` (354-414),
    unreachable `return` (`report_processor.py:383-386`), effectively-dead `except` branch
    (`report_processor.py:169-181`), `ReportProcessor.excel_instance` (48), `cleanup()` no-op (801-804),
    `EXTRACTION_OPTIONS['auto_refresh_if_missing']` and `['show_timestamp']` (never read),
    `FIELD_DISPLAY_NAMES` (read only by dead code), the `'description'` branch of `create_report` (53-65, no
    config has it), `_calculate_value` (no mapping has a calculation), `power_query_extractor.__version__`,
    `ExtractorPanel.target_folder` attribute after `__init__`, the client-dict key `'path'` and the result keys
    `'timestamp'`, `'original_file'`, `'refreshed_file'`, unused imports (`os` and `DEFAULT_SUFFIX_PATTERN` in report_processor,
    `Path` in data_consolidator, `sys` in extractor_main), the `workbook` parameter of `_calculate_value`.

---

## 2. File sections

(In the order the assignment lists them.)

### power_query_extractor/__init__.py (3 lines) - [CORE]

- **Purpose:** marks the folder as a Python package (so `power_query_extractor.xxx` can be imported) and
  declares a version string. READ.
- **Imports:** none.
- **Module-level constants and globals:** `__version__ = "1.0.0"` (line 3). NO READER FOUND anywhere in A
  (grep of every .py file). It disagrees with the main app's `'3.0.0'` (`gui/__init__.py:11`,
  `core/__init__.py:12`, `utils/__init__.py:38`) and the "v3.0" strings in `main.py:2`. Not a defect, just
  an unused label. HIGH.
- **Classes/functions:** none.
- **Business rules / edge cases / gaps:** none. No `__all__` and no re-exports: every importer uses the full
  sub-module path (`power_query_extractor.gui.extractor_window`, `power_query_extractor.extractor_main`, ...).
- **Excel touchpoints / files:** none.
- **Comments/docstrings:** line 1 docstring "Power Query Extractor Module for GST File Organizer" — matches.
- **History:** unchanged since edac6a3 (31-08-2025) (`git diff edac6a3 A/main` empty for this file). Identical
  in B/main. READ.

### power_query_extractor/config/__init__.py (1 line, no final newline) - [CONFIG]

- **Purpose:** package marker for `config`; contains only the docstring `"""Configuration for Power Query
  Extractor"""`. READ.
- **Imports / constants / functions:** none. No re-exports (importers use `..config.cell_mappings`).
- **History:** unchanged since edac6a3. Identical in B.

### power_query_extractor/config/cell_mappings.py (278 lines) - [CONFIG]

- **Purpose:** the extractor's configuration: which cell of which sheet in each ITC/Sales report is read in
  Step 4, which output sheet and column it is written to in the consolidated workbook, plus four run options;
  at import it also builds two derived look-up tables used by the rest of the code. READ.
- **Imports:** none (pure Python data + one loop).
- **Module-level constants and globals:**
  - Docstring 1-27: "Cell Mapping Configuration - MASTER VERSION ... This file controls what data is
    extracted and where it goes", with "HOW TO EDIT" instructions (each `EXTRACTION_CONFIG` section creates a
    new sheet; entry format; calculation types None / 'subtract' / 'add' / 'multiply' / 'divide').
  - `EXTRACTION_CONFIG` (33-242): dict of 9 output sheets; each has `'report_type'` (`'ITC'` or `'Sales'`) and
    `'mappings'` (list of dicts with `input_sheet`, `input_cell`, `output_column`, `calculation`). Full table
    below.
  - `EXTRACTION_OPTIONS` (248-253): `'skip_refresh': False` (249) "Set to True to skip Power Query refresh";
    `'auto_refresh_if_missing': True` (250) "Automatically refresh if no refreshed file found";
    `'show_timestamp': False` (251) "Add extraction timestamp to report"; `'missing_data_text': 'Missing'`
    (252) "What to show when data is not found".
  - `CELL_MAPPINGS = {}` (260) and `FIELD_DISPLAY_NAMES = {}` (261), filled by the loop 264-278 under the
    comment "Backward compatibility" / "Build old format for compatibility".
  - Side-effect globals left behind by the module-level loop (264-278): `sheet_name`, `config`,
    `report_type`, `mapping`, `sheet`, `field_name` remain as module attributes holding the LAST iteration's
    values (e.g. `cell_mappings.config` = the "Extra In GSTR2B" entry). Nobody imports them; harmless. HIGH.

- **Every mapping (this is [GST RULE] material — which sheet/cell a figure is taken from and which
  return/table name it is filed under; described, not judged; TIME-VARYING: return/table names such as
  GSTR-3B, GSTR-1, GSTR-2B, IMS, GSTR-9 Tables 6J/8D, and every cell address, which depends on the template
  version):**

  | # | Output sheet in consolidated workbook (EXTRACTION_CONFIG key) | report_type (which report file) | input_sheet (sheet inside the report) | input_cell | output_column (header in output sheet) | Derived field name (RAN) | calculation | Lines |
  |---|---|---|---|---|---|---|---|---|
  | 1 | GSTR-3B Difference | ITC | Diff GSTR-3B | J15 | Diff IGST | GSTR-3B_Difference_Diff_IGST | None | 38-43 |
  | 2 | GSTR-3B Difference | ITC | Diff GSTR-3B | K15 | Diff CGST | GSTR-3B_Difference_Diff_CGST | None | 44-49 |
  | 3 | GSTR-3B Difference | ITC | Diff GSTR-3B | L15 | Diff SGST | GSTR-3B_Difference_Diff_SGST | None | 50-55 |
  | 4 | GSTR-1 Difference | Sales | Diff - GSTR-1 | N1 | Diff Taxable Value | GSTR-1_Difference_Diff_Taxable_Value | None | 64-69 |
  | 5 | GSTR-1 Difference | Sales | Diff - GSTR-1 | O1 | Diff IGST | GSTR-1_Difference_Diff_IGST | None | 70-75 |
  | 6 | GSTR-1 Difference | Sales | Diff - GSTR-1 | P1 | Diff CGST | GSTR-1_Difference_Diff_CGST | None | 76-81 |
  | 7 | GSTR-1 Difference | Sales | Diff - GSTR-1 | Q1 | Diff SGST | GSTR-1_Difference_Diff_SGST | None | 82-87 |
  | 8 | Turnover GSTR-1 | Sales | Diff - GSTR-1 | I1 | Taxable Value | Turnover_GSTR-1_Taxable_Value | None | 95-100 |
  | 9 | Reco Diff | ITC | Reco | C37 | IGST Diff | Reco_Diff_IGST_Diff | None | 109-114 |
  | 10 | Reco Diff | ITC | Reco | D37 | CGST Diff | Reco_Diff_CGST_Diff | None | 115-120 |
  | 11 | Reco Diff | ITC | Reco | E37 | SGST Diff | Reco_Diff_SGST_Diff | None | 121-126 |
  | 12 | Table 6J Match | ITC | T6 of R9 | K29 | Table 6J Match | Table_6J_Match_Table_6J_Match | None | 133-138 |
  | 13 | Table 8D Limits | ITC | T8 of R9 | G6 | Table 8D Limits | Table_8D_Limits_Table_8D_Limits | None | 146-151 |
  | 14 | IMS Reco | ITC | IMS_Reco | AA1 | IGST | IMS_Reco_IGST | None | 159-164 |
  | 15 | IMS Reco | ITC | IMS_Reco | AD1 | CGST | IMS_Reco_CGST | None | 165-170 |
  | 16 | IMS Reco | ITC | IMS_Reco | AG1 | SGST | IMS_Reco_SGST | None | 171-176 |
  | 17 | Extra In Purchase | ITC | ExtraInPurchase | O1 | IGST | Extra_In_Purchase_IGST | None | 184-189 |
  | 18 | Extra In Purchase | ITC | ExtraInPurchase | P1 | CGST | Extra_In_Purchase_CGST | None | 190-195 |
  | 19 | Extra In Purchase | ITC | ExtraInPurchase | Q1 | SGST | Extra_In_Purchase_SGST | None | 196-201 |
  | 20 | Extra In GSTR2B | ITC | ExtraInGSTR2B | S1 | IGST | Extra_In_GSTR2B_IGST | None | 209-214 |
  | 21 | Extra In GSTR2B | ITC | ExtraInGSTR2B | T1 | CGST | Extra_In_GSTR2B_CGST | None | 215-220 |
  | 22 | Extra In GSTR2B | ITC | ExtraInGSTR2B | U1 | SGST | Extra_In_GSTR2B_SGST | None | 221-226 |

  Facts about the table (READ/RAN, no judgement): 17 ITC cells across 7 report sheets, 5 Sales cells on 1
  report sheet ("Diff - GSTR-1" feeds two output sheets: rows 4-7 and row 8). Every mapping is a single cell;
  no `calculation` is used anywhere (all `None`). Tax heads that appear: IGST, CGST, SGST (and "Taxable
  Value"); no other head appears in any mapping. Cell references inside one report type are all different
  (ITC: J15,K15,L15,C37,D37,E37,K29,G6,AA1,AD1,AG1,O1,P1,Q1,S1,T1,U1; Sales: N1,O1,P1,Q1,I1), so the
  "bare cell reference" keys stored by the extractor do not collide today (RAN). The comments above each
  block (34, 60, 91, 104, 129, 142, 155, 180, 205) only label the block (e.g. "# GSTR-3B Differences Sheet",
  "# Turnover Sheet"). Example template commented out at
  229-241. The report sheets named here ("Diff GSTR-3B", "Diff - GSTR-1", "Reco", "T6 of R9", "T8 of R9",
  "IMS_Reco", "ExtraInPurchase", "ExtraInGSTR2B") are not created by any code in A — they must exist in the
  author's ITC/Sales templates (no template is in the repo; READ, `git ls-tree` of A/main shows no .xlsx).

- **Derived structures (RAN, printed after import):**
  - `CELL_MAPPINGS = {'ITC': {'Diff GSTR-3B': {'J15': 'GSTR-3B_Difference_Diff_IGST', 'K15': ..., 'L15': ...},
    'Reco': {'C37','D37','E37'}, 'T6 of R9': {'K29'}, 'T8 of R9': {'G6'}, 'IMS_Reco': {'AA1','AD1','AG1'},
    'ExtraInPurchase': {'O1','P1','Q1'}, 'ExtraInGSTR2B': {'S1','T1','U1'}}, 'Sales': {'Diff - GSTR-1':
    {'N1','O1','P1','Q1','I1'}}}` — shape `{report_type: {input_sheet: {input_cell: field_name}}}` (cells
    shown without their field names for brevity; each maps to the field name in the table above).
  - `FIELD_DISPLAY_NAMES = {field_name: output_column}` (22 entries; e.g. `'IMS_Reco_IGST': 'IGST'`).
  - Field name rule (276): `f"{output_sheet}_{output_column}".replace(' ', '_')`.
- **The build loop 264-278, step by step:** for each output sheet → take its `report_type` → create
  `CELL_MAPPINGS[report_type]` if missing → for each mapping, only if BOTH `input_sheet` and `input_cell`
  are truthy (270) → create `CELL_MAPPINGS[report_type][input_sheet]` if missing → compute field name →
  `CELL_MAPPINGS[report_type][input_sheet][input_cell] = field_name` (277) → `FIELD_DISPLAY_NAMES[field_name]
  = output_column` (278). Raises `KeyError` at import if an entry lacks `'report_type'`, `'mappings'`,
  `'input_sheet'`, `'input_cell'` or `'output_column'` keys (READ).
- **Who reads what (whole-project grep):**
  - `EXTRACTION_CONFIG`: `data_consolidator.py:24, 45` (sheets and columns), `:223, 246` (summary list).
  - `EXTRACTION_OPTIONS`: `report_processor.py:134-135` reads `skip_refresh`; `extractor_window.py:486-487`
    WRITES `skip_refresh` at run time; `data_consolidator.py:24, 129, 154` read `missing_data_text`.
    `auto_refresh_if_missing` and `show_timestamp`: NO READERS (dead options). The behaviour
    `auto_refresh_if_missing` describes is hard-coded instead (`report_processor.py:156-159`: with skip on and
    nothing found it always refreshes). HIGH.
  - `CELL_MAPPINGS`: `report_processor.py:37, 546` (decides which cells are read — this is the structure the
    extractor actually uses, despite the "Backward compatibility" label at 259); `data_consolidator.py:11,
    256` (only in dead `_write_section`).
  - `FIELD_DISPLAY_NAMES`: `data_consolidator.py:11, 275` (only in dead `_write_section`) → effectively unused.
  - No test reads this file (grep of tests/).
- **Business rules in this file:**
  - [GST RULE] TIME-VARYING — the 22 mappings in the table above (which report, sheet and cell each figure is
    taken from, and under which return/table label it is reported). 33-242.
  - [GST RULE] TIME-VARYING — output-sheet names tie figures to returns/tables: "GSTR-3B Difference",
    "GSTR-1 Difference", "Turnover GSTR-1", "Reco Diff", "Table 6J Match", "Table 8D Limits", "IMS Reco",
    "Extra In Purchase", "Extra In GSTR2B" (35, 61, 92, 105, 130, 143, 156, 181, 206). "T6 of R9" / "T8 of R9"
    are the template's sheet names for (by their wording) Tables 6 and 8 of GSTR-9 — I describe the names only.
  - Display rule — `missing_data_text = 'Missing'` (252) is what a missing/failed value is replaced with in the
    output (used at `data_consolidator.py:129, 154`).
  - Run option — `skip_refresh = False` default (249).
- **Edge cases explicitly handled:** mappings with empty `input_sheet`/`input_cell` are left out of
  `CELL_MAPPINGS` (270) — meant for calculated columns.
- **Gaps, likely failures, silent bugs:**
  - Configuration is Python code: a typo (missing comma/bracket) makes the module fail to import, which makes
    `report_processor`, `data_consolidator` and `extractor_window` fail to import. In the main app Step 4
    then shows the "Power Query refresh is unavailable ... needs pywin32" notice (`gui/tabs/extract_tab.py:
    24-30, 88-100`) — the text always blames pywin32 although `except Exception` catches any import error
    (the real error is shown after "Details:"). HIGH.
  - Same input cell mapped twice for the same report type and sheet (e.g. two output sheets both wanting
    `Reco!C37`): line 277 overwrites the first field name, so the first output column is never extracted and
    shows "Missing" (READ; not the case today, RAN). MEDIUM impact if the author ever reuses a cell.
  - A range such as `'A1:B2'` in `input_cell` is not supported: COM returns a block of values (tuple of
    tuples), and openpyxl refuses it at write time (RAN: `ValueError: Cannot convert ((1.0, 2.0),) to
    Excel`), which aborts the whole consolidated report. COM part MEDIUM (not testable here).
  - Output-sheet names must obey Excel rules: more than 31 characters → openpyxl only warns (RAN: "Title is
    more than 31 characters..."), any of `/ \ ? * [ ] :` → `ValueError` (RAN) → whole report fails. A sheet
    named "Summary" would be renamed "Summary1" silently (RAN). Current names are 8-18 characters (RAN).
  - Single-word `output_column` values (`IGST`, `CGST`, `SGST`) interact badly with the consolidator's
    substring lookup (finding 1). Multi-word columns are safe only because field names use `_` and columns use
    spaces (RAN).
  - Two entries that differ only by space vs underscore (e.g. "A B" and "A_B") produce the same field name and
    overwrite each other in `FIELD_DISPLAY_NAMES`/`CELL_MAPPINGS`. LOW.
  - A duplicated output-sheet key in the dict literal would silently keep only the last one (Python rule). LOW.
  - No validation of the configuration at import (types, cell syntax, sheet names). HIGH.
- **Excel touchpoints:** none here (data only).
- **Files read/written:** none.
- **Comments vs code:**
  - "Backward compatibility" (259) / "Build old format for compatibility" (263): misleading — `CELL_MAPPINGS`
    is the live structure the extractor uses to decide which cells to read (`report_processor.py:546`). HIGH.
  - The docstring's description of `'calculation'` matches `_calculate_value` (`data_consolidator.py:205-214`),
    which additionally returns `'Unknown'` for other types, `'Div/0'` for division by zero, `'Error'` for
    non-numbers. No mapping uses it.
  - `auto_refresh_if_missing` / `show_timestamp` comments describe behaviour no code implements (READ).
- **Pattern notes:** configuration as importable Python dicts, read via `from ..config.cell_mappings import`
  (sometimes at top of file, sometimes inside functions: `report_processor.py:134`, `data_consolidator.py:24,
  223`, `extractor_window.py:486`). Main app uses `utils/constants.py` dicts in the same style.
  `EXTRACTION_OPTIONS` is mutated at run time by the GUI (shared, process-wide state).
- **History (READ, git):** created in edac6a3 with the first three output sheets (rows 1-8 of the table);
  `power_query_extractor/config/cell_mappings copy.py` (63 lines, older format: `CELL_MAPPINGS` with field
  names `itc_diff_igst`, `sales_diff_tv`, ... for the same cells J15/K15/L15 and N1/O1/P1/Q1, plus
  `FIELD_DISPLAY_NAMES` like "Diff IGST (ITC)") existed from edac6a3 and was deleted in 47a1492.
  47a1492 (30-10-2025, the shared split point) added rows 9-22 (Reco Diff, Table 6J Match, Table 8D Limits,
  IMS Reco, Extra In Purchase, Extra In GSTR2B). No change after the split; the file is byte-identical in
  A/main and B/main (`git diff A/main B/main` empty). HIGH.

### power_query_extractor/core/__init__.py (1 line, no final newline) - [CORE]

- **Purpose:** package marker for `core`; only the docstring `"""Core processing components"""`. READ.
- **Imports / constants / functions:** none; no re-exports.
- **History:** unchanged since edac6a3. Identical in B.

### power_query_extractor/core/data_consolidator.py (432 lines) - [CORE]

- **Purpose:** after a Step 4 run, writes one new Excel workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`
  (with openpyxl, no Excel needed) into the level-1 ("Annual Statement-...") folder: a "Summary" sheet plus
  one sheet per `EXTRACTION_CONFIG` entry, one row per processed client. READ + RAN.
- **Imports:** standard library `logging` (5), `pathlib.Path` (6, UNUSED — RAN AST check), `datetime` (7);
  third-party `openpyxl.Workbook` (8), `openpyxl.styles` `Font, PatternFill, Alignment, Border, Side` (9),
  `openpyxl.utils.get_column_letter` (10); project `..config.cell_mappings` `FIELD_DISPLAY_NAMES,
  CELL_MAPPINGS` (11 — used only by the dead `_write_section`), plus function-local imports
  `EXTRACTION_CONFIG, EXTRACTION_OPTIONS` (24) and `EXTRACTION_CONFIG` (223). No OS-specific modules.
- **Module-level constants and globals:** `logger` (13). Styling values are inline literals (see below).
- **class DataConsolidator (16-432):** no `__init__`, no state. Instantiated once per panel at
  `extractor_window.py:113`.

- **`create_report(self, results, output_folder)` (20-186)** — what it actually does:
  1. 24: imports `EXTRACTION_CONFIG`, `EXTRACTION_OPTIONS` (comment 23 "Import the new configuration").
  2. 27-31: new `Workbook()`; removes its default sheet.
  3. 34-42: styles — `header_font` bold, white, size 12; `header_fill` solid `0078D4` (blue);
     `section_font` bold 14; `border` thin on all four sides.
  4. 45: for every output sheet in `EXTRACTION_CONFIG` (file order): log "Creating sheet: ..." (46);
     `wb.create_sheet(sheet_name)` (49).
     - 52-71: if the config entry has a `'description'` list (NO entry has one → this branch never runs),
       each line goes to A1, A2, ... (first line bold 14, others italic 11, each merged A:J) and
       `start_row = len + 2`; otherwise A1 = `"<sheet name> Report"` in bold 14, merged A1:C1, `start_row = 3`.
     - 74-80: builds `['Client Name', 'Successfully Run']` and writes them, UNFORMATTED, into row 3 columns
       1-2 — leftover code; immediately overwritten by 86-98. 83 sets `current_row` (repeated at 101).
     - 86-98: rebuilds headers = 'Client Name', 'Successfully Run', then every mapping's `output_column` in
       config order; writes them in row `start_row` with header font, blue fill, centred, bordered.
     - 101-102: `current_row = start_row + 1`; `report_type = sheet_config['report_type']`.
     - 104-166: for each client result, in the order processed:
       - column 1 = `result['client']` (the client folder name) with border (106);
       - `report_data = result.get(report_type.lower(), {})` → the `'itc'` or `'sales'` part (109);
       - `is_success = report_data['status']['success']` (via `.get`, default False) (110);
       - column 2 = `'✓'` (green `107C10`, bold) if success else `'✗'` (orange-red `D83B01`, bold) (112-121);
       - columns 3+ (124-164), per mapping: start `value = None`;
         - if NOT success → `value = 'Missing'` (`missing_data_text`) — **even when values were extracted**
           (128-129);
         - elif the mapping has a `calculation` → `_calculate_value(report_data['extracted_values'], ...)`
           (130-136; never happens today);
         - else (137-150): look through `report_data['data']` in insertion order and take the FIRST key that
           CONTAINS `output_column` as a substring (`if mapping['output_column'] in field_name`, 142-145);
           only if that gave `None` (no match, or the matched value is None) fall back to the exact key
           `f"{sheet_name}_{output_column}".replace(' ', '_')` (147-150). `mapping['input_cell']` must exist
           (KeyError otherwise).
         - 153-154: `None` or `''` → `'Missing'`. Zero stays 0.
         - 157-158: write the value with a border; 161-162: if `int`/`float` and `!= 0` → number format
           `'#,##0.00'` (so `True`/`False`, being ints in Python, also get it; 0 keeps "General").
       - `current_row += 1` (166).
     - 169: `_autofit_columns(ws)`.
  5. 172: `_create_summary_sheet(wb, results)` (inserted as the FIRST sheet).
  6. 175-177: name `PQ_Extraction_Report_{datetime.now():%Y%m%d_%H%M}.xlsx`; path `output_folder / name`.
  7. 179: `wb.save(report_path)` — silently overwrites a file of the same name (same minute); 180 logs.
  8. 182: returns the `Path`. 184-186: any exception → `logger.error(..., exc_info=True)` and **re-raised**.
  - Parameters: `results` = list of `process_client` results (section on report_processor);
    `output_folder` = a `Path` (the level-1 folder). Both used.
  - Side effects: one file created/overwritten. Does not create folders (missing folder → error raised).
  - Callers: `power_query_extractor/gui/extractor_window.py:945` only (worker thread). No test calls it.
  - **RAN (fake data, distinct number per cell; `/root/merge-scratch/run/A-R2/out/PQ_Extraction_Report_20261021_0405.xlsx`):**
    sheet order `Summary, GSTR-3B Difference, GSTR-1 Difference, Turnover GSTR-1, Reco Diff, Table 6J Match,
    Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B`; each data sheet: row 1 title merged A1:C1,
    row 2 blank, row 3 headers, rows 4+ clients. With J15=100, K15=101, L15=102, AA1=108, AD1=109, AG1=110,
    O1=111, ..., U1=116, the sheets **IMS Reco, Extra In Purchase and Extra In GSTR2B all showed 100, 101, 102**
    (the GSTR-3B difference values), because `'IGST' in 'GSTR-3B_Difference_Diff_IGST'` is true for the first
    stored key. The 13 multi-word columns were correct (their spaces never match the underscore field names).
    A client whose ITC status was "failed" but whose data dict held all values showed `✗` and `Missing` in
    every ITC column.

- **`_calculate_value(self, extracted_values, calculation, workbook)` (188-218):** reads `type`, `cell1`,
  `cell2` from the calculation dict; looks the two cells up in `extracted_values` by BARE cell reference
  (default 0) — these keys are written by `report_processor.py:588` and do not include the sheet name, so
  `cell1='O1'` means "whichever sheet's O1 was read for this report type"; converts each with
  `float(v) if v else 0` (None, '' and 0 → 0; any non-number text → returns `'Error'` via bare `except`,
  201-202); returns `num1 - num2` / `+` / `*` / `num1 / num2` (or text `'Div/0'` if `num2 == 0`) / `'Unknown'`
  for any other type; outer `except` logs and returns `'Error'`. Parameter `workbook` is NEVER USED.
  Results are not rounded. Callers: `create_report` line 132, reached only if a mapping has a `calculation`
  — none does, so the function is dormant (HIGH).
  - [GST RULE] (dormant): blank or missing cells count as 0 in a calculation; division by zero yields the text
    `Div/0`; non-numeric text yields `Error`.

- **`_create_summary_sheet(self, wb, results)` (220-248):** local import of `EXTRACTION_CONFIG` (223, comment
  "# ADD THIS LINE:"); creates sheet "Summary" at position 0; A1 "Extraction Summary" (bold 16); A3
  `Generated: <YYYY-MM-DD HH:MM:SS>`; A4 `Total Clients: <n>`; A5 `ITC Reports Processed: <x>/<n>`; A6
  `Sales Reports Processed: <y>/<n>`, where x/y count results whose `status.success` is truthy (235-236); A8
  "Sheets Created:" (bold); A9 onwards `• <output sheet name>` for each config key. No widths set (RAN). No
  return value. The denominator counts every client, including those not selected for that report type
  (a Sales-only client counts as an ITC non-success). Caller: `create_report` 172. RAN (output above:
  "ITC Reports Processed: 1/2").

- **`_write_section(self, ws, results, report_key, report_type, start_row, header_font, header_fill, border,
  subsection_font)` (250-352): NO CALLERS FOUND (dead; dead since the initial commit, because this file is
  unchanged since edac6a3).** What it would do: collect every field name of `CELL_MAPPINGS[report_type]` plus
  every key actually present in the results' `data` (which would include bare cell references such as `J15`),
  sort them, write headers 'Client Name', 'Refresh Status', display names (`FIELD_DISPLAY_NAMES`), 'Error
  Details'; one row per client with '✓ Success'/'✗ Failed', values ('-' for None), the error text in orange
  size 9; then a "Summary:" row "x/y successful". Returns the last row. It is the only code that would put the
  **error reason** into the workbook. Probably the original one-sheet "ITC section / Sales section" layout
  replaced by the per-sheet `EXTRACTION_CONFIG` design (MEDIUM inference from the name and from the old
  "cell_mappings copy.py" format).

- **`_create_details_sheet(self, wb, results)` (354-414): NO CALLERS FOUND (dead).** Would add sheet
  "Extraction Details": per client "Client: <name>", "ITC Report:", "Status: Success/Failed", "Validation:
  <validation_message>" (lines 379-382 marked "# ADD THESE LINES"), "Extracted Values:" with every
  `field: value`, then the same for Sales. The "# ADD THESE LINES" edits (379, 402) were pasted into a method
  nothing calls, so the refresh-validation message never reaches any output file. HIGH.

- **`_autofit_columns(self, ws)` (416-432):** for each column, the longest `str(value)` among truthy cells
  (0 and empty ignored; bare `except: pass`); width = `min(max_length + 2, 50)`; always set (the `> 0` test at
  431 is always true). The merged title in A1 counts toward column A. Caller: `create_report` 169.

- **Business rules in this file:**
  - Reporting rule (author-owned, decides whether a figure is shown): a report whose status is not success
    shows `Missing` in every column, regardless of extracted values (128-129). [GST RULE — tagged under
    "when in doubt", because it decides whether figures appear.]
  - [GST RULE] rounding (display only): non-zero numbers are shown with format `#,##0.00` — two decimals and
    thousands separators; the stored cell value keeps full precision (161-162). Zero is shown as plain `0`.
  - Missing/blank value → `'Missing'` (153-154).
  - "Successfully Run" = ✓ exactly when `status.success` is true (110-121); see report_processor for what
    makes it true.
  - Dormant [GST RULE]s in `_calculate_value` (above).
- **Edge cases handled:** missing `'status'`/`'data'` keys (`.get` with defaults, 109-110, 139); blank/None
  values (153); division by zero (212); non-numeric calculation input (201).
- **Gaps, likely failures and silent bugs:**
  - Substring look-up cross-wires 9 columns (finding 1; RAN). HIGH.
  - Figures dropped when status is false; ✓ rows with all `Missing` when extraction itself failed
    (`report_processor.py:615-617` returns `{}` but the refresh status stays success). HIGH.
  - No error reason, validation message, file name or refresh time in the workbook (only ✓/✗) — the
    methods that wrote them are dead. HIGH.
  - Excel error cells (#N/A, #REF! ...) come back through COM as large negative integers (MEDIUM — general
    pywin32 behaviour, not testable here); they would be written as numbers with `#,##0.00`
    (RAN: openpyxl accepts and saves -2146826246). How to check: put `=NA()` in a mapped cell and run Step 4.
  - A date value from COM arrives as a timezone-aware datetime (MEDIUM, same caveat); openpyxl then fails the
    WHOLE save (RAN: `TypeError: Excel does not support timezones in datetimes`). Mapped cells are amounts, so
    LOW likelihood.
  - Any exception (bad sheet name, range value, file open in Excel, network folder gone) aborts the whole
    report and is re-raised; the extracted results then exist only in the GUI log and the log file — no
    retry, no fallback location. HIGH.
  - Same-minute rerun overwrites the earlier consolidated report without warning (READ, minute-resolution
    name). HIGH.
  - Only the clients processed in THIS run appear; nothing is merged with earlier runs. HIGH.
  - The leftover first header write (77-80) is harmless (overwritten). HIGH.
- **Excel workbook touchpoints:** openpyxl `Workbook()` create (27); `wb.remove(default sheet)` (31);
  `wb.create_sheet` per config entry (49) and "Summary" at index 0 (225); `ws.merge_cells` (62, 70);
  `wb.save(report_path)` (179). No COM, no loading of existing workbooks. It never saves over a Power Query
  workbook (it always creates a new one), so Standard 12's openpyxl warning does not arise here.
- **Files written:** `<level-1 folder>/PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` (format note: year-first,
  unlike the rest of the app's DDMMYY stamps — `utils/helpers.py:25-35`, `report_processor.py:128`).
- **Comments and docstrings vs code:** paste-instruction leftovers "# ... rest of the code" (75), "# Use
  start_row instead of 3" (79), "# ... rest of header formatting" (80), "# Instead of 4" (83), "# ADD THIS
  LINE:" (222), "# ADD THESE LINES" (379, 402). Docstring of `create_report` "multiple custom sheets" — true.
  Module docstring "Data consolidation into master report" — true.
- **Pattern notes:** errors are logged and re-raised (unlike `report_processor`, which converts every error into
  a result dict). Colours `0078D4`, `107C10`, `D83B01` match the GUI palette. No thread concerns (runs once on the
  worker thread, touches no widgets).
- **History:** unchanged on A/main since edac6a3 (`git diff edac6a3 A/main` empty for this file); B changed it
  after the split (19 lines, B's readers).

### power_query_extractor/core/refresh_naming.py (66 lines) - [CORE]

- **Purpose:** small pure functions that decide whether a file name is a "refreshed copy" of a report, based
  on the user's "Refreshed File Suffix" pattern, plus the legacy `_refreshed_` marker. READ + RAN.
- **Imports:** none.
- **Module-level constants:**
  - `DEFAULT_SUFFIX_PATTERN = "_Refreshed_{timestamp}"` (14) — read only by tests
    (`tests/test_refresh_naming.py:11, 18, 39, 53, 112`); imported but UNUSED in `report_processor.py:38`;
    the same literal is repeated in `report_processor.py:51` and `extractor_window.py:297` instead of using it.
  - `TIMESTAMP_TOKEN = "{timestamp}"` (15) — used at 39 only. The writer side
    (`report_processor.py:129`) repeats the literal `"{timestamp}"`.
  - `LEGACY_MARKER = "_refreshed_"` (20) — "Always accepted in addition to the configured pattern" (17-19).
- **`refresh_markers(suffix_pattern=None)` (23-40):** if the pattern is empty/None → `['_refreshed_']` (37-38);
  else LOWER-CASES the pattern, splits it on `{timestamp}` and keeps the non-empty pieces (39); if nothing is
  left (pattern was just `{timestamp}`) → `['_refreshed_']` (40). Returns a list of lower-case strings.
  Callers: `refresh_naming.py:52`; tests `tests/test_refresh_naming.py:18, 21, 24, 29, 33`.
- **`is_refreshed_name(filename, suffix_pattern=None)` (43-55):** lower-cases the name (49); not ending in
  `.xlsx` → False (50-51) (so `.xlsm`/`.xlsb` are never recognised); if ALL markers occur ANYWHERE in the name
  → True (52-54); else True if `_refreshed_` occurs anywhere (55). Callers: `extractor_window.py:64`,
  `refresh_naming.py:66`; tests `tests/test_refresh_naming.py:38, 43, 49, 53, 54, 57, 61, 65, 68`.
- **`is_refreshed_copy_of(filename, original_stem, suffix_pattern=None)` (58-66):** True if the lower-cased
  name STARTS WITH the lower-cased original stem (name without `.xlsx`) and `is_refreshed_name` is true.
  Callers: `report_processor.py:144`; tests `tests/test_refresh_naming.py:76, 80, 84`.
- **Business rules:** none of GST. Naming rule: "a refreshed file is any `.xlsx` whose name contains every
  literal piece of the configured suffix, or `_refreshed_`" (READ).
- **Edge cases handled:** empty/None pattern; pattern with no literal part; upper/lower case in names; `.xlsx`
  only; legacy files after a suffix change (docstring 17-19, 46-47).
- **Gaps (RAN, `python -c` probes, original name `ITC_Report_CLIENT_DL_210726_1130.xlsx`):**
  - Markers are searched in the WHOLE file name, not in the suffix position. Patterns whose literal part also
    occurs in the normal report name make the ORIGINAL count as refreshed: `_R{timestamp}` → marker `_r`
    (inside "itc**_r**eport") → True; `_{timestamp}` → marker `_` → True; `_Report{timestamp}` → True;
    `_{TIMESTAMP}` → True. `is_refreshed_copy_of(original)` is then True as well, so Skip Refresh can pick the
    original (see report_processor RAN case 5). HIGH.
  - A client whose sanitised name contains the word "Refreshed" (e.g. `..._R_Refreshed_Co_...`) makes its
    original report match the legacy marker. LOW likelihood. (RAN: `sanitize_filename('R_Refreshed_Co')` keeps it.)
  - Pieces are not checked in order (`_v{timestamp}_final` accepts "..._final..._v..."). LOW.
  - Reader lower-cases the pattern (so `{TIMESTAMP}` acts as the token here) but the writer
    (`report_processor.py:129`) replaces only lower-case `{timestamp}`, so with `{TIMESTAMP}` every refresh
    writes the literal name `..._{TIMESTAMP}.xlsx` and overwrites the previous one (RAN case 6). MEDIUM impact.
  - `is_refreshed_copy_of` is a prefix test: another report whose stem merely starts with the same text would
    count (RAN: `..._11300_Refreshed_x.xlsx` accepted for stem `..._1130`). LOW (stems end in a full timestamp).
- **Excel touchpoints / files:** none (names only).
- **Comments vs code:** module docstring (1-12) — accurate; "so all three agree" (11) refers to writer +
  two readers. The function docstrings' examples match RAN results.
- **Pattern notes:** pure, side-effect-free, unit-tested (20 tests added in 2aaa2b1). One of the few places
  that centralises a rule both GUI and core use.
- **History:** added in 2aaa2b1 (25-07-2026, A only) "Recognise refreshed files by the CONFIGURED suffix, not a
  hardcoded one"; before that both readers hard-coded `_Refreshed_` (glob `"{stem}_Refreshed_*.xlsx"` in
  report_processor, `'_refreshed_' in name` in extractor_window). Not present in B.

### power_query_extractor/core/report_processor.py (804 lines) - [CORE]

- **Purpose:** for one client, finds the ITC and/or Sales report in the client's chosen version folder, makes
  a timestamped copy, refreshes Power Query in that copy by remote-controlling Excel (key presses), checks a
  marker cell, saves, then opens the copy again in a second, hidden Excel and reads the configured cells.
  Returns a result dict per report. READ (COM parts cannot run on Linux); non-COM parts RAN with stubs.
- **Imports:**
  - Standard library: `os` (5, UNUSED — RAN AST), `time` (6), `shutil` (7), `logging` (8), `pathlib.Path` (9),
    `datetime` (10).
  - OS-specific (Windows/pywin32): `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` —
    **GUARDED** by `try/except ImportError` (20-36). On failure each name is set to `None`,
    `WIN32_AVAILABLE = False`, and a warning is logged AT IMPORT TIME (34-36: "pywin32 not available (...).
    Power Query refresh is disabled; install it with: pip install pywin32"). No platform check — it relies on
    the import failing. A second, local `import win32api` sits inside `_safe_open_workbook` (762), inside a
    bare `try/except`. COM ProgIDs (the names Windows uses to find a program's remote-control interface) used:
    `"Excel.Application"` (218, 536) and `"WScript.Shell"` (260, Windows' scripting helper that can bring a window
    to the front and type keys).
  - Project: `..config.cell_mappings.CELL_MAPPINGS` (37); `..core.refresh_naming` `DEFAULT_SUFFIX_PATTERN`
    (UNUSED) and `is_refreshed_copy_of` (38); `utils.excel_com.find_open_workbook` (39) — an ABSOLUTE import of
    the top-level `utils` package, so the project root must be on `sys.path` (true for `main.py` and
    `launch_extractor.py:10`); function-local `from ..config.cell_mappings import EXTRACTION_OPTIONS` (134).
- **Module-level constants and globals:** `WIN32_AVAILABLE` (26/33) — read at 66 and by
  `tests/test_optional_pywin32.py:46`; `logger` (41). All other operational values are literals inside the
  methods — listed under "Operational constants" below.

- **class ReportProcessor (44-804).** One instance per panel (`extractor_window.py:112`), reused for every
  client of every run, and shared by any concurrent run (see finding 2).

- **`__init__` (47-51):** `excel_instance = None` (48, NEVER READ — dead), `log_callback = None` (49),
  `wait_time = 10` (50, "Default wait time in seconds"), `suffix_pattern = "_Refreshed_{timestamp}"` (51).
  `wait_time`/`suffix_pattern` are overwritten by the GUI before each run (`extractor_window.py:864-865`).

- **`_log(self, message, level='info')` (53-60):** always `logger.info(message)` (56) — whatever the level;
  then, if a callback is set, `log_callback("  " + message, level)` (59-60). All 97 calls in this file use the
  default level, so every processor line (including "❌ ..." errors) is shown in the GUI log in the plain
  'info' colour (READ: grep finds no `_log(..., '<level>')`). Messages logged only with `logger.warning/error`
  (18 calls, e.g. 230, 234, 246, 270, 274, 316, 339, 377, 389, 553, 571, 592, 597, 616) never reach the GUI.
  Callers: 97 sites in this file.

- **`process_client(self, client_data, log_callback=None)` (62-106):**
  - Stores the callback (64) — instance attribute, so a second concurrent run would replace the first's.
  - If `not WIN32_AVAILABLE` (66-77): logs "❌ pywin32 is not installed for this Python, so Power Query cannot
    be refreshed. Install it with: pip install pywin32" and returns `{'client': name, 'timestamp': now,
    'itc': copy, 'sales': copy}` where both copies are `{'status': {'success': False, 'error': message},
    'data': {}, 'extracted_values': {}}` (a shallow `dict(empty)` copy — both share the same inner
    `status` dict; harmless, nobody mutates it). The GUI run then still writes a consolidated report full
    of ✗/"Missing" and offers to open it (READ via `extractor_window.py:941-955`). HIGH.
  - Otherwise builds `results` (79-84) and, for ITC then Sales: if `client_data.get('process_itc', True)`
    (87) / `'process_sales'` (97) → `_process_report(client_data, 'ITC'|'Sales')`, else a "Not selected for
    processing" failure dict (90-94, 100-104). The `True` defaults are "for backward compatibility" (87, 97).
  - Returns `{'client': str, 'timestamp': datetime, 'itc': report_result, 'sales': report_result}`.
    `'timestamp'` is never read anywhere (grep). Needs keys `name`, `latest_version`.
  - Callers: `power_query_extractor/gui/extractor_window.py:914` (worker thread);
    `tests/test_optional_pywin32.py:50`.

- **`_process_report(self, client_data, report_type)` (108-207):** the per-report pipeline.
  1. 112-113: `latest_version.glob(f"{report_type}_Report_*.xlsx")` — `ITC_Report_*.xlsx` or
     `Sales_Report_*.xlsx`. **This also matches earlier refreshed copies** (`..._Refreshed_....xlsx` start
     with the same prefix). 115-123: none → `{'status': {'success': False, 'error': "No ITC report found"}, ...}`.
  2. 125: `original_file = report_files[0]` — the FIRST in directory order. On NTFS, listings are
     alphabetical (case-insensitive) and the original (`..._1130.xlsx`, `.` sorts before `_`) usually comes
     first with the default suffix (MEDIUM — NTFS ordering, not testable here); on other file systems (some
     NAS shares) order is arbitrary. RAN (Linux):
     in a folder holding `X.xlsx` and `X_Refreshed_1.xlsx`, `report_files[0]` was the refreshed copy, and the
     run produced `X_Refreshed_1_Refreshed_211026_0406.xlsx` (a copy of a copy). With a custom suffix that
     sorts before `.` (e.g. starting with a space or `-`), NTFS order would also put the copy first (MEDIUM).
  3. 128-130: `timestamp = now.strftime("%d%m%y_%H%M")` (DDMMYY_HHMM, minute resolution); suffix =
     `suffix_pattern.replace("{timestamp}", timestamp)` (case-sensitive); `refreshed_file =
     <same folder>/<original stem><suffix>.xlsx`.
  4. 132-135: comment "ADD THE NEW CODE HERE (notice the indentation - 3 levels)" (paste leftover); reads
     `EXTRACTION_OPTIONS['skip_refresh']` — the module dict the GUI checkbox mutates.
  5. Skip-refresh branch (137-159): candidates = `glob(f"{original_stem}*.xlsx")` filtered by
     `is_refreshed_copy_of(name, stem, suffix_pattern)` (142-146); if any → `refreshed_file` = the one with the
     newest mtime (149), log "Skip refresh enabled - using existing: <name>" (150), status
     `{'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}`
     (151-155); else log "No existing refreshed file found - will refresh" and fall through to refresh (157-159).
     **Nothing checks that the chosen file was ever successfully refreshed** (see 6 below). The glob uses the
     stem as a pattern, so a stem containing `[`/`]` matches nothing (RAN) — but `sanitize_filename`
     (`utils/helpers.py:96`) already replaces `[ ]` in client names (RAN), so only hand-renamed files are affected.
  6. Refresh branch (162-181): `shutil.copy2(original, refreshed_file)` (164) — **copy2 also copies the
     original's date-modified time**, so until Excel saves the copy it looks as old as the original; an existing
     file of the same name (same minute) is silently overwritten; if the copy fails midway (network) a partial
     file may remain under the "Refreshed" name (MEDIUM). Then `_refresh_power_query_simple(refreshed_file)`
     (168). The `except` at 169-181 (special-casing "Open method of Workbooks class failed" / "-2147352567")
     is **effectively dead**: `_refresh_power_query_simple` catches every exception itself (388-393) and returns
     a dict. HIGH.
  7. 183-185: log "Waiting for file to be released..." and `time.sleep(3)` — also in skip mode.
  8. 188: `_extract_data(refreshed_file, report_type)` — "regardless of refresh status" (187).
  9. 190-196: returns `{'status': refresh_status, 'data': extracted, 'extracted_values': extracted (the SAME
     dict object), 'original_file': str, 'refreshed_file': str}`. The two file keys are never read (grep).
  10. 198-207: any exception (e.g. copy failure, SameFileError) → `logger.error(exc_info=True)` and
      `{'status': {'success': False, 'error': str(e)}, 'data': {}, 'extracted_values': {}}`. Not shown in
      the GUI except as the per-report "⚠️ ... - <error>" line (`extractor_window.py:923`).
  - RAN (stubbed Excel steps, `/root/merge-scratch/run/A-R2/out/tmp1fluxiqu`): default suffix → copy
    `..._Refreshed_211026_0406.xlsx`, refresh+extract called on it; **empty suffix** → `shutil.SameFileError`
    caught → status error "... are the same file" (original untouched); suffix `/sub{timestamp}` →
    `FileNotFoundError` (sub-folder) caught; skip with default suffix → existing copy reused, success True;
    **skip with `_R{timestamp}` → the ORIGINAL file was used for extraction and reported "Skipped - Using
    existing refreshed file"**; `_{TIMESTAMP}` → file literally named `..._{TIMESTAMP}.xlsx`; Sales report
    missing → "No Sales report found".
  - Callers: `report_processor.py:88, 98`.

- **`_refresh_power_query_simple(self, excel_file)` (209-415) — the Excel refresh, step by step (READ):**
  1. 215: `pythoncom.CoInitialize()` (prepares COM for this worker thread).
  2. 217-218: "Creating Excel instance..." → `win32com.client.Dispatch("Excel.Application")` (may attach to
     a running Excel — see Standard 12 comparison below).
  3. 221-234: up to 5 attempts (2 s apart): `excel.Visible = True` ("Keep it visible to avoid issues"),
     `excel.DisplayAlerts = False`, `excel.WindowState = -4137` (xlMaximized). After 5 failures:
     "Could not set Excel properties, continuing anyway..." (logger only).
  4. 236-249: "Opening Excel file: <name>"; up to 3 attempts of `_safe_open_workbook(excel, file)`; on an error
     containing "Call was rejected by callee" (Excel busy) wait 3 s and retry; any other error is re-raised
     (→ outer except 388). **If all 3 attempts are "rejected", the loop ends with `wb = None` and the code
     carries on** (waits, key presses), failing later at `wb.Save()` (inner except 376-381). HIGH.
  5. 251-255: "Waiting for workbook to fully load..." — 5 × 1 s, logging "Loading... i/5 seconds".
  6. 258-276: `WScript.Shell` object; 3 attempts of `shell.AppActivate(excel.Caption)` (bring Excel's window
     to the front by its title), 1 s apart; bare `except` around each. If never activated: logger warning
     "Could not activate Excel window, continuing anyway..." — **the key presses are still sent to whatever
     window has the focus** (another program, an e-mail being typed...). 276: sleep 1.
  7. 278-293: "Starting refresh sequence..." → `SendKeys("%a")` (Alt+A: opens the Data ribbon tab's key
     tips), sleep 1, `SendKeys("r")`, sleep 1, `SendKeys("a")` → "Refresh All command executed". This depends
     on English ribbon key tips, an unlocked desktop and Excel keeping the focus (MEDIUM — cannot test). 295:
     stray comment "# Rest of your code remains same...".
  8. 297-302: "Waiting for Power Query refresh to complete... (<wait_time> seconds)"; `wait_time` × 1 s.
     Nothing checks whether the refresh actually finished.
  9. 305-322: error-dialog check 5 times, 1 s apart: `_find_error_window()`; if found → read its text
     (`_get_window_text`), logger warning, `_close_error_window`, stop checking. Else "✓ No errors detected".
  10. 325-330: "Giving additional time for complex queries..." — `max(10, wait_time)` × 1 s (logs every 2 s).
  11. 333-342: one more error-dialog check if none was found earlier.
  12. 345: `_validate_refresh(wb)` → (ok, message).
  13. 348-351: `wb.Save()` **always — even if an error dialog was seen or validation failed**; sleep 2.
  14. 354-357: validation failure → `error_detected = True`, message "Refresh validation failed: <msg>".
  15. 360-363: `wb.Close(SaveChanges=True)`; sleep 1.
  16. 365-374: logs "⚠️ Refresh completed with errors: ..." or "✅ Refresh completed successfully"; returns
      `{'success': not error_detected, 'error': message or None, 'validation_message': message}`.
  17. 376-381: any exception in steps 6-16 → logger "❌ Error during refresh" → `{'success': False, 'error': str(e)}`
      (no `validation_message`).
  18. 383-386: `return {'success': True, 'error': None}` — **UNREACHABLE** (both branches above return). HIGH.
  19. 388-393: exception in steps 1-5 → logger error with traceback → `{'success': False, 'error': str(e)}`.
  20. `finally` 394-415 (always): "Cleaning up Excel instance..."; `if wb: wb.Close(SaveChanges=False)` (the
      workbook is usually already closed → error swallowed by bare `except: pass`); `if excel: excel.Quit()`
      (bare except); `pythoncom.CoUninitialize()` (bare except) — called while the local variables `excel` and
      `wb` still hold COM objects (they are released only when the function's frame ends, after
      uninitialising). Standard 12 (rebuild_pipeline_FINAL.md:872-875) asks for the Excel work in a separate
      function and `CoUninitialize()` after it returns. HIGH (code), MEDIUM (effect: possible stray EXCEL.EXE
      or COM warnings).
  - Total fixed waiting per refreshed report with default 10 s (when nothing fails): 5 (load) + 1 + 2 (keys)
    + 10 (wait) + 5 (dialog checks) + 10 (extended) + 2 (after save) + 1 (after close) = **36 s**, plus 3 s in
    `_process_report` and Excel start/open times. With 60 s: 5+1+2+60+5+60+2+1+3 = **139 s**. Two reports per
    client double it (READ arithmetic, HIGH).
  - Callers: `report_processor.py:168`.

- **`_find_error_window(self)` (417-453):** enumerates **every top-level window on the desktop**
  (`win32gui.EnumWindows`), keeping those that are visible and enabled, whose title (lower-cased) contains any
  of `'microsoft excel', 'error', 'problem', 'cannot', 'failed', 'unable', 'could not', 'exception', 'invalid',
  'missing'` (425-436) AND whose window class is `'#32770'` (standard Windows dialog), `'NUIDialog'` (Office
  dialog) or `'bosa_sdm_msword'` (441). Returns the first such handle or None; all errors swallowed (444-445,
  452-453). Not limited to Excel: a dialog of ANY program (including this app's own Tk message boxes, which on
  Windows are `#32770` dialogs titled e.g. "Error") qualifies (MEDIUM-HIGH). `'bosa_sdm_msword'` is, by its
  name, a Word dialog class; Excel's classic dialogs use a different class name (I believe `bosa_sdm_XL9` —
  LOW, not sure; check with `win32gui.GetClassName` on a live Excel dialog). Callers: 312, 335.
- **`_get_window_text(self, hwnd)` (455-490):** window title plus the text of child controls whose class
  contains "static" or "text" and whose text is longer than 5 characters (466-470); joined, whitespace
  collapsed, first 200 characters; fallback "Power Query error occurred". This text goes into the log file and
  the result's `error`, then the GUI log. Callers: 315, 338.
- **`_close_error_window(self, hwnd)` (492-525):** looks through the dialog's child windows for a button whose
  text contains (substring, lower-case) `'ok'`, `'close'`, `'yes'` or `'continue'` and **clicks it**
  (`BM_CLICK`, 508); stops at the first. If none clicked → posts `WM_CLOSE` (518-520). Sleep 0.5 s. Errors only
  at debug level. Clicking "Yes" on an unrelated program's dialog is possible because of the broad search above
  (MEDIUM). Parameter `unused` in the inner callback is indeed unused (500). Callers: 317, 340.

- **`_extract_data(self, excel_file, report_type)` (527-636) — reading values back:**
  1. 533-539: `CoInitialize()`; a NEW `Dispatch("Excel.Application")` (536); `Visible = False`,
     `DisplayAlerts = False`; logs "✓ Excel instance created (hidden mode)" (if Dispatch attached to an
     already-running Excel, that Excel is hidden — MEDIUM).
  2. 541-543: logs "Opening file in read-only mode" / "✓ File opened in read-only mode", but
     `_safe_open_workbook` opens it normally (not read-only) — log text disagrees with code. HIGH.
  3. 546-554: `mappings = CELL_MAPPINGS.get(report_type, {})`; logs counts; none → warning and `{}`.
  4. 556-601: for each report sheet name: find a sheet whose `Name == sheet_name` EXACTLY (case- and
     space-sensitive; 564-568; compare `_validate_refresh`, which is case-insensitive); not found → logger
     warning and every field of that sheet = `None` (570-576). Found → for each cell: `ws.Range(cell).Value`
     (584) stored under the field name AND under the bare cell reference (585-588); per-cell error → both
     `None` (591-594). Any sheet-level error → all its fields `None` (596-601).
  5. 603-611: logs "Extraction complete. Total values extracted: <n>" — n counts field names AND cell keys,
     so it is about double the number of cells — then every key and value (figures go into the GUI log and the
     log file).
  6. 613: returns the dict `{field_name: value, cell_ref: value, ...}` (insertion order: sheets in
     `CELL_MAPPINGS` order, within a sheet field then cell). 615-617: any other error → logger error and `{}`.
  7. `finally` 618-636: `wb.Close(SaveChanges=False)`, `excel.Quit()`, `CoUninitialize()` — each in bare
     `try/except: pass`; same CoUninitialize-while-referenced pattern as above.
  - Values are what Excel shows after opening the file (Excel recalculates on open; if the queries are set to
    refresh on open, a background refresh may start — LOW, not testable).
  - Callers: `report_processor.py:188`.

- **`_validate_refresh(self, wb)` (638-685) — the only "did the refresh work" test:** finds a sheet whose
  lower-cased name is `'info'` (645-648) → else (False, "Info sheet not found"); reads `Info!BB2` (656);
  `None` → (False, "Cell BB2 is empty"); text blank after strip → "Cell BB2 is blank"; `'0'`/`'0.0'` → "Cell BB2
  contains 0"; upper-cased text in `['#REF!', '#VALUE!', '#NAME?', '#NULL!', '#DIV/0!', '#N/A', '#NUM!',
  '#ERROR!']` → "Cell BB2 contains error: ..."; otherwise (True, "Validated - Client Name: <value>"). Any
  exception → (False, "Validation error: ..."). RAN with fake sheets: None/''/'   '/0/0.0/'0'/'#N/A'/'#n/a' →
  False; `'SomeName'`, `12.5`, and **`-2146826246` (how COM is expected to return an #N/A cell — MEDIUM) →
  True**; no Info sheet → False; `wb=None` → False ("Validation error: 'NoneType' object has no attribute
  'Sheets'"). It cannot tell a freshly loaded value from one already in the file (READ, HIGH).
  - [GST RULE — tagged under "when in doubt"; A-X-EXCEL treats it as technical]: a report counts as refreshed
    only if its `Info` sheet's cell `BB2` (comment 654: "client name") is non-empty, non-zero and not an error
    text. TIME-VARYING (depends on template layout). Its result decides whether that report's figures appear
    or become "Missing" in the consolidated workbook.
  - Callers: `report_processor.py:345`.

- **`_safe_open_workbook(self, excel, file_path)` (687-799) — opening with fallbacks (RAN with a fake Excel):**
  1. 689: `Path(file_path).absolute()`.
  2. 692-694: missing → log "❌ File does not exist" and raise `FileNotFoundError`.
  3. 697-703: opens the file in Python and reads 1 byte as an access test; failure → log and re-raise.
  4. 705-712: `find_open_workbook(excel, path)` (`utils/excel_com.py`) — if this Excel already has it open,
     returns that workbook ("📌 File already open, using existing workbook").
  5. 714-732: path length logged; if > 218 characters ("Excel's typical limit"): `/` → `\`, then prefix
     `\\?\` (local) or `\\?\UNC\` (network `\\server\...`) — Windows "extended-length path" syntax. Whether
     Excel's `Workbooks.Open` accepts such paths: NOT SURE (LOW); later attempts fall back to plain paths.
  6. Attempt 1 (735-741): `Workbooks.Open(file_str)` — no options (so no `UpdateLinks`, no
     `IgnoreReadOnlyRecommended`, not read-only).
  7. Attempt 2 (744-753): **closes every workbook in this Excel whose file NAME (not full path) equals this
     file's name, without saving** (746-749), waits 3 s, opens again. In an Excel the user is working in, a
     different workbook with the same file name would be closed and its unsaved changes lost (MEDIUM —
     depends on Dispatch attaching).
  8. Attempt 3 (756-783): waits 2 s; tries the 8.3 short path (`win32api.GetShortPathName`) if different
     (761-766; any error there, including a failed Open, is reported as "Could not get short path"); then, if
     the extended prefix was added, `Open(plain path)` (771-774) — and if that fails the "all parameters"
     attempt is skipped; else `Open(Filename=..., UpdateLinks=0, ReadOnly=False,
     IgnoreReadOnlyRecommended=True, Notify=False)` (777-783).
  9. 784-799: all failed → logs the error and, for "-2147352567" (a generic "exception occurred" code that
     Excel uses for most failures), prints likely causes (open elsewhere, corrupted, path too long, OneDrive/
     SharePoint sync) and "Path length is definitely an issue!" if > 218; re-raises.
  - RAN (Linux, 232-character path, fake Excel whose `Open` always fails): three `Open` calls were made —
    `\\?\...` twice, then the plain path; log lines exactly as listed; the error was re-raised.
  - `workbook.ReadOnly` is never checked after opening (Standard 12, 857-859). HIGH.
  - Callers: `report_processor.py:241, 542`.

- **`cleanup(self)` (801-804):** does nothing ("Nothing to cleanup with this approach as we create/destroy Excel
  each time"). Caller: `extractor_window.py:966`.

- **Business rules in this file:**
  - Report file pattern: `<ITC|Sales>_Report_*.xlsx` in the chosen version folder (112); first match used (125).
  - Refreshed copy name: `<original stem><suffix>.xlsx`, suffix = pattern with `{timestamp}` →
    `DDMMYY_HHMM` (128-130).
  - Skip refresh: reuse the newest file recognised as a refreshed copy; otherwise refresh (137-159).
  - Refresh validation (Info!BB2) — [GST RULE, when in doubt] TIME-VARYING (638-685).
  - A report is "success" only if no error dialog was detected AND validation passed AND no exception (365-374);
    or Skip Refresh found a candidate (151-155).
  - Extraction always happens, even after a failed refresh (187-188).
- **Operational constants (all READ):** wait default 10 s (50); property retries 5 × 2 s (222-232);
  `xlMaximized = -4137` (227); open retries 3 × 3 s on "Call was rejected by callee" (239-247); load wait 5 × 1 s
  (253-255); AppActivate 3 × 1 s, then 1 s (264-276); keys `"%a"`, `"r"`, `"a"` with 1 s gaps (282-292);
  dialog checks 5 × 1 s (310-319); extended wait `max(10, wait_time)` (326); after-save 2 s (351); after-close 1 s
  (363); release wait 3 s (185); dialog title words (425-436), dialog classes (441), button words (506); window
  text > 5 chars, cut at 200 (469, 484); close pause 0.5 s (522); long-path threshold 218 (720, 795); second-open
  pause 3 s (750), third-open pause 2 s (758); error-text cuts `[:100]` (171, 176, 753) and `[:200]` (741);
  error strings "Open method of Workbooks class failed", "-2147352567" (170, 788), "Call was rejected by
  callee" (245); Info sheet `info` / cell `BB2` / error list (646, 656, 675); timestamp `%d%m%y_%H%M` (128).
- **Edge cases explicitly handled:** pywin32 missing (66-77); report missing (115-123); report not selected
  (89-94, 99-104); Excel busy on open (245-247); window activation failure (269-274); error dialogs (310-342);
  missing sheet / unreadable cell (570-576, 591-601); long paths (720-732); file already open in this Excel
  (709-712); file inaccessible (697-703).
- **Gaps, likely failures and silent bugs (beyond those above):**
  - Standard 12 comparison (rebuild_pipeline_FINAL.md:844-878), technical only: `Dispatch` instead of
    `DispatchEx` (852) at 218 and 536; refresh Excel visible (853 asks hidden) at 225; refresh by keys instead of
    `BackgroundQuery = False` + `RefreshAll()` + `CalculateUntilAsyncQueriesDone()` + `Refreshing` check
    (860-867); `IgnoreReadOnlyRecommended` only in the last resort (857); `ReadOnly` never checked (857-859);
    `Quit()` in `finally` — yes (856); `CoInitialize` in the worker — yes; `CoUninitialize` after the Excel
    function returns — no (872-875). HIGH.
  - If Dispatch attaches to the user's own Excel (MEDIUM-HIGH: the behaviour Standard 12 warns about), the
    user's Excel is maximised, keyed into, then quit with alerts off — as I recall Microsoft's documentation for
    `Application.Quit`, with DisplayAlerts False unsaved workbooks are closed without saving (MEDIUM; not
    looked up — no web access). Two Step 4 runs, Step 3 + Step 4
    (`core/excel_handler.py:127` also uses Dispatch), or the standalone extractor next to the main app would all
    share one Excel, and one run's `Quit()` ends the other's work (MEDIUM).
  - Error-dialog search is desktop-wide and auto-clicks OK/Yes/Close/Continue (MEDIUM-HIGH).
  - A failed or partial refresh still leaves a `_Refreshed_` file (copy2 at 164 happens before Excel opens it),
    later counted as "last refreshed" and reused by Skip Refresh with success True. HIGH.
  - Both Excel sessions run on a worker thread: COM objects are created and used only on that thread (good).
  - Very long runs: no cancel, no per-report time-out other than the fixed waits.
  - Two runs in the same minute → same refreshed name → copy2 overwrites the earlier copy (HIGH); if that file
    is open in Excel the copy fails (PermissionError → report error).
  - Suffix text is not validated (empty, `/`, `:` and other characters Windows forbids; `:` on NTFS even means
    an alternate data stream — NOT SURE what Excel would then do, LOW). RAN for empty and `/`.
  - On a network disconnect: glob returns nothing ("No ... report found") or copy/open fail; each failing report
    still costs the internal retry sleeps (≈ 5 s in `_safe_open_workbook`, ×2 for refresh and extraction) plus
    3 s (READ arithmetic).
  - Logging: client names, file names and every extracted figure are written to the log file via `logger.info`
    (56) — see extractor_main / main.py for where that file lives.
- **Excel workbook touchpoints:** `Dispatch("Excel.Application")` 218 (visible) and 536 (hidden);
  `Workbooks.Open` 737, 751, 766, 774, 777-783; `wb.Close(SaveChanges=False)` on same-named workbooks 749;
  `wb.Save()` 349 (the refreshed copy); `wb.Close(SaveChanges=True)` 361; `wb.Close(SaveChanges=False)` 399,
  622; `excel.Quit()` 406, 628; `Range(...).Value` reads 584 (mapped cells) and 656 (Info!BB2); keyboard
  "Refresh All" 282-292; `excel.Workbooks` iteration 746 (+ `utils/excel_com.py:26`). No openpyxl here.
- **Files and folders:** reads `<version folder>/<ITC|Sales>_Report_*.xlsx`; writes (copy) `<version
  folder>/<stem><suffix>.xlsx` (164) and Excel saves it (349, 361); never deletes, moves or renames anything
  (READ, grep). Excel itself creates `~$<name>.xlsx` owner files while a workbook is open.
- **Comments, docstrings vs code:** module docstring "Report processing with Power Query refresh" — true.
  `_refresh_power_query_simple` docstring "Simple refresh approach with proper error handling" (its docstring
  was dead text until 064c634 removed a stray `print` above it). "Open ... in read-only mode" (541, 543) —
  false. "Keep it visible to avoid issues" (225) — consistent with a key-press approach. Paste leftovers: "ADD
  THE NEW CODE HERE (notice the indentation - 3 levels)" (132), "ADD THIS: Wait for file to be fully released"
  (183), "Rest of your code remains same..." (295), "Increased from 0.5" (283, 288), "Increased from 3 to 5"
  (253). Import-guard comment 11-19 explains c463256 (accurate). `utils/excel_com.py` docstring's claim that
  this file "needs extended `\\?\` paths plus close-by-name retries" because it opens a NETWORK file "still
  locked from a just-finished Power Query refresh" — the code does contain those fallbacks (720-783); whether
  they are needed is not testable here.
- **Pattern notes:** errors converted to result dicts (never raised to the GUI) — opposite of
  `data_consolidator`; bare `except:` used 12 times in this file (269, 401, 408, 414, 444, 452, 471, 511, 624,
  630, 635, 767; RAN grep) plus twice in data_consolidator (201, 427), swallowing everything including
  KeyboardInterrupt; GUI logging via an injected callback;
  guarded optional import (same pattern as `gui/tabs/extract_tab.py:24-30`; `core/excel_handler.py:21-28` uses a
  platform check + try). Two separate Excel sessions per report (refresh, then extraction).
- **History (A/main):** created edac6a3 (736 lines); bd176bf (30-10-2025, shared) made wait time and suffix
  configurable and added per-report selection (`process_itc`/`process_sales`); 064c634 (23-07-2026) removed a
  stray `print`; ba2a753 (23-07-2026) moved the "already open" loop into `utils/excel_com.find_open_workbook`;
  2aaa2b1 (25-07-2026) switched Skip Refresh to `is_refreshed_copy_of` (and added the unused
  `DEFAULT_SUFFIX_PATTERN` import); c463256 (26-07-2026) guarded the pywin32 imports and added the
  `WIN32_AVAILABLE` refusal. B changed this file separately after the split (44 lines; B's readers).

### power_query_extractor/extractor_main.py (82 lines) - [CORE] (standalone entry point)

- **Purpose:** sets up logging, works out which target folder to open (from the main app's cache or older
  fallback files), then opens the standalone extractor window. READ + RAN.
- **Imports:** standard library `sys` (5, UNUSED — RAN AST), `logging` (6), `json` (7), `pathlib.Path` (8);
  project `from .gui.extractor_window import PowerQueryExtractorApp` (9) — a RELATIVE import, so the file only
  works as part of the package (RAN: `python power_query_extractor/extractor_main.py` →
  `ImportError: attempted relative import with no known parent package`; `python -m
  power_query_extractor.extractor_main` or `launch_extractor.py` work). No OS-specific modules.
- **Module-level code (runs at import):**
  - 12-19: `logging.basicConfig(level=INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[FileHandler('pq_extractor.log'), StreamHandler()])`. The log file name is RELATIVE, so it is
    created in the **current working folder** at import (RAN: importing from
    `/root/merge-scratch/run/A-R2/cwdtest` created `pq_extractor.log` there; launching `launch_extractor.py`
    from `/root/merge-scratch/run/A-R2/cwd2` created it in `cwd2`, not in the app folder). If that folder is not
    writable, the import fails with a permission error (→ `launch_extractor.py` shows "Failed to launch") (READ,
    HIGH). Because line 9 runs BEFORE line 12, the pywin32 warning from `report_processor` import is emitted
    before logging is configured and does not reach `pq_extractor.log` (RAN: the log lacked it).
    `basicConfig` is a no-op if logging was already configured (Python rule) — which is why
    `gui/tabs/extract_tab.py:8-10` deliberately avoids importing this module.
  - 21: `logger`.
- **class PowerQueryExtractor (24-72):**
  - **`__init__(self, target_folder=None)` (27-34):** stores the folder; if empty → `load_cached_folder()`;
    logs "Initializing Power Query Extractor with folder: <folder>" (34 — the folder path goes into the log).
    Callers: `launch_extractor.py:17`, `extractor_main.py:77`. Never called with a folder anywhere.
  - **`load_cached_folder(self)` (36-63):** tries, in order: `Path.home() / '.gst_organizer_cache.json'` (41 —
    the file the main app writes, `gui/utils/cache_manager.py:15, 45-46`), then three paths RELATIVE to the
    current folder: `gst_organizer_cache.json` (42), `.cache/app_state.json` (43), `config/last_session.json`
    (44). For each existing file: `json.load` (default text encoding), then the first non-empty of keys
    `'target_folder'`, `'target_path'`, `'output_folder'`, `'last_target'` (55) is returned; any error →
    warning and next file (59-60). Nothing found → "No cached folder found" and `''` (62-63). RAN: with a home
    cache holding `target_folder` it returned that; with the home cache removed it returned `output_folder` from
    a `gst_organizer_cache.json` in the current folder. **No code in A writes the three fallback files or the
    three fallback keys** (grep; `.gitignore:64` lists `config/last_session.json`) — they are legacy guesses.
    The folder returned is not checked for existence. Callers: `extractor_main.py:32`.
  - **`run(self)` (65-72):** `PowerQueryExtractorApp(self.target_folder)` then `app.mainloop()`; an exception
    is logged with traceback and re-raised. Callers: `launch_extractor.py:18`, `extractor_main.py:78`.
- **`main()` (75-78):** `PowerQueryExtractor().run()`. Callers: `extractor_main.py:82` (`if __name__ ==
  "__main__"`, 81-82) — reachable only via `python -m`.
- **Business rules:** none of GST. Cache precedence rule above.
- **Edge cases handled:** missing/corrupt cache files (try/except per file), empty values.
- **Gaps:** log location in the current folder and containing client names/figures (Standard 13,
  rebuild_pipeline_FINAL.md:1034-1041, asks for logs in `%LOCALAPPDATA%\[App Name]\` and for app files to be
  found from the app folder, never the current folder) — technical, HIGH; relative cache fallbacks depend on the
  current folder (HIGH); the cache is read twice in the standalone path (here, and `ExtractorPanel.load_from_cache`
  when this returns `''`) — harmless; a cached folder on a disconnected network drive makes the window's
  automatic scan run against it on the UI thread at start-up (MEDIUM: may freeze while Windows times out).
  No tests exist for this file (grep of tests/: only a docstring mention, `tests/test_extractor_standalone.py:8`).
- **Excel touchpoints:** none directly.
- **Files read/written:** writes `./pq_extractor.log` (append, every run); reads `~/.gst_organizer_cache.json`
  and possibly `./gst_organizer_cache.json`, `./.cache/app_state.json`, `./config/last_session.json`.
- **Comments vs code:** 39-40 "The main organizer writes here ... check it FIRST" — true (c1583d6). The design
  spec (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md`, "Window-level concerns" item 3) says the
  `basicConfig` "moves into the standalone entry point" — it did not move; it is still module-level here. HIGH.
- **Pattern notes:** the only module in my set that configures logging; `main.py:20-27` does the same for the
  main app with `gst_organizer.log` (also relative to the current folder).
- **History:** created edac6a3; only change c1583d6 (23-07-2026) added the home-folder path first (the old
  audit's "cache path mismatch" fix). B never changed this file after the split (`git diff 47a1492 B/main`
  empty), so B's copy lacks the home-folder line and still imports `.gui.extractor_window`, which B deleted
  (B's readers to confirm what that means for B).

### power_query_extractor/gui/__init__.py (1 line, no final newline) - [UI]

- **Purpose:** package marker; only `"""GUI components for Power Query Extractor"""`. No imports, no
  re-exports. Unchanged since edac6a3. READ.

### power_query_extractor/gui/extractor_window.py (1063 lines) - [UI]

- **Purpose:** the whole extractor user interface and its run logic: a parent-agnostic panel
  (`ExtractorPanel`) used both as the main app's "Step 4: Extract" tab and inside a thin standalone window
  (`PowerQueryExtractorApp`); scans the output tree for clients, lets the user tick ITC/Sales per client,
  runs `ReportProcessor` on a worker thread and `DataConsolidator` at the end. READ + RAN.
- **Imports:** standard library `tkinter as tk` (5), `tkinter.ttk, messagebox, filedialog` (6), `os` (7),
  `logging` (8), `threading` (9), `json` (10), `pathlib.Path` (11), `datetime` (12); project
  `..core.report_processor.ReportProcessor` (13), `..core.data_consolidator.DataConsolidator` (14),
  `..core.refresh_naming.is_refreshed_name` (15); function-local `..config.cell_mappings.EXTRACTION_OPTIONS`
  (486). OS-specific: `os.startfile` (973, Windows-only) — guarded by `try/except Exception` (972-975: on other
  systems it just logs "Could not open report"). Importing this module imports `report_processor`, whose
  pywin32 imports are guarded, so this module imports on any OS (RAN).
- **Module-level constants and globals:** `logger` (17). Operational literals: wait-time spin box 5-60,
  default 10 (270-274); suffix default `"_Refreshed_{timestamp}"` (297); scroll speed 3 units per wheel notch
  (471; `gui/tabs/setup_tab.py:138` names the same value `SCROLL_UNITS_PER_NOTCH`); window 1000×700, minimum
  900×600 (1043-1044); folder patterns `"Annual Statement-*"` (623) and `"Version-*"` (641) (duplicating
  `utils/constants.py:82, 84` instead of importing them); report prefixes `itc_report_` / `sales_report_` with
  `.xlsx` (36-39, 66-70); display time format `"%Y-%m-%d %H:%M"` (80); log time `"%H:%M:%S"` (1024).

- **`_classify_version_reports(version_folder)` (20-42):** one `os.scandir` of the version folder; `has_itc`
  = any name (lower-cased) starting `itc_report_` and ending `.xlsx`; `has_sales` likewise `sales_report_`
  (the `elif` means one name cannot set both). Refreshed copies count too. Any `OSError` (folder missing, not a
  folder, network error) → `(False, False)` silently. Returns `(bool, bool)`. Callers: `extractor_window.py:651`.
  Not unit-tested directly (only through scans in `tests/test_extractor_panel.py`); the design spec's claim
  that it is unit-tested is only half true (`_latest_refresh_times` is). READ.
- **`_latest_refresh_times(version_folder, suffix_pattern=None)` (45-82):** one `os.scandir`; skips names
  that `is_refreshed_name(name, suffix_pattern)` rejects (64); for `itc_report_*` / `sales_report_*` keeps the
  newest `entry.stat().st_mtime`; `OSError` → swallowed (74-75); returns two strings formatted
  `"%Y-%m-%d %H:%M"` (local time of the file's date-modified) or `None` each (77-82). Because the refreshed
  copy is made with `shutil.copy2` (keeps the original's time) and only re-dated when Excel saves it, a
  refresh that failed before saving shows the ORIGINAL report's time as "last refresh" (READ, HIGH). Callers:
  `extractor_window.py:787`; tests `tests/test_refresh_naming.py:102, 112, 121`.
- **`configure_extractor_styles()` (85-95):** `ttk.Style().theme_use('clam')`; `TNotebook` background
  `#f0f0f0`; `TNotebook.Tab` padding `[20, 10]`. Application-wide; called only by the standalone window
  (1046) — docstring explains why. RAN: theme is `clam` in the standalone window.

- **class ExtractorPanel(ttk.Frame) (98-1029):**
  - **`__init__(self, parent, target_folder=None, show_header=True, show_status_bar=True, auto_load=True)`
    (107-139):** creates `ReportProcessor()` (112) and `DataConsolidator()` (113); `client_vars = {}`;
    `cache_file = Path.home() / '.gst_organizer_cache.json'` (118); `is_processing = False`; stores flags;
    `_init_vars()`; `create_widgets()`; then: `auto_load=False` → only pre-fill the folder box (132-134);
    `auto_load=True` with a folder → pre-fill and `scan_folder()` (135-137); `auto_load=True` without →
    `load_from_cache()` (138-139). `self.target_folder` (111) is never read again (dead attribute). `self.level1_folder`
    is NOT initialised here — it first exists after a successful scan (632). Callers:
    `gui/tabs/extract_tab.py:74-80` (embedded: header off, status bar off, auto_load False);
    `extractor_window.py:1048` (standalone: all defaults); tests `tests/test_extractor_panel.py:39`.
  - **`_init_vars` (141-145):** `folder_path` StringVar, `status_var` StringVar "Ready", `progress_var`
    DoubleVar. Caller: 124.
  - **`create_widgets` (147-182):** optional banner (150-161): blue `#0078D4` frame 60 px high with
    "🔄 Power Query Report Extractor" (Segoe UI 18 bold, white); a `ttk.Notebook` (164-165) with two sub-tabs
    "Setup" (168-170) and "Processing" (173-175), white backgrounds; optional status bar (178-179); binds
    `<<NotebookTabChanged>>` → `on_tab_changed` (182). Caller: 125.
  - **`create_setup_tab` (184-481):** layout (two columns):
    - LEFT column, fixed width 400 px (`pack_propagate(False)`, 191-193):
      - LabelFrame "Target Folder" (196-204): label "Folder Path:"; entry bound to `folder_path` (width 35);
        buttons "📁 Browse" (grey `#6c757d`, `command=self.browse_folder`, 226-235) and "🔍 Scan" (green
        `#28a745`, bold, `command=self.scan_folder`, 237-246).
      - LabelFrame "Processing Options" (249-257): "Power Query Wait Time:" with a `Spinbox` 5-60, default 10
        (270-279) and hint "seconds (Min: 5, Max: 60)"; "Refreshed File Suffix:" entry, default
        `_Refreshed_{timestamp}` (297-304), hint "Use {timestamp} for datetime"; check box "Skip Refresh\n(Use
        existing refreshed files)" bound to `skip_refresh_var`, `command=self.on_skip_refresh_toggle` (315-325).
      - Button "🚀 Start Processing" (328-340): blue, bold 11, `state='disabled'` initially, hand cursor,
        `command=self.start_processing`.
    - RIGHT column (343-481): LabelFrame "Select Clients & Reports"; buttons "✓ Select All" (teal `#17a2b8`,
      → `select_all_clients`), "✗ Deselect All" (grey, → `deselect_all_clients`), "ITC Only" and "Sales Only"
      (blue, → `select_itc_only` / `select_sales_only`) (361-403); header row "Client Name" (width 25), "ITC"
      (8), "Sales" (8), "Last Refresh Status" (410-420); a `Canvas` + vertical `ttk.Scrollbar` holding
      `scrollable_frame` (422-432, 480-481); `<Configure>` on the frame updates the scroll region (426-429,
      a lambda — not in a loop, so no late-binding issue).
    - Mouse wheel (448-478, from bf2f1dd): `self._wheel_accum = 0.0`; inner `_in_scroll_area(widget)` walks
      `.master` links up to the canvas (450-456); inner `_on_mousewheel(event)` (458-476): ignore if the canvas
      no longer exists / pointer not over this canvas / nothing to scroll (return None so others get the
      event); otherwise accumulate `-delta/120*3` and scroll whole units; return "break". Bound once with
      `canvas.bind_all("<MouseWheel>", ..., add='+')` (478) — application-wide, adding to (never replacing)
      other handlers. `<MouseWheel>` is the Windows/macOS event; Linux wheels send Button-4/5 → no wheel
      scrolling on Linux (Windows-first assumption, LOW). Each new panel adds one more global handler; handlers
      of destroyed panels stay registered but return early (459-463). Test guard:
      `tests/test_scroll_bindings.py` (RAN, passes).
    - Caller: 170.
  - **`on_skip_refresh_toggle` (483-492):** writes the tick into the module dict
    `EXTRACTION_OPTIONS['skip_refresh']` (487) — process-wide, read by every later report
    (`report_processor.py:134-135`, even mid-run, because the check box stays enabled during a run — RAN) —
    and logs "Skip refresh enabled - will use existing refreshed files" / "Skip refresh disabled - will refresh
    all files". The check box always starts unticked (315) whatever `EXTRACTION_OPTIONS` says; not saved.
    Caller: `command=` at 320.
  - **`create_processing_tab` (494-571):** LabelFrame "Processing Progress" with label "Ready to process" and a
    `ttk.Progressbar` 0-100 (500-528); LabelFrame "Processing Log" with a `Text` (height 20, word wrap,
    Consolas 9, background `#f8f9fa`) and scrollbar (531-556); log wheel handler (559-565, from 4c42758) scrolls
    `int(-delta/120)` units and always returns "break" (touch-pad deltas below 120 scroll 0 — the accumulation
    fix of bf2f1dd was not applied here); colour tags `info` black, `success` `#28a745`, `warning` `#ffc107`
    (amber on near-white — low contrast), `error` `#dc3545` (568-571). The log `Text` stays editable (user can
    type into it). Caller: 175.
  - **`create_status_bar` (573-585):** grey `#e0e0e0` strip, 25 px, label bound to `status_var`. Caller: 179
    (standalone only). **This is the ONLY widget showing `status_var`** — embedded, every `status_var.set(...)`
    (617, 627, 663, 667, 952, 1019) is invisible (READ, HIGH).
  - **`load_from_cache` (587-599):** if `~/.gst_organizer_cache.json` exists → `json.load` → `target_folder`
    → if non-empty: fill the box, log "Loaded folder from cache", `scan_folder()`. Errors → logger warning only.
    Reads only the `target_folder` key (the main app writes it at `gui/main_window.py:364`). Caller: 139.
  - **`browse_folder` (601-609):** `filedialog.askdirectory(title="Select Target Folder", initialdir=<box> or
    None)`; if chosen → set and `scan_folder()` immediately. Caller: `command=` 229. Stays enabled during runs.
  - **`scan_folder` (611-668) — batch and version selection (runs on the UI thread):**
    1. Empty box → warning dialog "No Folder" / "Please select a folder first" (613-615).
    2. `status_var` "Scanning folder..." and log "Scanning folder for clients..." (617-618).
    3. `Path(box).glob("Annual Statement-*")` (622-623) — matches FILES too (no `is_dir` test, unlike
       `core/file_organizer.py:179`). None → log "No Annual Statement folders found" (warning), status "No
       processed folders found", return — **without clearing the previous client list or `level1_folder`**
       (RAN: list still showed the previous folder's clients and "Start Processing" stayed enabled).
    4. `latest_folder = max(..., key=st_mtime)` (631) → stored as `self.level1_folder` (632) BEFORE the client
       loop; log "Using folder: <name>".
    5. For each entry of that folder (638-659): a folder not starting with `_` → its `Version-*` entries (glob,
       files included) → if any: newest by mtime (648; comment 643-647 explains why not by name — c1583d6) →
       `_classify_version_reports` → append `{'name': folder name, 'path': folder, 'latest_version': Path,
       'has_itc', 'has_sales'}`. Client folders without any `Version-*` are silently left out. The `'path'` key
       is never read (dead). Order = directory order (alphabetical on NTFS; arbitrary elsewhere — RAN Linux
       listed CLIENTB before CLIENTA). No sorting.
    6. `display_clients(clients)`; log "Found N clients" (success); status "Found N clients" (661-663).
    7. Any exception → log "Error scanning folder: <e>" (error), status "Error scanning folder", logger error
       with traceback (665-668). RAN: a FILE named `Annual Statement-notes.txt` that is newer than the real
       folders → `NotADirectoryError` → "Error scanning folder".
    - RAN results: (a) parent folder with two batches → the newer-by-mtime batch chosen; (b) **pointing the box
      directly at an older `Annual Statement-...` folder → "No processed folders found"** (it looks for
      `Annual Statement-*` INSIDE the chosen folder), contrary to the bfcb24a commit message ("so an older
      'Annual Statement-...' batch can still be refreshed without touching Step 1") and the design spec ("Step 4
      behaviour"); (c) creating `~$...xlsx` in the OLDER batch folder made it the chosen batch on the next scan;
      (d) creating `~$ITC_Report_A.xlsx` in an older `Version-*` folder made that older version the chosen one;
      (e) nonexistent folder → "No processed folders found" (glob of a missing folder returns nothing).
    - Callers: `command=` 240; 137, 597, 609; `gui/handlers/processing_handler.py:169` (Step 3 "full" workflow
      auto-advance, UI thread via `after`); tests `tests/test_main_window_tabs.py:30-34, 69` (wrapped/called).
  - **`display_clients(clients)` (670-757):** destroys all rows and clears `client_vars` (673-675); no clients →
    italic grey "No clients found" (677-685) and return — **the Start button is NOT disabled** in that case.
    For each client: `get_refresh_status` (690); `itc_var` / `sales_var` BooleanVars pre-ticked when the report
    exists (693-694); stores `{'data', 'itc_var', 'sales_var', 'itc_status', 'sales_status'}` (696-702) — the
    two status keys are written (and rewritten at 776-777) but never read; row: name label (width 30 — header
    above is width 25, so columns do not line up), ITC and Sales check boxes (disabled when that report is
    absent), status label (grey, width 40) with `format_refresh_status` text; label reference kept (753).
    756-757: **if any clients → Start button enabled — also while a run is in progress** (RAN: state went
    `disabled` → `normal` after a Scan during a run). Callers: 661.
  - **`refresh_status_display` (759-782):** for every row, recompute the two times and update the label in
    place (keeps ticks); `TclError` (row destroyed) ignored. Runs on the UI thread at the end of each run
    (via `_reset_after_processing`), doing one directory listing per client there. Callers: 986; tests
    `tests/test_extractor_panel.py:90, 108, 120, 133`.
  - **`get_refresh_status(client)` (784-790):** `_latest_refresh_times(client['latest_version'],
    self._configured_suffix())`; any exception → `(None, None)`. Note: an unreachable folder therefore reads
    "Never refreshed" rather than "unknown". Callers: 690, 775.
  - **`_configured_suffix` (792-797):** current text of the suffix box, or None if the widget does not exist.
    Uses the box's CURRENT text, so after a restart (box back to default) files made with a custom suffix are
    recognised only through the legacy `_refreshed_` marker. Callers: 788.
  - **`format_refresh_status(itc, sales)` (799-808):** "ITC: <t> | Sales: <t>", "ITC: <t> | Sales: Never",
    "ITC: Never | Sales: <t>", or "Never refreshed". Callers: 740, 780.
  - **`select_all_clients` (810-816):** ticks ITC/Sales only where the report exists. `deselect_all_clients`
    (818-822): unticks both everywhere. `select_itc_only` (824-829): ITC where available, Sales off.
    `select_sales_only` (831-836): ITC off, Sales where available. Callers: `command=` at 364, 375, 386, 397.
  - **`on_tab_changed(event)` (838-843):** if the "Processing" sub-tab is selected during a run →
    `update_idletasks()` (UI thread). Effectively a no-op refresh. Caller: binding at 182.
  - **`start_processing` (845-883) (UI thread):** collects every client with at least one tick into a SHALLOW
    copy of its data dict plus `process_itc` / `process_sales` (849-857); none → warning "No Selection" /
    "Please select at least one report to process" (859-861); copies `wait_time_var.get()` and
    `suffix_pattern_var.get()` into the shared processor (864-865) — **an invalid spin-box text (e.g. "abc" or
    empty) raises `TclError` here (RAN) and the click silently does nothing** (Tk reports callback errors only to
    the console); values outside 5-60 typed by hand are accepted (RAN: 0 accepted; "12.7" → 12); switches to the
    Processing sub-tab (868); disables the Start button (871); `is_processing = True` (872); clears the log
    (875); starts a **daemon** `threading.Thread(target=_process_clients_thread, args=(selected,))` (878-883).
    No check of the main app's own Step 3 `is_processing`, nor of a run already going in this panel (only the
    disabled button, which a re-scan re-enables). Caller: `command=` 331.
  - **`_process_clients_thread(selected_clients)` (885-966) (worker thread):**
    1. Log header lines; **reads `wait_time_var.get()` and `suffix_pattern_var.get()` from the worker thread**
       (890-891) — Tk variable access off the UI thread, the one exception to the `_ui` rule (works with
       Python's thread-enabled Tcl, which forwards the call to the UI thread and waits — RAN under a real
       mainloop; from how Python's tkinter is built it would raise "main thread is not in main loop" if the
       main loop were not running — MEDIUM-HIGH).
    2. Per client (897-938): log "📋 Processing <name> - Reports: ITC, Sales"; progress `i/total*100` with
       "Processing <name>" (910-911); `processor.process_client(client, log_callback=self.log_message)` (914);
       append result; per report: success → "✅ <ITC|Sales> Report - Refreshed successfully" (shown also when
       Skip Refresh merely reused a file); failure → "⚠️ <...> - <error>" and, if data came back, "(Data
       extracted despite error)"; not selected → "⏭️ <...> - Skipped".
    3. 941-945: "📊 Creating consolidated report...", progress 95, `consolidator.create_report(results,
       self.level1_folder)` — reads `self.level1_folder` NOW, so a re-scan during the run redirects the output
       (RAN: saved into the re-scanned folder).
    4. 947-955: progress 100 "Processing complete"; logs "✅ Processing complete!", "📄 Report saved: <name>",
       "📂 Location: <folder>"; `status_var` "Processing complete" (then overwritten with "Ready" by the reset,
       1019); schedules `_offer_to_open_report`.
    5. 957-960: any exception (in practice: consolidated-report failure, or `self.level1_folder` missing) → log
       "❌ Error: <e>", logger error, error dialog "Error" / "Processing failed: <e>" (scheduled on the UI
       thread). All per-client results are lost in that case.
    6. `finally` (961-966): `is_processing = False` (plain attribute written from the worker); schedule
       `_reset_after_processing`; `processor.cleanup()` (no-op).
    - Caller: `threading.Thread(target=...)` at 878-880.
  - **`_offer_to_open_report(report_path)` (968-975) (UI thread):** "Complete" / "Processing complete!\n\nOpen
    the consolidated report?" → Yes → `os.startfile(report_path)` (Windows default program); failure logged only.
    Caller: `_ui` at 955.
  - **`_reset_after_processing` (977-986) (UI thread):** Start button → normal; progress 0, label/status
    "Ready"; `refresh_status_display()`. Caller: `_ui` at 963.
  - **`_ui(func, *args, **kwargs)` (989-1010):** `self.after(0, lambda: func(*args, **kwargs))` — schedules the
    call on the Tk thread; arguments are bound at call time (each call has its own lambda, no late-binding
    bug); swallows `TclError` (window gone) and `RuntimeError` ("main thread is not in main loop") so a failed
    screen update never kills the worker before its `finally`. Callers: 952, 955, 960, 963, 1014, 1025.
  - **`update_progress(value, message)` (1012-1014)** → `_ui(_apply_progress, ...)`; **`_apply_progress`**
    (1016-1019) sets the bar, the progress label and `status_var`. Callers: 911, 942, 947 (worker); 985 (UI);
    test `tests/test_extractor_panel.py:73`.
  - **`log_message(message, level='info')` (1021-1025):** timestamp `[HH:MM:SS]` taken immediately, then
    `_ui(_append_log, ...)`; **`_append_log`** (1027-1029) inserts with the colour tag and scrolls to the end.
    Used as the processor's `log_callback` (914). Callers: 490, 492, 596, 618, 626, 634, 662, 666, 888-892, 907,
    921-938, 941, 948-950, 958 and every `ReportProcessor._log`; test `tests/test_extractor_panel.py:72`.
    The log is never trimmed (grows during long runs) and is not saved to a file by the panel (the processor's
    `logger.info` copy goes to the log file).
- **class PowerQueryExtractorApp(tk.Tk) (1032-1063):** standalone shell.
  - **`__init__(target_folder=None)` (1039-1049):** title "Power Query Extractor", geometry 1000×700, minimum
    900×600, `configure_extractor_styles()`, then `ExtractorPanel(self, target_folder=target_folder)` with all
    defaults (banner, status bar, auto-load) packed to fill. No `WM_DELETE_WINDOW` (close-button) handler —
    closing during a run ends the program and kills the daemon worker without its `finally` clean-up (READ,
    HIGH). Callers: `extractor_main.py:68`; test `tests/test_extractor_standalone.py:31` (subprocess).
  - **`__getattr__(name)` (1051-1063):** forwards unknown attributes to the panel (so old code using
    `app.log_message` / `app.client_vars` still works). Side effect: it REPLACES `tk.Tk`'s own `__getattr__`,
    which normally forwards unknown attributes to the Tcl interpreter (RAN: `tk.Tk.__dict__` has `__getattr__`
    "Delegate attribute access to the interpreter object"; `plain_tk.eval('info patchlevel')` works, but
    `app.eval(...)` raises `AttributeError: 'ExtractorPanel' object has no attribute 'eval'`). Nothing in A uses
    such calls, so LOW.
- **UI elements summary (all READ; layout RAN under xvfb):** see the sub-items above. Dialogs: "No Folder"
  warning (614), "No Selection" warning (860), "Complete" yes/no (970-971), "Error" / "Processing failed: ..."
  (960), folder picker "Select Target Folder" (604). Status texts: "Ready", "Scanning folder...", "No processed
  folders found", "Found N clients", "Error scanning folder", "Processing <client>", "Creating consolidated
  report", "Processing complete". Enable/disable: Start button disabled at creation and at run start; enabled by
  any scan that finds clients (even mid-run) and after a run. Nothing else is ever disabled (Browse, Scan, wait,
  suffix, skip, select buttons, ticks all stay active during a run — RAN for the skip check box). There is no
  Stop button. Fonts: Segoe UI 18/11/10/9, Consolas 9 (Windows fonts; Tk substitutes elsewhere). Colours:
  `#0078D4`, `#28a745`, `#6c757d`, `#17a2b8`, `#e0e0e0`, `#f8f9fa`, `#666`, white backgrounds hard-coded on
  every `tk` widget (the main app's dark mode is handled by `gui/utils/dark_mode_manager.py`, A-R4's file).
- **Business rules in this file:** which batch is processed = newest `Annual Statement-*` by date modified
  (631); which version per client = newest `Version-*` by date modified (648); client folders starting with `_`
  are skipped (639); a report is offered only if a file `itc_report_*.xlsx` / `sales_report_*.xlsx` exists
  (36-39); default selection = every available report (693-694). None is a GST rule; they decide WHICH data set
  is reported (technical, author should know).
- **Edge cases explicitly handled:** empty folder box (613); no batch (625); no clients (677); unreadable folders
  (`OSError` in both helpers); destroyed rows (781-782); window gone while the worker reports (1003-1010);
  `os.startfile` failure (974-975); widget not yet created (796).
- **Gaps (in addition to those above):**
  - Embedded mode: status messages invisible (finding 9); the Setup sub-tab shows no feedback when a scan finds
    nothing (the log is on the other sub-tab). HIGH.
  - Scanning runs on the UI thread (glob + `stat` + two directory listings per client, over the network) — the
    window freezes for large or slow folders and on a dropped network drive (MEDIUM-HIGH).
  - Wait time, suffix and Skip Refresh are not persisted (no write to any cache; READ grep). HIGH.
  - Shared state touched from both threads without locks: `is_processing` (872, 962), `level1_folder` (632, 945),
    `EXTRACTION_OPTIONS['skip_refresh']` (487 vs `report_processor.py:135`), processor attributes (864-865 vs
    worker use, and `log_callback` per `process_client`). Benign individually; harmful together with the
    re-enabled Start button (RAN).
  - Two copies of the app (main app Step 4 + `launch_extractor.py`) on the same target: nothing prevents it;
    both drive Excel and both write `PQ_Extraction_Report_<minute>.xlsx` into the same folder (READ). MEDIUM-HIGH.
  - `messagebox`/`filedialog` calls pass no `parent`, so dialogs attach to the default root window (fine in
    both homes). LOW.
- **Excel touchpoints:** none directly; it triggers `ReportProcessor` (COM) and `DataConsolidator` (openpyxl),
  and opens the finished workbook with `os.startfile` (973).
- **Files/settings/caches:** reads `~/.gst_organizer_cache.json` key `target_folder` (118, 590-595); lists
  `<target>/Annual Statement-*`, `<batch>/<client>/Version-*`, `<version>/*` (623, 638, 641, 33, 61); nothing
  written by this file itself.
- **Comments and docstrings vs code:** module docstring "Two Tab Design" — true. `_classify_version_reports`
  docstring (21-29) — accurate. `ExtractorPanel` docstring "Nothing here touches window-level state" — true
  except the deliberate application-wide wheel binding (478, documented at 434-447). `_ui` docstring "Every
  worker-thread UI update goes through here" — two Tk variable reads (890-891) do not. "Open ... report"
  docstrings accurate. Comment 643-647 explains the mtime choice (c1583d6). Comment 129-131 (auto_load) accurate.
- **Pattern notes:** UI built with `tk` widgets inside a `ttk.Frame`; behaviour duplicated across the two homes is
  avoided by the panel split (good); worker-thread → UI marshalling via `after(0)` (same idea as
  `gui/handlers/processing_handler.py::_log_async`, per the design spec); error handling: UI-thread methods catch
  and log; dialogs only for "no folder", "no selection", completion and fatal run errors.
- **History (A/main):** created edac6a3 (601 lines, `PowerQueryExtractorApp(tk.Tk)` with everything);
  4c42758 (29-10-2025) log-wheel scrolling; bd176bf (30-10-2025) two-column Setup with wait time, suffix,
  ITC/Sales ticks and "Last Refresh Status" (338 of today's lines); 275d9a4 (23-07-2026) single-listing helpers;
  c1583d6 cache path + mtime version choice; c1f425e (25-07-2026) `_ui` marshalling; 2acffed timestamps updated
  after a run; a677670 the Panel/App split; 2aaa2b1 configured-suffix recognition; bfcb24a (26-07-2026)
  `auto_load`; bf2f1dd (27-07-2026) wheel binding rewrite. `git blame` (RAN): 388 lines from edac6a3, 338 bd176bf,
  88 a677670, 64 275d9a4, 56 c1f425e, 42 bf2f1dd, 36 2acffed, 20 2aaa2b1, 11 4c42758, 10 bfcb24a, 10 c1583d6.
  B deleted this file after the split (B is PyQt5).

### launch_extractor.py (30 lines, no final newline) - [CORE] (entry point)

- **Purpose:** the file a user runs to open the extractor on its own: makes the project folder importable,
  opens the standalone window, and shows an error box if start-up fails. READ + RAN.
- **Imports:** `sys` (6), `os` (7); inside `main`: `from power_query_extractor.extractor_main import
  PowerQueryExtractor` (14); in the error path: `tkinter as tk`, `tkinter.messagebox` (21-22). No OS-specific
  modules.
- **Module-level code:** `sys.path.append(<folder of this file>)` (10, comment "Add current directory to path" —
  it is actually the script's folder, not the current directory). APPENDED, i.e. searched LAST: when Python
  already puts the script folder first (normal `python launch_extractor.py`) this is redundant; under a Python
  that does not add the script folder (Standard 13 notes the portable embeddable Python does not) the project's
  `utils`/`core` packages would lose to any same-named installed package (LOW). Standard 13's `main.py` pattern
  uses `sys.path.insert(0, APP_DIR)`.
- **`main()` (12-27):** imports and runs `PowerQueryExtractor()` (folder from cache — comment 16 "it will
  auto-load from cache") → `.run()`. Any exception (import error, unwritable `pq_extractor.log` folder, window
  construction error, re-raised errors from `run`) → creates a NEW hidden `tk.Tk()` root (never destroyed),
  shows "Error" / `Failed to launch: <error>`, then re-raises (the traceback goes to the console — invisible
  under `pythonw`, and no log file of it is written unless logging was already configured). If tkinter itself
  is the problem, the error path fails too. Callers: line 30 (`if __name__ == "__main__"`, 29-30). No tests run
  this file (the standalone test constructs `PowerQueryExtractorApp` directly, despite its docstring line 8).
- **RAN:** `xvfb-run python launch_extractor.py` from `/root/merge-scratch/run/A-R2/cwd2` with an empty home:
  window started (killed by timeout after 8 s), `pq_extractor.log` written in `cwd2` with "No cached folder
  found" and "Initializing Power Query Extractor with folder: ".
- **Business rules:** none.
- **Gaps:** no single-instance guard (two extractors can run at once); the window's own errors after start-up
  are not covered (Tk shows callback errors only on the console); Standard 13 (rebuild_pipeline_FINAL.md:900-1061)
  expects a `.bat` launcher with `pythonw.exe -B -s`, `sys.excepthook` / `threading.excepthook` writing to
  `%LOCALAPPDATA%` — none of that exists here (technical comparison, HIGH).
- **Excel touchpoints / files:** none directly (see extractor_main for the log file).
- **Comments vs code:** docstring "Standalone launcher for Power Query Extractor / Run this file separately
  after GST Organizer processing" — still true; since bfcb24a the same panel is also Step 4 of the main app.
- **History:** unchanged on A/main since edac6a3. B rewrote it (32 lines changed; B's readers).

### utils/excel_com.py (29 lines) - [UTIL]

- **Purpose:** one shared helper that answers "is this file already open in this Excel?" for both COM users
  (Step 3's `core/excel_handler.py` and Step 4's `report_processor.py`). READ + RAN (tests).
- **Imports:** `pathlib.Path` (16). No OS-specific modules (it only uses an Excel object handed in).
- **Module-level:** docstring 1-15 (raw string) explaining that the two `_safe_open_workbook` methods
  deliberately keep DIFFERENT open-with-fallback strategies (Step 3 opens a freshly copied LOCAL temp file;
  Step 4 opens a NETWORK file that may still be locked), and "Do NOT merge those two strategies without
  validating BOTH flows against real Microsoft Excel".
- **`find_open_workbook(excel, abs_path)` (19-29):** loops over `excel.Workbooks`; returns the first workbook
  whose `Path(wb.FullName).absolute() == abs_path`; else `None`. `abs_path` must already be an absolute `Path`
  (both callers pass `Path(file_path).absolute()`: `core/excel_handler.py:326`, `report_processor.py:689`).
  On Windows, `Path` equality ignores letter case. A file reached by a mapped drive letter (`Z:\...`) and the
  same file by UNC (`\\server\share\...`) are NOT equal → not detected (MEDIUM). For a workbook stored on
  OneDrive/SharePoint, `FullName` is a web address → never equal (LOW). COM errors while looping are not caught
  here (they propagate to the caller's retry logic). Callers: `core/excel_handler.py:337`,
  `power_query_extractor/core/report_processor.py:709`; tests `tests/test_excel_com.py:23, 28, 32` (RAN: pass).
- **Business rules:** none.
- **Gaps:** "already open" detection works only within the Excel instance the caller holds; a workbook open in a
  DIFFERENT Excel process is invisible to it (READ, HIGH) — the Step 4 caller then opens it again (Excel may
  open it read-only, which nobody checks).
- **Excel touchpoints:** reads `excel.Workbooks` and `wb.FullName` (26-27). No open/save.
- **History:** added in ba2a753 (23-07-2026, A only), extracting an identical loop from both callers. Not in B.

---

## 3. Focus topics asked for in the assignment (summaries with pointers; details are in the file sections)

### 3.1 Where every mapped figure actually lands in the consolidated workbook (RAN)

Each data sheet: column A "Client Name", column B "Successfully Run" (✓/✗), then one column per mapping in
config order. "Actually filled from" = what the code does today when the report's status is success.

| Output sheet | Col | Header | Configured source | Actually filled from (RAN) |
|---|---|---|---|---|
| GSTR-3B Difference | C / D / E | Diff IGST / Diff CGST / Diff SGST | ITC `Diff GSTR-3B` J15 / K15 / L15 | same (correct) |
| GSTR-1 Difference | C / D / E / F | Diff Taxable Value / Diff IGST / Diff CGST / Diff SGST | Sales `Diff - GSTR-1` N1 / O1 / P1 / Q1 | same |
| Turnover GSTR-1 | C | Taxable Value | Sales `Diff - GSTR-1` I1 | same |
| Reco Diff | C / D / E | IGST Diff / CGST Diff / SGST Diff | ITC `Reco` C37 / D37 / E37 | same |
| Table 6J Match | C | Table 6J Match | ITC `T6 of R9` K29 | same |
| Table 8D Limits | C | Table 8D Limits | ITC `T8 of R9` G6 | same |
| IMS Reco | C / D / E | IGST / CGST / SGST | ITC `IMS_Reco` AA1 / AD1 / AG1 | **`Diff GSTR-3B` J15 / K15 / L15** whenever those were read (non-empty); configured cell only if J15/K15/L15 came back empty |
| Extra In Purchase | C / D / E | IGST / CGST / SGST | ITC `ExtraInPurchase` O1 / P1 / Q1 | **`Diff GSTR-3B` J15 / K15 / L15** (same condition) |
| Extra In GSTR2B | C / D / E | IGST / CGST / SGST | ITC `ExtraInGSTR2B` S1 / T1 / U1 | **`Diff GSTR-3B` J15 / K15 / L15** (same condition) |

All rows: if the report's status is not success → every column of that report = "Missing"; empty value →
"Missing"; non-zero number → format `#,##0.00`. [GST RULE] material (which cell feeds which column) — the
author decides which is intended; the notes only record the mismatch between config and code.

### 3.2 How a report is refreshed through Excel COM, and how values are read back (READ)

1. Find `<ITC|Sales>_Report_*.xlsx` in the client's chosen version folder; take the first listed
   (`report_processor.py:112-125`).
2. Copy it to `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` (or the custom suffix) with `shutil.copy2` (128-164) — or,
   with Skip Refresh, reuse the newest copy whose name looks refreshed (137-159).
3. **Excel session 1 (visible):** `CoInitialize`; `Dispatch("Excel.Application")`; Visible, alerts off,
   maximised (215-234); open the copy with fallbacks (`_safe_open_workbook`, 687-799); wait 5 s; bring Excel to
   the front (`AppActivate`); press **Alt+A, R, A** (Data → Refresh All) (258-293); wait `wait_time` s; look for
   error dialogs on the whole desktop for 5 s, reading and clicking them away (305-322); wait `max(10,
   wait_time)` s; final dialog check; check `Info!BB2` (638-685); `Save`; `Close(SaveChanges=True)`; `Quit`;
   `CoUninitialize` (394-415). Result: success only if no dialog was seen and BB2 passed.
4. Sleep 3 s (185).
5. **Excel session 2 (hidden):** new `Dispatch`; open the same copy; for each mapped sheet find it by exact name;
   read `Range(cell).Value` for each mapped cell; store under field name and bare cell ref; close without
   saving; `Quit`; `CoUninitialize` (527-636).
6. Result per report `{'status', 'data', 'extracted_values', 'original_file', 'refreshed_file'}`; the GUI collects
   one per client and passes the list to `DataConsolidator.create_report` (`extractor_window.py:945`).
- No `RefreshAll()`, `BackgroundQuery`, `CalculateUntilAsyncQueriesDone()` or `Refreshing` check anywhere (grep).
  The design spec keeps the key-press approach out of scope on purpose
  (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md`, "Out of scope").

### 3.3 Version-folder selection and the cache path (READ + RAN)

- Folder box → `glob("Annual Statement-*")` INSIDE it → newest by date modified (`extractor_window.py:623-632`) →
  each sub-folder not starting with `_` → newest `Version-*` by date modified (641-648). Before c1583d6 the
  version was chosen by name (wrong across months, per the commit message).
- Cache: the main app writes `~/.gst_organizer_cache.json` (`gui/utils/cache_manager.py:15, 45-46`; keys from
  `gui/main_window.py:360-370`, `target_folder` at 364; plus `recent_folders`, `cache_manager.py:43`). The
  extractor reads the same file and the same `target_folder` key (`extractor_main.py:41, 51-58`;
  `extractor_window.py:118, 590-595`). Path and key agree (HIGH). Embedded, the panel does not read the cache
  at all: its folder comes from Step 1's `target_folder` variable (`gui/tabs/extract_tab.py:76, 108-131`) and from
  `gui/handlers/processing_handler.py:166-168` after a "full" Step 3.

### 3.4 How the consolidated output is written (READ + RAN)

File `<chosen Annual Statement folder>/PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, openpyxl, new workbook every
run (`data_consolidator.py:175-179`). Sheets: "Summary" (first) + the 9 configured sheets. Summary: A1 title,
A3 generated time, A4 total clients, A5/A6 "ITC/Sales Reports Processed: x/n", A8 "Sheets Created:", A9+ list.
Data sheets: row 1 "<sheet> Report" (merged A1:C1), row 3 headers (blue), row 4+ one row per processed client,
columns as in 3.1. Same-minute rerun overwrites; failure loses the run's results (no fallback).

### 3.5 Refresh naming (suffix rules) — writer vs readers

| Role | Where | Rule |
|---|---|---|
| Writer | `report_processor.py:128-130` | `<original stem>` + suffix with `{timestamp}` (exact lower-case token) replaced by `DDMMYY_HHMM` + `.xlsx` |
| Reader: "has a report" | `extractor_window.py:36-39` | any `itc_report_*.xlsx` / `sales_report_*.xlsx` — refreshed copies included |
| Reader: "the original" | `report_processor.py:112-125` | first `<type>_Report_*.xlsx` in the listing — refreshed copies NOT excluded |
| Reader: "last refreshed" | `extractor_window.py:45-82` + `refresh_naming.is_refreshed_name` | name contains every literal piece of the CURRENT suffix box text (lower-cased), or `_refreshed_`; time = file's date modified |
| Reader: Skip Refresh | `report_processor.py:142-149` + `is_refreshed_copy_of` | name starts with the original stem AND is_refreshed_name; newest by date modified wins |

Disagreements: upper-case `{TIMESTAMP}` (reader accepts, writer does not replace); literal pieces that also occur
in normal names (`_r`, `_`, `_report`) make originals look refreshed (RAN); "original" picking does not exclude
copies (RAN on Linux); the suffix is not remembered between launches.

### 3.6 Threading and how the worker updates the Tk window (READ + RAN)

- UI thread: everything in `ExtractorPanel` except `_process_clients_thread`. Scans, cache reads and the final
  `refresh_status_display` run here (network I/O on the UI thread).
- One daemon worker per "Start Processing" (`extractor_window.py:878-883`): runs every client sequentially, both
  Excel sessions (COM initialised per session on this thread), and the openpyxl save.
- Worker → UI: `log_message`, `update_progress`, status, dialogs and the reset all go through `_ui` →
  `self.after(0, ...)` (989-1010), with values bound at call time. Exceptions: `wait_time_var.get()` and
  `suffix_pattern_var.get()` at 890-891 (direct Tk reads), and plain attributes `is_processing`, `level1_folder`
  shared without locks. RAN under xvfb with a stub processor: log/progress/dialog marshalling worked; a re-scan
  during the run re-enabled Start and changed the output folder.
- No cancellation; daemon thread dies without clean-up when the window closes (no close handler in standalone;
  main window checks only Step 3).

### 3.7 ExtractorPanel (embedded Step 4) vs the standalone window

| Aspect | Standalone (`launch_extractor.py` → `extractor_main.py` → `PowerQueryExtractorApp`) | Embedded (`gui/tabs/extract_tab.py:74-80`) |
|---|---|---|
| Window | own `tk.Tk`, "Power Query Extractor", 1000×700, min 900×600 | a tab "🔄 Step 4: Extract" in the main window |
| ttk theme | `clam` + notebook styling (`configure_extractor_styles`) | host's theme (styling deliberately not applied) |
| Banner | blue "🔄 Power Query Report Extractor" | suppressed (host shows its own strip, `extract_tab.py:54-65`) |
| Status bar / `status_var` | shown | suppressed → status texts invisible |
| Start-up scan | yes (`auto_load=True`): cache folder → immediate scan | no (`auto_load=False`): folder pre-filled from Step 1; user presses Scan, or Step 3 "full" mode triggers one |
| Folder source | `~/.gst_organizer_cache.json` (+ CWD fallbacks) | Step 1 `target_folder`, followed until the user overrides it here |
| Logging | `basicConfig` → `./pq_extractor.log` + console | main app's `basicConfig` → `./gst_organizer.log` + stdout (`main.py:20-27`) |
| Close during run | no handler: window closes, worker killed | `gui/main_window.py:344-355` asks only if Step 3 runs; else closes, worker killed |
| Attribute access | `__getattr__` forwards to the panel (replaces Tk's own forwarding) | direct |
| Same code | yes — one `ExtractorPanel` class (a677670) | yes |

---

## 4. Cross-file observations

### 4.1 Shared files, caches and settings — write path vs read path

| Shared thing | Written at | Read at | Agree? |
|---|---|---|---|
| `~/.gst_organizer_cache.json` | `gui/utils/cache_manager.py:45-46` (content `gui/main_window.py:360-370`, `cache_manager.py:43`) | `extractor_main.py:41, 51-58`; `extractor_window.py:118, 590-595`; `gui/main_window.py:121` | Yes: same path, JSON, key `target_folder` (HIGH) |
| `./gst_organizer_cache.json`, `./.cache/app_state.json`, `./config/last_session.json`; keys `target_path`, `output_folder`, `last_target` | NO WRITER in A | `extractor_main.py:42-44, 55` | Reader-only legacy (HIGH) |
| `EXTRACTION_OPTIONS['skip_refresh']` | default `cell_mappings.py:249`; run-time `extractor_window.py:487` | `report_processor.py:134-135` | Yes, but the check box ignores the file default (315) |
| `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` | `gui/handlers/processing_handler.py:297-298, 355-356` (`<type>_Report_<client>_<state>_<DDMMYY_HHMM>.xlsx`, timestamp from 186/294) | `report_processor.py:112`; `extractor_window.py:36-39, 66-70` | Prefix agrees; readers also match refreshed copies |
| Refreshed copies | `report_processor.py:128-130, 164` (+ Excel `Save` 349, `Close` 361) | `extractor_window.py:64`; `report_processor.py:112, 142-146`; `refresh_naming.py` | See 3.5 (partly inconsistent) |
| Folder names `Annual Statement-<DDMMYY HHMM>`, `Version-<DDMMYY HHMM>` | `core/file_organizer.py:103-107, 166-171` via `utils/constants.py:82, 84`, timestamp `core/file_organizer.py:43` / `utils/helpers.py:25-27` | `extractor_window.py:623, 641` (literal patterns, not the constants) | Names agree; extractor does not filter files out (organizer does, `file_organizer.py:179`) |
| "Latest batch" rule | organizer rerun/resume: newest dir by mtime (`core/file_organizer.py:176-185`) | extractor: newest entry by mtime (631) | Same rule (both move when Excel lock files appear) |
| `Info!BB2` | not written by any code in A (template / Power Query) | `report_processor.py:656` | Cannot verify without the templates |
| `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` | `data_consolidator.py:175-179` | `extractor_window.py:973` (opens it) | Yes |
| Log files | `./pq_extractor.log` (`extractor_main.py:16`); `./gst_organizer.log` (`main.py:24`) | people; `.gitignore:40-42` ignores both | Both relative to the current folder |

### 4.2 Inconsistent patterns

- Error handling: `report_processor` converts every error into a result dict (never raises); `data_consolidator`
  logs and re-raises; `extractor_window` catches and logs, dialogs only at run end.
- Optional-import guards: `report_processor.py:20-36` (`except ImportError`), `gui/tabs/extract_tab.py:24-30`
  (`except Exception`, message always says pywin32), `core/excel_handler.py:21-28` (platform check + try; A-R1).
- Sheet-name matching: exact in `_extract_data` (565), case-insensitive in `_validate_refresh` (646).
- Timestamp formats: folders `DDMMYY HHMM`; reports and refreshed copies `DDMMYY_HHMM`; consolidated file
  `YYYYMMDD_HHMM`; client list `YYYY-MM-DD HH:MM`; summary sheet `YYYY-MM-DD HH:MM:SS`; log `HH:MM:SS`.
- Duplicated literals instead of shared constants: default suffix (`refresh_naming.py:14`, `report_processor.py:51`,
  `extractor_window.py:297`); folder patterns (`utils/constants.py:82, 84` vs `extractor_window.py:623, 641`);
  wheel speed 3 (`gui/tabs/setup_tab.py:138` named, `extractor_window.py:471` literal).
- Mouse wheel: the client list accumulates small deltas (471-475), the log box does not (561).
- Both COM users use `Dispatch`, not `DispatchEx` (`report_processor.py:218, 536`; `core/excel_handler.py:127`).
  Step 3 goes further and first asks for the user's running Excel on purpose
  (`core/excel_handler.py:123-125`, `GetActiveObject("Excel.Application")`, "Using existing Excel instance"), so
  while Step 3 runs, Step 4's `Dispatch` and Step 3 may well be driving the same Excel (READ; MEDIUM on effect).
- Thread-to-UI marshalling: consistent through `_ui`, except 890-891.

### 4.3 Dead code found (all READ; confirmed by whole-project grep)

`data_consolidator._write_section` (250-352), `data_consolidator._create_details_sheet` (354-414); the
`'description'` branch (53-65); leftover header write (77-80) and duplicate `current_row` (83); `_calculate_value`
dormant (no calculations configured) and its unused `workbook` parameter; `report_processor` unreachable return
(383-386), effectively dead `except` (169-181), `excel_instance` (48), `cleanup()` no-op (801-804), unused
imports `os` (5) and `DEFAULT_SUFFIX_PATTERN` (38) and the module-level `win32api` (only the local import at 762
is used); `EXTRACTION_OPTIONS['auto_refresh_if_missing']`, `['show_timestamp']`; `FIELD_DISPLAY_NAMES` (only dead
reader); `power_query_extractor.__version__`; `ExtractorPanel.target_folder` (111); client dict key `'path'`
(655); `client_vars[...]['itc_status'/'sales_status']` (700-701, 776-777); result keys `'timestamp'`,
`'original_file'`, `'refreshed_file'`; `extractor_main` unused `sys`; `data_consolidator` unused `Path`;
`on_tab_changed` does nothing useful. Test coverage: no test touches `data_consolidator`, `extractor_main`,
`launch_extractor`, `_process_report`, `_refresh_power_query_simple`, `_extract_data`, `_validate_refresh`,
`_safe_open_workbook`, `start_processing` or `_process_clients_thread` (grep of tests/). RAN: the 7 test files
that do touch my code (`test_refresh_naming`, `test_extractor_panel`, `test_extractor_standalone`,
`test_excel_com`, `test_optional_pywin32`, `test_scroll_bindings`, `test_main_window_tabs`) → **113 passed**.

### 4.4 A vs B quick facts for my files (READ, git)

`cell_mappings.py`, `power_query_extractor/__init__.py` and the three sub-package `__init__.py` files are
identical in A/main and B/main (unchanged since 47a1492/edac6a3). `refresh_naming.py` and `utils/excel_com.py`
exist only in A. B deleted `gui/extractor_window.py` and changed `launch_extractor.py`, `data_consolidator.py`
and `report_processor.py` after the split; B's `extractor_main.py` is the split-point version (it still imports
the deleted `.gui.extractor_window`) (B readers cover those).

---

## 5. RAN log (all in `/root/merge-scratch/run/A-R2/`; nothing written to A-copy, B-copy or /home/user)

1. `python -B -c` import of `cell_mappings` → printed `CELL_MAPPINGS`, `FIELD_DISPLAY_NAMES`, options, sheet-name
   lengths, `__version__` (section 2 tables).
2. `DataConsolidator().create_report(...)` with fake results → `out/PQ_Extraction_Report_20261021_0405.xlsx`,
   reloaded with openpyxl and printed every sheet (cross-wiring, Missing-on-failure, layout, formats, widths).
3. `refresh_naming` probes with 11 suffix patterns and edge names.
4. `ReportProcessor._process_report` with `_refresh_power_query_simple`/`_extract_data`/`time.sleep` stubbed, in
   `out/tmp1fluxiqu/` (default, empty, `/sub`, skip, skip with `_R`, `{TIMESTAMP}`, missing Sales, `[1]` stem with
   Linux directory order).
5. `_validate_refresh` with fake sheets (11 values) and `_safe_open_workbook` with a fake Excel and a 232-char path.
6. `gui_probe.py` under `xvfb-run`: embedded panel, scan, invalid/0/12.7 spin-box values, worker run with a
   blocking fake processor, re-scan during the run, file named like a batch folder, Tk `__getattr__`.
7. `getattr_probe.py`: standalone window theme and `eval` forwarding.
8. `extractor_main` import in `cwdtest/` (log created in the current folder; cache precedence with fake HOME).
9. Direct script run of `extractor_main.py` (relative-import error); `launch_extractor.py` from `cwd2/` for 8 s.
10. `scan_probe.py` / `ver_probe.py`: batch choice by mtime, pointing at an older batch, `~$` lock-file effect on
    batch and version choice, nonexistent folder, empty box.
11. openpyxl edge values: range tuple, timezone-aware datetime, `True`, `-2146826246`, duplicate "Summary",
    32-character and `/` sheet titles.
12. AST scan for unused imports and exact line counts; `grep` for bare `except:`; `git blame` attribution.
13. `pytest -p no:cacheprovider` on the 7 related test files (copy in `A/`, fake HOME) → 113 passed in 2.17 s.

---

## 6. Possible client names (locations only)

None found. Checked: all lines of my 12 files at bf2f1dd; every version of `cell_mappings.py`,
`cell_mappings copy.py`, `data_consolidator.py`, `report_processor.py`, `extractor_main.py`,
`extractor_window.py`, `launch_extractor.py` at edac6a3, bd176bf, 47a1492 and A/main (grep for company-style
words and GSTIN-shaped strings); the messages of the 16 A/main commits touching these files. Test fixtures that
call my code use only the generic "ABC" / "ABC-DL" (e.g. `tests/test_extractor_panel.py:28-31`). I did not open
`gst_organizer.log`.

---

## 7. Questions for the author

1. For output sheets "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B", are the configured cells
   (`IMS_Reco` AA1/AD1/AG1, `ExtraInPurchase` O1/P1/Q1, `ExtraInGSTR2B` S1/T1/U1) the ones you intend? Today the
   code fills those columns from `Diff GSTR-3B` J15/K15/L15 instead. [GST RULE — your call which cells.]
2. When a refresh is judged "failed" (Info!BB2 check or an error dialog), should the extracted figures still be
   shown (perhaps flagged), or replaced by "Missing" as now?
3. Is `Info!BB2` filled by Power Query during the refresh (so it is empty in an un-refreshed report), in BOTH the
   ITC and the Sales templates? If it already holds a value before refresh, the check cannot detect a failed
   refresh.
4. Do you or your staff keep your own Excel workbooks open while Step 4 runs, or work in other programs at the same
   time? (Step 4 may take over a running Excel, and it types Alt+A, R, A into whatever window is in front.)
5. Which language is Office installed in on the office PCs? (The Alt+A, R, A key sequence is the English ribbon.)
6. Do you ever need Step 4 to process an OLDER `Annual Statement-...` batch? Pointing at it does not work today; the
   newest batch by "date modified" is always used.
7. Should the wait time, the refreshed-file suffix and the Skip Refresh choice be remembered between sessions?
   Have you ever changed the suffix from `_Refreshed_{timestamp}`?
8. Are the output folders on a Windows server share (NTFS) or on a NAS device? (Directory order decides which
   `ITC_Report_*.xlsx` is treated as the original.)
9. Have the template cell positions (J15, K29, G6, AA1, ...) or sheet names ever changed between template versions,
   and do old batches use older templates? (TIME-VARYING.)
10. Does `T6 of R9!K29` ("Table 6J Match") hold a number or a text such as Match/Mismatch? (Only affects
    formatting — technical.)
11. Would you want the consolidated workbook to show WHY a client failed (the never-used "Extraction Details" sheet
    suggests this was once intended)?
12. Is `launch_extractor.py` (the separate window) still used, or only Step 4 inside the main app?

---

## 8. Limits (not INCOMPLETE — every assigned line was read)

Could not be executed here (Linux, no Excel, no pywin32) and are therefore READ-only with the stated confidence:
whether `Dispatch` attaches to a running Excel; how pywin32 returns error cells and dates; whether
`Workbooks.Open` accepts `\\?\` paths; the Alt+A, R, A key tips; Excel/Office dialog class names; NTFS directory
order and directory "date modified" behaviour on the office share; behaviour of a `:` in the suffix. How to check
each is given next to the finding.
