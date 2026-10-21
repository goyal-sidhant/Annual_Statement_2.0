# B-R1 working notes - project B (AnnualStatement_oldv2, snapshot b0e7d0e) - core/* and utils/*

Reader: B-R1. Files covered (read in full, every line, with the Read tool): `core/__init__.py`, `core/excel_handler.py`,
`core/file_organizer.py`, `core/file_parser.py`, `utils/__init__.py`, `utils/constants.py`, `utils/helpers.py`.
Other files were read only to find callers or understand a call (listed where used).

## How to read these notes

- **READ** = I read the code and it says this. **RAN** = I executed the code (copies only, in
  `/root/merge-scratch/run/B-R1/`) and saw this result. Scripts: `/root/merge-scratch/run/B-R1/scripts/t_parse.py`,
  `t_helpers.py`, `t_flow.py`, `t_modes.py`, `t_excel.py`, `t_fakecom.py`, `t_resume_name.py`. All test file names
  are invented placeholders (`ClientA`, `ClientB Private Limited`, `ClientC`, `Delhi`, `Goa`, `Tamil Nadu`); no client data used.
- Confidence for every inference: **HIGH / MEDIUM / LOW**.
- **[GST RULE]** = code decides something the author owns (which return/file/cell a thing belongs to, state-code logic,
  periods, required returns...). I only describe these; I do not judge them. **TIME-VARYING** = a value that may change over time.
- Runs used Python 3.14.6 + openpyxl 3.1.5 on Linux (`/root/merge-scratch/venv-B/bin/python`). There is no Excel and no
  pywin32 here, so the Excel-through-COM code was traced by reading, plus a simulation with fake Python objects
  standing in for Excel (`t_fakecom.py`) - that simulation proves only the Python control flow (which calls are made in
  which order), NOT how real Excel reacts.
- Line counts: the Read tool shows one more line than `wc -l` for every file here, because none of these files ends with a
  line-break after its last line. I give the Read-tool count (the last line number you can cite).

### Plain-English glossary (terms used below)

- **regex / regular expression**: a text pattern used to recognise file names (for example "starts with `GSTR3B-`").
- **COM**: Windows' mechanism that lets one program (here Python) drive another (here Excel) remotely.
- **pywin32 / win32com / pythoncom / win32api**: the Python add-on package (and its modules) that provides COM and other Windows calls.
- **openpyxl**: a Python library that reads and writes .xlsx files directly, without Excel.
- **Power Query (PQ)**: Excel's built-in data-import feature; its queries are stored inside the workbook.
- **glob**: a wildcard file search such as `*.xlsx`.
- **stem / suffix**: for `Report.xlsx`, stem = `Report`, suffix = `.xlsx`.
- **mtime**: a file's or folder's "last modified" time stamp kept by the operating system.
- **dict / defaultdict**: Python's key-to-value table; a defaultdict creates an empty entry automatically when a missing key is used.
- **worker thread / QThread**: a background line of execution so the window does not freeze; Qt "signals" carry messages back to the window.
- **exception**: Python's error object; "caught/swallowed" = the code handles it and carries on.
- **8.3 short path**: old Windows short form of a long path (e.g. `C:\PROGRA~1\...`).
- **UNC path**: a network path such as `\\server\share\folder`.
- **dead code**: code that is defined but never used/called anywhere in the snapshot.

## Git history of the assigned files (B/main)

| File | Commits that touched it (newest first) |
|---|---|
| core/__init__.py | edac6a3 only (unchanged since initial commit 31-08-2025) |
| core/excel_handler.py | c258047 (07-02-2026), edac6a3 |
| core/file_organizer.py | c258047, edac6a3 |
| core/file_parser.py | 7e09e5c (08-02-2026), c258047, 6237597 (29-10-2025), edac6a3 |
| utils/__init__.py | edac6a3 only |
| utils/constants.py | edac6a3 only |
| utils/helpers.py | 7e09e5c, c258047, 6237597, edac6a3 |

(READ, `git log B/main -- <file>`.) Only c258047 and 7e09e5c are after the split point 47a1492 (B-only); 6237597 and edac6a3 are shared with A.
A also changed core/excel_handler.py, core/file_organizer.py, core/file_parser.py, utils/__init__.py, utils/constants.py,
utils/helpers.py after the split and added utils/dependency_installer.py, utils/excel_com.py, utils/requirements_check.py
(`git diff 47a1492 A/main --stat -- core utils`, READ) - those are other readers' business; mentioned only so nobody assumes these files are identical in A and B.

- **6237597** (29-10-2025, shared): file_parser.py gained `self.progress_callback = None` and `_update_progress()` (callback errors
  now caught and logged) and `scan_folder` stores the callback; helpers.py `get_state_code` changed `from utils.constants import
  STATE_CODE_MAPPING` to `from .constants import STATE_CODE_MAPPING` (commit message: "Fixed circular import error"). (READ, `git show 6237597`.)
- **c258047** (07-02-2026, B-only, the PyQt5 rewrite): in these files it only removed `# ADD THIS/THESE` comments (excel_handler.py:39,
  file_organizer.py:15, file_parser.py import, helpers.py:15), deleted many `logger.info("[DEBUG] ...")` lines (excel_handler,
  file_parser) and replaced a few with `logger.debug(...)`, and replaced bare `except:` with `except Exception:` (excel_handler.py:62,
  113, 200, 210, 216), `except (KeyError, ValueError) as e: logger.warning(...)` (excel_handler.py:251-252, before: `except: pass`)
  and `except (TypeError, ValueError): pass` (excel_handler.py:534). No logic change in these files. (READ, `git show c258047 -- core utils`.)
- **7e09e5c** (08-02-2026, B-only; message is about Setup/Validation views, it does not mention these changes): file_parser.py
  `_process_file` rewritten to do one `stat()` and one 8-byte read per file, with three new "variation" reasons (stat failure is
  silently skipped; `< 1024` bytes = "File too small to be valid Excel"; bad signature = "Invalid or corrupted Excel file"; read
  error = "Cannot read file: ..."); new progress messages ("Finding Excel files...", "No Excel files found", "Found N files,
  processing...", "[i/N] name"); stopped importing `validate_excel_file` and `get_file_info`. helpers.py `find_excel_files` no longer
  calls `validate_excel_file` on each file (before: invalid/small files were silently dropped; after: they reach the parser and are
  reported as variations) and now skips names starting `~$`. (READ, `git show 7e09e5c -- core utils`.) Behaviour change: since
  7e09e5c, files under 1 KB or with a wrong signature appear in the "variations" list instead of vanishing (HIGH).

---

### core/__init__.py (13 lines) - [CORE]

- **Purpose**: package marker for `core` that imports and re-exports the three core classes.
- **Imports**: project only - `.file_parser.FileParser` (line 6), `.file_organizer.FileOrganizer` (7), `.excel_handler.ExcelHandler` (8). No OS-specific imports.
- **Module-level constants/globals**: `__all__ = ['FileParser', 'FileOrganizer', 'ExcelHandler']` (10); `__version__ = '3.0.0'` (12); `__author__ = 'GST File Organizer Team'` (13).
- **Classes/functions**: none.
- **Side effect that matters (HIGH, RAN)**: any `import core.<anything>` runs this file first, so importing only the parser
  (gui/controllers/scan_controller.py:8 `from core.file_parser import FileParser`) also imports `core.excel_handler`, which at import
  time (a) loads openpyxl, (b) runs `warnings.filterwarnings('ignore')` for the whole process (excel_handler.py:17) and (c) on Windows
  tries to import win32com/pythoncom and writes a root-logger message (excel_handler.py:21-28). RAN: after `import core.file_parser`,
  `sys.modules` contains `core.excel_handler` and `openpyxl`, and `warnings.filters[0]` is `('ignore', None, Warning, None, 0)`.
- **Callers**: nobody uses `from core import ...`; all callers import submodules directly (scan_controller.py:8, processing_controller.py:10-11). `__all__`, `__version__`, `__author__` are read by nothing (dead metadata).
- **Comments/version strings**: docstring "Core package for GST File Organizer v3.0" and `__version__ = '3.0.0'` disagree with the app's own "v4.0" (main.py:2, main.py:42, gui/__init__.py:11 `__version__ = '4.0.0'`) - stale since c258047 (HIGH, READ).
- **Pattern notes**: re-export facade pattern (`__init__` + `__all__`); new public classes must be added here by hand.

---

### core/excel_handler.py (627 lines) - [CORE]

- **Purpose**: makes each client's ITC and Sales report by copying a template workbook and writing folder paths and file names
  into cells of its "Links" sheet (through Excel COM on Windows when pywin32 imports, otherwise through openpyxl), and writes the
  run's summary workbook with openpyxl.

#### Imports
- Standard library: `os` (6, **unused**), `time` (7), `shutil` (8), `logging` (9), `tempfile` (10), `pathlib.Path` (11), `typing`
  `Dict, List, Optional, Any, Union` (12; `List` and `Optional` unused), `warnings` (13), `platform` (14).
- Third-party: `win32com.client` (23) and `pythoncom` (24) - **guarded**: only inside `if platform.system() == 'Windows':` (21) and
  `try/except ImportError` (22, 27). openpyxl `load_workbook, Workbook` (31), `openpyxl.styles Font, PatternFill, Alignment, Border, Side` (32),
  `openpyxl.utils.get_column_letter` (33) - unconditional.
- Project: `utils.constants EXCEL_TEMPLATE_MAPPING, EXCEL_SAFETY` (35); `utils.helpers sanitize_filename, get_safe_timestamp,
  ensure_path_exists, validate_excel_file, clean_windows_path, extract_filename_without_extension, get_short_path, get_state_code`
  (36-40). `sanitize_filename`, `get_safe_timestamp`, `validate_excel_file` are **imported but never used** in this file (HIGH, READ + AST check).
- Note: if pywin32 is installed but broken so that the import raises something other than `ImportError` (for example a DLL load
  error), line 22-25 would not catch it and the whole module import would fail (LOW - not tested; check by renaming a pywin32 DLL on a test PC).

#### Module-level constants and globals
- line 17 `warnings.filterwarnings('ignore')` - switches off **every** Python warning in the whole program, not just in this file
  (comment line 16 "Suppress warnings"). Runs at import time. (HIGH, READ + RAN.)
- line 20 `USE_WIN32COM = False`; set to `True` at line 25 only if on Windows and both imports succeed. Decides COM vs openpyxl in
  `create_report_from_template` (83) and gates `__del__` (59).
- line 26 `logging.info("Using win32com for Excel operations (Power Query support)")`; line 28 `logging.warning("win32com not available.
  Install with: pip install pywin32")` - these use the **root** logger (`logging.info`, not the module `logger`), unlike the rest of
  the file. If no logging was configured yet, `logging.info()` silently configures the root logger with a default console handler
  (standard Python behaviour), after which a later `logging.basicConfig(...)` does nothing. main.py configures logging at lines 16-23
  before importing the GUI, so through main.py this is harmless (HIGH, READ); launch_extractor.py configures no logging before
  importing the GUI (launch_extractor.py:13-14), so there the Windows-only line 26/28 would configure the root logger first (MEDIUM, READ; Windows-only, not run).
- line 42 `logger = logging.getLogger(__name__)`.
- Operational constants hard-coded in code (not in constants.py): temp folder name `"gst_excel_temp"` (52); `time.sleep(0.5)` after the
  template copy (125, comment "Give Windows time to complete the copy") and after closing (180, "Small delay"); default sheet name
  `'Links'` (133, 235); fallback sheet search substring `'link'` (145); COM error code string `"-2147352567"` (292); path-length threshold
  `218` characters (293); log truncation `value[:50]` (165) and `short_path[-50:]` (300); summary column width cap `50` and padding `+2` (537);
  header colour `"366092"`, white font `"FFFFFF"`, sizes 12 and 16 (352-359).

#### class ExcelHandler (45-627)

**`__init__(self)` (50-55)** - READ.
- Sets `self.temp_dir = Path(tempfile.gettempdir()) / "gst_excel_temp"` (52) and creates it (53, via `ensure_path_exists`; any
  failure raises out of the constructor). On Windows this is `%TEMP%\gst_excel_temp`; RAN on Linux: `/tmp/gst_excel_temp` created.
- `self.safety_config = EXCEL_SAFETY` (54); `self.excel_app = None` (55).
- **`temp_dir` is never used anywhere else** in the project and never deleted (grep, HIGH): a side-effect-only empty folder.
  **`safety_config` is never read** (grep, HIGH), so none of the EXCEL_SAFETY settings (retries, delays, prefixes, minimum size) are applied anywhere.
  **`excel_app` is never assigned anything but `None`** (grep: only lines 55, 59, 61, 207), which matters for lines 59 and 207 below.
- Caller: gui/controllers/processing_controller.py:54 (`excel_handler = ExcelHandler()`, once per processing run, inside the worker thread).

**`__del__(self)` (57-63)** - READ.
- If `USE_WIN32COM and self.excel_app`: calls `self.excel_app.Quit()`, swallowing any exception (62-63 `except Exception: pass`).
- Because `excel_app` is always `None`, the body never runs: effectively dead (HIGH). Called implicitly by Python when the object is garbage-collected.

**`create_report_from_template(self, template_path, output_path, data_mappings, report_type) -> bool` (65-93)** - READ + RAN.
- Logs at debug level (72). Converts both paths to `Path` (74-75).
- If the template does not exist: raises `FileNotFoundError("Template not found: ...")` (78-79) - NOT a `False` return. (RAN.)
- Creates the output's parent folder (81).
- If `USE_WIN32COM`: returns `_create_report_win32com(...)` (83-87). Else logs warning "Not on Windows - Power Query may not be
  preserved!" (90) and returns `_create_report_openpyxl(...)` (91-93). The message says "Not on Windows" even when the real cause is
  "on Windows but pywin32 missing" (HIGH, READ).
- Parameters: `report_type` is used as the key into EXCEL_TEMPLATE_MAPPING ('ITC' or 'Sales'). An unknown type gives an empty
  mapping, writes nothing, and still returns True (RAN: `'Other'` -> True).
- Caller: processing_controller.py:225 (inside `ProcessingWorker._create_report`, which wraps it in try/except at 223-239 and checks
  the template exists first at 215). Output path is built by the caller as `folders['version'] / f"{report_name}.xlsx"`
  (processing_controller.py:220) with report_name `ITC_Report_{sanitized client}_{state code}_{DDMMYY_HHMM}` or
  `Sales_Report_...` (processing_controller.py:120, 128) - always the `.xlsx` extension, whatever the template's extension.

**`_create_report_win32com(self, template_path, output_path, data_mappings, report_type) -> bool` (95-217)** - READ (+ fake-COM simulation, see t_fakecom.py).
Step by step:
1. `pythoncom.CoInitialize()` (107) - COM set-up for this thread (this runs in the processing worker thread).
2. Gets Excel (110-115): first `win32com.client.GetActiveObject("Excel.Application")` - **attaches to an Excel the user already has
   open** if there is one (log "Using existing Excel instance"); if that raises, `win32com.client.Dispatch("Excel.Application")`
   (log "Created new Excel instance"). Not `DispatchEx`.
3. Sets on that Excel: `Visible = False` (118), `DisplayAlerts = False` (119), `ScreenUpdating = False` (120). If the instance is the
   user's own Excel, **the user's Excel window is hidden**, and `Visible` is never set back to True anywhere (HIGH, READ; simulated).
4. `shutil.copy2(template, output)` (124) - copies the template file to the report's final name **before** opening it; silently
   overwrites an existing file of that name (same-minute re-runs produce the same name). Then `time.sleep(0.5)` (125).
5. Opens the copy with `_safe_open_workbook(excel, output_path)` (129).
6. Looks up `EXCEL_TEMPLATE_MAPPING[report_type]` (132; `{}` if unknown) and sheet name (133, default 'Links').
7. Finds the sheet (136-153): first exact name match ignoring case (137-140); else the first sheet whose name contains "link"
   (144-148, logs warning "Using sheet '<name>' instead of 'Links'"); else **adds a new sheet and names it 'Links'** (150-153, logs
   error "Links sheet not found! Creating one..."), so the run continues and reports success even though the template had no link sheet (HIGH; simulated: returns True).
8. Writes cells (159-167): for each `cell_ref -> data_key` of the mapping, `value = data_mappings.get(data_key, '')`; **only if the value
   is non-empty** (162) it sets `ws.Range(cell_ref).Value = str(value)` and logs `"  B2 = <first 50 chars>..."` at INFO (165) - so
   folder paths (which contain client names) go into the log. Empty values are **skipped, leaving whatever the template already had
   in that cell** (HIGH; RAN with the openpyxl twin, see below). Each cell write has its own try/except that logs and continues (166-167).
9. `wb.Save()` (173) then `wb.Close(SaveChanges=True)` (176) - saved twice; `wb = None` (177); `time.sleep(0.5)` (180).
10. Verifies only that the output file exists (183-189): returns True (logs size) or False ("Output file not found!"). No check that
    the cells were written, that it opens, or its size (EXCEL_SAFETY `verification_size_min` 10240 is not used).
11. `except Exception` (191-193): logs with traceback, returns False. **The template copy made at step 4 stays on disk under the
    report's name** (no clean-up anywhere) - so a failed run leaves an `ITC_Report_...xlsx`/`Sales_Report_...xlsx` that is just the
    unfilled template; the Extract tab later treats any `ITC_Report_*.xlsx` in a version folder as a report
    (gui/pyqt_main_window.py:461-462; power_query_extractor/core/report_processor.py:76-77) (HIGH, READ).
12. `finally` (195-217): if `wb` still set (only when an error happened after opening) `wb.Close(SaveChanges=False)` swallowing errors
    (197-201). If `excel` was obtained: `ScreenUpdating = True`, `DisplayAlerts = True` (204-205; not inside try - if Excel has already
    died, this raises from inside `finally` and replaces the original error/return value; MEDIUM, READ), then **because
    `self.excel_app is None` is always true (207), it always calls `excel.Quit()`** (209), swallowing errors. The comment at 206 says
    "Don't quit if Excel was already open", but nothing records whether the instance was the user's - so when step 2 attached to the
    user's own Excel, the program **quits the user's Excel**, with `DisplayAlerts` just switched back to True and the window still hidden
    (HIGH that Quit is called on the attached instance - simulated, fake log shows `('excel.Quit', 'USER-OWN-EXCEL')`; what real Excel then does with the user's unsaved workbooks is not tested - MEDIUM that it would show a hidden "save changes?" prompt or close the user's work).
13. `pythoncom.CoUninitialize()` (215) in the same `finally`, while the local variables `excel` and `wb` may still hold COM objects
    (they are released only when the function returns) (HIGH, READ).
- Calls per client: 2 reports -> 2 x (start/attach Excel, open, save, quit) with 1 s of fixed sleeps per report.
- No Power Query refresh happens here (refresh is done later by power_query_extractor) (HIGH, READ).
- Workbooks are opened with no extra arguments (`Workbooks.Open(file_str)`, 286): no `UpdateLinks`, `ReadOnly` or `IgnoreReadOnlyRecommended`
  argument, and no check of `workbook.ReadOnly` after opening. (Contrast: report_processor.py:725-731 passes `UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False` in its last attempt.)
- Caller: `create_report_from_template` (85).

**`_create_report_openpyxl(self, template_path, output_path, data_mappings, report_type) -> bool` (219-262)** - READ + RAN.
- Logs warning "Using openpyxl - Power Query content may be lost!" (224). Copies template to output (228). Loads it
  `load_workbook(str(output_path), data_only=False, keep_links=True)` (231) - `keep_vba` left at default False, so macros in an .xlsm
  template would not be kept (MEDIUM, from openpyxl's documented default; not RAN with a real macro file).
- Finds the sheet by exact name ignoring case only (238-242) - **no "contains link" fallback and no sheet creation**, unlike the COM
  twin (inconsistent). If not found, nothing is written but it still saves and returns True (RAN: template with sheet "Other" -> True;
  sheet "MyLinks" -> True with B2 empty; sheet "LINKS" -> written).
- Writes non-empty values only (246-250); errors `KeyError/ValueError` per cell are logged as warnings and skipped (251-252). A merged
  cell raises `AttributeError` ("'MergedCell' object attribute 'value' is read-only"), which is NOT in that tuple, so it falls to the
  outer `except` (260-262) -> logs "Fallback method error: ..." and returns False, **leaving the copied template on disk** (RAN).
- Saves over the output and closes (255-256); returns `output_path.exists()` (258).
- RAN: empty values leave the template's own cell content (a template value placed in B10 survived in the ITC report of a client
  with no IMS file; a value in B9 survived in every ITC report, since ITC does not map B9).
- RAN: an `.xltx` template copied to an `.xlsx` name keeps the "template" content type inside the file
  (`[Content_Types].xml` still says `template.main`) - Excel would likely refuse or mis-handle it (MEDIUM, not tested in Excel). The
  template picker only offers `*.xlsx *.xlsm` (gui/views/setup_view.py:612, 619), so this needs a typed-in path.
- **Reachability (HIGH, RAN on Linux + READ for Windows)**: the GUI cannot even start without pywin32, because
  `power_query_extractor/core/report_processor.py:11-15` imports `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui`
  unconditionally and is imported by `gui/controllers/extraction_controller.py:10` -> `gui/pyqt_main_window.py:27`. RAN:
  `import gui.pyqt_main_window` -> `ModuleNotFoundError: No module named 'pythoncom'`, while `import core.excel_handler` alone works
  with `USE_WIN32COM = False`. So through the app, this openpyxl fallback is reachable only in the odd case where `pythoncom` imports
  but `win32com.client` does not. In practice it is dead code inside the app (MEDIUM-HIGH).
- Caller: `create_report_from_template` (91).

**`_safe_open_workbook(self, excel, file_path)` (264-306)** - READ + simulated.
- Makes the path absolute (266); if missing, logs and raises `FileNotFoundError` (269-271).
- Loops over `excel.Workbooks` (274-277): if one is already open at the same full path, returns it (log "File already open, using
  existing workbook") - the caller then writes into it, saves and **closes it** (that would close a workbook the user had open; edge case, HIGH by reading).
- Logs path length (279-281). Tries `excel.Workbooks.Open(file_str)` (284-286).
- On failure (287-306): logs the first `args` item as "Error code" (288-289). If the message contains `"-2147352567"`, or contains
  the word "path", or the path is longer than 218 characters (292-293), it asks `get_short_path()` for the 8.3 short form and, if
  different, opens that (296-301); errors there are logged (302-303). Then a bare `raise` (306) re-raises the **original** error
  (simulated: original message propagated, caller returned False).
- On UNC/network paths and volumes without 8.3 names, `GetShortPathName` usually returns the same path, so no retry happens (LOW - not tested).
- Different from the same-named method in power_query_extractor/core/report_processor.py:638-747 (which also tries the `\\?\` long-path
  prefix, closing same-named workbooks, sleeps 3 s/2 s, and a parameterised Open) - two diverged copies of one idea.
- Caller: `_create_report_win32com` (129) only (report_processor has its own `_safe_open_workbook`, called at report_processor.py:197, 493).

**`create_summary_report(self, output_path, report_data) -> bool` (308-346)** - READ + RAN.
- Creates parent folder (314-315), new `Workbook()` (318), removes default sheet "Sheet" (321-322), builds styles (325), then 5 sheets
  in this order: Summary, Client Status, File Mapping, Errors, Variations (328-332); auto-fits column widths on every sheet (335-336);
  saves and closes (339-340); returns `output_path.exists()` (342). Any exception: logged with traceback, returns False (344-346).
- Always openpyxl (comment 317 "no Power Query needed").
- Caller: processing_controller.py:190, path `level1_folder / f"GST_Processing_Summary_{DDMMYY_HHMM}.xlsx"` (processing_controller.py:187-188),
  only if at least one client succeeded (processing_controller.py:184). Overwrites an existing summary of the same minute silently (openpyxl save; HIGH).
- RAN: produced sheets `['Summary', 'Client Status', 'File Mapping', 'Errors', 'Variations']`.

**`_create_styles(self) -> dict` (348-367)** - header: bold, size 12, white font `FFFFFF`, solid fill `366092`, centred; title: bold size 16,
centred; border: thin on all four sides. Caller: 325.

**`_create_summary_sheet(self, wb, data, styles)` (369-399)** - sheet "Summary"; A1 = "GST FILE PROCESSING SUMMARY" (title font, merged
A1:B1, 374-376); from row 3 (394), label in column A (bold) and value in B for these 12 rows (379-392), each read with `.get` and a default:
"Processing Date" (`timestamp`, ''), "Total Clients" (`total_clients`, 0), "Successful" (`successful_clients`), "Failed" (`failed_clients`),
"Total Files" (`total_files`), "Reports Generated" (`reports_generated`), "ITC Reports Created" (`itc_reports_created`, comment "# NEW"),
"Sales Reports Created" (`sales_reports_created`, "# NEW"), "Report Creation Errors" (`report_errors`, "# NEW"), "Processing Mode"
(`processing_mode`), "Target Folder" (`target_folder`), "Include Client Name in Folders" ("Yes"/"No" from `include_client_name`).
"Processing Date" holds the `DDMMYY HHMM` text from `get_timestamp()` (processing_controller.py:44, 59), not a real date (HIGH). Caller: 328.

**`_create_client_status_sheet` (401-428)** - sheet "Client Status"; headers "Client", "State", "Status", "Files", "Missing", "Complete %"
(406) styled (407-412); one row per entry of `data['clients']` (416): client, state, status, file_count, number of missing types,
completeness as text `f"{x:.1f}%"` (422 - text, not a number); thin borders on 6 columns (425-426). The caller only appends
**successful** clients to `clients`, always with `status: 'Success'` (processing_controller.py:136-147), so failed clients never appear here (HIGH, READ). Caller: 329.

**`_create_file_mapping_sheet` (430-456)** - sheet "File Mapping"; headers "Original File", "Client", "Type", "Destination", "Status" (435);
rows from `data['file_mappings']` (445-456). Caller fills it from each file result (processing_controller.py:162-169; "Client" is the
client key). Caller: 330.

**`_create_errors_sheet` (458-489)** - sheet "Errors"; headers "Time", "Client", "Type", "Error", "Action" (463); if no errors, A2 =
"No errors encountered" merged A2:E2 (474-476); else rows with `time`, `client`, `type`, `message`, `action` defaulting to "Review"
(478-489). Caller adds only per-client exceptions (processing_controller.py:175-180); report-creation failures are counted
(`report_errors`) but not listed here (HIGH, READ). Caller: 331.

**`_create_variations_sheet` (491-522)** - sheet "Variations"; headers "Filename", "Issue", "Expected Pattern", "Size", "Action" (496);
if none, A2 = "No file variations found" merged A2:E2 (507-509); else rows with Action always "Rename file" (516). **The caller always
passes `'variations': []`** (processing_controller.py:73, never filled), so this sheet always says "No file variations found" even when
the scan found variations (HIGH, READ; RAN with an empty list). Caller: 332.

**`_autofit_columns(self, ws)` (524-538)** - for each column, width = min(longest text length + 2, 50) (530-538); type/value errors per cell
ignored (534-535). Caller: 336. (A different `_autofit_columns` exists in data_consolidator.py:401.)

**`prepare_template_data(self, client_info, folders, report_type) -> Dict[str, str]` (540-627)** - READ + RAN.
- Debug log (546). `state_code = get_state_code(client_info['state'])` (550) is computed **outside** the try and **never used**; if
  `client_info` has no 'state', this raises `KeyError` to the caller (RAN).
- try (552-625):
  - **ITC** (553-588): `gstr3b_folder` = Windows-style path of `folders['gstr3b']` - **always**, even when the client has no GSTR-3B
    file (555; RAN). If any 'Annual Report' file: `annual_folder` = version folder path, `annual_filename` = stem of the **first**
    Annual Report file's `name` (558-563), else both '' (565-566). Same for 'GSTR-2B Reco' -> `gstr2b_folder` = itc folder,
    `gstr2b_filename` (569-577), and 'IMS Reco' -> `ims_folder` = itc folder, `ims_filename` (580-588).
  - **Sales** (590-622): 'Sales' -> `sales_folder` = sales folder, `sales_filename` (592-600); 'Annual Report' -> `annual_folder` =
    version folder, `annual_filename` (603-611); 'Sales Reco' -> `sales_reco_folder` = sales folder, `sales_reco_filename` (614-622).
  - Any other report type: returns `{}` (RAN).
- `except Exception` (624-625) logs "Error preparing template data" and returns whatever was filled so far (RAN: a file entry with no
  'name' -> returns only `{'sales_folder': ...}`; the error is invisible to the user and the report is still created) (HIGH).
- "First file" = first in the list built by the parser, i.e. alphabetical order of full paths from `find_excel_files` (helpers.py:262).
  Other files of the same type are copied but never referenced in the report (HIGH, READ; RAN: two GSTR-2B files -> first used).
- File names given are **stems without extension** (`extract_filename_without_extension`); whether the template re-adds ".xlsx" (and
  what happens for an ".xls" source) depends on the template's queries, which are not in the repository (LOW - check a template's Power Query source steps).
- The `name` used is the name **after** copying, because `FileOrganizer.organize_files` overwrote it with the sanitised file name
  (file_organizer.py:224-227) - but only for files copied successfully in this run; for a file that was skipped or failed, the
  original (unsanitised) name is sent while the folder holds the sanitised name (RAN, see file_organizer section).
- Caller: processing_controller.py:224.

#### Business rules in this file
- [GST RULE] (template contract; borderline - kept as a GST rule because it decides which source file feeds which part of the ITC/Sales report)
  The ITC report gets: GSTR-3B export folder, Annual Report folder+name, GSTR-2B Reco folder+name, IMS Reco folder+name; the Sales report
  gets: Sales folder+name, Annual Report folder+name, Sales Reco folder+name (553-622 with cell addresses from constants.py:96-120). TIME-VARYING
  (cell addresses depend on the template version).
- [GST RULE] When several files of one type exist, only the first (alphabetical) feeds the report (561-562, 572-573, 583-584, 595-596, 606-607, 617-618).
- [GST RULE] A missing source file means "leave the template's own cell value" (162, 249) rather than clearing the cell (described only).
- Annual Report files are expected in the version folder (560, 605), GSTR-2B and IMS in the "itc" folder (571, 582), Sales and Sales Reco in the
  "sales" folder (594, 616) - must agree with file_organizer.py:288-295 (it does; HIGH, READ).
- Summary workbook layout and texts (369-522) - not GST.
- Operational: 0.5 s waits (125, 180); 218-char path threshold (293); error code -2147352567 (292); column width cap 50 (537).

#### Edge cases explicitly handled
- Template missing -> `FileNotFoundError` (78-79). Output folder missing -> created (81, 315).
- Links sheet named in another case / containing "link" / absent (COM: 137-153; openpyxl: exact case-insensitive only 239-242).
- Workbook already open in the attached Excel (274-277). Long path / path error -> 8.3 short path retry (292-303).
- Per-cell write failures (166-167, 251-252). Non-string values converted with `str()` (163, 250).
- Empty error / variation lists in the summary (474-476, 507-509). Non-measurable cell values in autofit (534-535).

#### Gaps, likely failures and silent bugs
1. Takes over the user's own running Excel (111), hides it (118) and quits it at the end (207-209) - comment at 206 says the opposite (HIGH that it is coded so; simulated).
2. A failed report leaves the raw template copy under the report's final name, later picked up by the Extract tab (124 before 129; no clean-up) (HIGH).
3. Missing source files leave stale template cell values (162; RAN with openpyxl twin) (HIGH).
4. "Links" sheet missing -> COM path creates an empty "Links" sheet and reports success (150-153); openpyxl path writes nothing and reports success (244) (HIGH).
5. Report verification is only "file exists" (183-189, 258) (HIGH).
6. Same-minute re-run overwrites an existing report of the same name with no backup (124, 228) (HIGH, RAN: existing file overwritten, returned True).
7. `warnings.filterwarnings('ignore')` hides all warnings in the program (17) (HIGH).
8. Unused: `temp_dir` folder (52-53), `safety_config` (54), `excel_app` (55), `__del__` body (59-63), `state_code` (550), imports `os`, `List`, `Optional`, `sanitize_filename`, `get_safe_timestamp`, `validate_excel_file` (HIGH).
9. On a PC without pywin32, the user is never told in the window that the Power Query-safe path was not used: the only signals are log lines (28, 90, 224) (HIGH, READ; grep finds no `USE_WIN32COM` use in gui/).
10. `excel.ScreenUpdating = True` / `DisplayAlerts = True` in `finally` are not protected (204-205); if Excel crashed, the exception from here replaces the return value (the worker's `_create_report` catches it, so the client is not lost) (MEDIUM).
11. Opens with no `IgnoreReadOnlyRecommended` and never checks `workbook.ReadOnly`, so if Excel opens the copy read-only (for example someone else has it open) the write and `Save` would fail and the run returns False; in practice the preceding `shutil.copy2` (124) would usually already fail on a file that is open elsewhere (MEDIUM, READ). (No refresh happens here, so refresh-completion checks do not apply to this file.)
12. Thread model: COM work runs inside the processing QThread; `CoInitialize`/`CoUninitialize` are paired per report call (107, 215) (HIGH).
13. Network share: if the share drops during Save, the exception path returns False and leaves the half-made report (MEDIUM).
14. Two app copies at once on the same target and minute: both copy to the same report name; no locking (MEDIUM).
- Closure/lambda, mutable defaults, `is`/`==`, off-by-one: none found in this file (READ).

#### Excel workbook touchpoints
| Line | Action | Workbook |
|---|---|---|
| 107 | `pythoncom.CoInitialize()` | - |
| 111 / 114 | `GetActiveObject` / `Dispatch("Excel.Application")` | Excel application (possibly the user's own) |
| 118-120 | Visible/DisplayAlerts/ScreenUpdating = False | Excel application |
| 124 | `shutil.copy2(template, output)` | template -> report file |
| 274-277 | reuse already-open workbook | report file |
| 286 / 301 | `Workbooks.Open(path)` / open by 8.3 short path | report file |
| 152-153 | `Sheets.Add()`, rename to 'Links' | report file |
| 163 | `Range(cell).Value = str(value)` | report file, Links sheet |
| 173, 176 | `wb.Save()`, `wb.Close(SaveChanges=True)` | report file |
| 199 | `wb.Close(SaveChanges=False)` on error | report file |
| 204-205 | ScreenUpdating/DisplayAlerts = True | Excel application |
| 209 | `excel.Quit()` (always) | Excel application |
| 61 | `excel_app.Quit()` in `__del__` (never runs) | - |
| 215 | `pythoncom.CoUninitialize()` | - |
| 228, 231, 250, 255, 256 | copy2, `load_workbook(keep_links=True)`, cell writes, `save`, `close` | template -> report file (openpyxl) |
| 318, 322, 371/403/432/460/493, 339, 340 | `Workbook()`, remove 'Sheet', `create_sheet`, `save`, `close` | `GST_Processing_Summary_*.xlsx` |
No Power Query refresh in this file.

#### Files, folders, settings, caches and logs
- Creates `Path(tempfile.gettempdir()) / "gst_excel_temp"` (52-53) - never used, never removed.
- Writes the ITC/Sales report at the caller's `output_path` (124/228 then 173/255), and the summary workbook at the caller's path (339).
- Log messages through `logger` (module name `core.excel_handler`) and root `logging` (26, 28). INFO messages include full output paths
  and the cell values written (165) - i.e. folder paths containing client names go into the program's log file (`gst_organizer.log`,
  opened by main.py:20 relative to the current folder - B-R2's file).
- No settings or caches read or written.

#### Comments, docstrings, version strings
- Module docstring "Excel Handler for GST File Organizer v3.0 / WINDOWS VERSION - Uses Excel COM to preserve Power Query and all Excel
  features" (1-4): "v3.0" is stale (app says v4.0). "WINDOWS VERSION" is true for the COM path only.
- Comment 30 "Fallback to openpyxl if not on Windows" - the fallback also triggers on Windows without pywin32.
- Comment 206 "Don't quit if Excel was already open" - **contradicted by the code** (207 always true).
- Comment 125 "Give Windows time to complete the copy" - `shutil.copy2` has finished when it returns; the sleep is a guess-based wait (MEDIUM).
- Comments 386-388 "# NEW" mark the three report-count rows (added in edac6a3 or earlier; no history before it).
- Docstring 311 "Create summary report (doesn't need Power Query)" - agrees with code.
- No TODO/FIXME, no commented-out code.

#### Pattern notes
- Error handling: broad `except Exception` + `logger.error` + `return False` in public methods; `raise` for missing template. Mixed with
  root-logger calls at import (26, 28).
- Config access: imports constants directly (EXCEL_TEMPLATE_MAPPING); many operational numbers are inline literals instead.
- OS-specific imports: guarded here (platform check + try/except), but **unguarded** in power_query_extractor/core/report_processor.py:11-15 (inconsistent pattern across the project; HIGH).
- Threading: called from the processing QThread; no Qt objects touched here.

---

### core/file_organizer.py (409 lines) - [CORE]

- **Purpose**: creates the dated output folder tree for each client and **copies** (never moves) the client's recognised files into it,
  sanitising file names, with per-mode handling of files that already exist, and writes a per-client text report.

#### Imports
- Standard library: `logging` (6), `pathlib.Path` (7), `typing Dict, List, Optional, Any, Tuple, Union` (8; `Tuple` unused), `datetime` (9).
  `shutil` is imported locally inside `cleanup_on_error` (374).
- Project: `utils.constants FOLDER_STRUCTURE, PROCESSING_MODES` (11); `utils.helpers ensure_path_exists, safe_copy_file, create_backup,
  get_timestamp, sanitize_filename, safe_path_join, get_state_code, create_client_state_key` (12-16).
- No OS-specific modules.

#### Module-level constants and globals
- `logger = logging.getLogger(__name__)` (18). Hard-coded in code: prefix `"Annual Statement-"` for finding existing level-1 folders (162,
  duplicates constants.py:82 instead of reusing it); type-to-folder table (288-295); report file name pattern
  `"_Organization_Report_{client}_{state}.txt"` (313); separator line `"=" * 50` (318, 325, 336, 348); default max length 35 (30, 42).

#### class FileOrganizer (21-409)

**`__init__(self, target_folder, processing_mode='fresh', include_client_name=False, client_folder_settings=None, client_name_max_length=35)` (26-54)** - READ + RAN.
- Stores `target_folder` as Path (38), mode (39), `include_client_name` (40), `client_folder_settings or {}` (41),
  `client_name_max_length` if > 0 else 35 (42), and `self.timestamp = get_timestamp()` (43) - one `DDMMYY HHMM` stamp for the whole run.
- Tracking lists `created_folders`, `copied_files`, `errors` (46-48). Calls `_validate_inputs()` (51). Logs (53-54).
- Docstring (31-37) documents only `target_folder` and `processing_mode`.
- No mutable default argument problem (`client_folder_settings=None` then `or {}`) (READ).
- Caller: processing_controller.py:46-52 (once per run, in the worker thread) with `settings.target_folder`, `settings.processing_mode`,
  `include_client_name=settings.include_client_name_in_folders`, `client_folder_settings=self.client_folder_settings`
  (from `ProcessingState.client_folder_settings`, gui/pyqt_main_window.py:370), `client_name_max_length=settings.client_name_max_length`.
- **`client_folder_settings` is always an empty dict in B** (HIGH, READ): `ProcessingState.client_folder_settings` is created empty
  (models/app_state.py:83) and nothing in B ever writes to it (grep: only app_state.py:83, pyqt_main_window.py:370 and pass-through
  lines). In the old Tkinter GUI it was filled per client by a toggle (`git show c258047:gui/handlers/client_handler.py` lines 49,
  176-215, deleted in a62d3cc). So the "per-client include name" branch at line 126 is dead in B; only the global checkbox works.
  History note: that old Tk toggle used the key `f"{client}-{state name}"` (client_handler.py:178 in c258047) while this file looks up
  `f"{client}-{state_code}"` (118) - they would only have matched when the state was already written as a code (MEDIUM, READ of old tree).
- `processing_controller.py:44` calls `get_timestamp()` separately from line 43 here; if the minute changes between the two calls, the
  report/summary names carry a stamp one minute different from the folder names (LOW impact, HIGH that two calls exist).

**`_validate_inputs(self)` (56-65)** - raises `ValueError("Invalid processing mode: ...")` unless the mode is a key of PROCESSING_MODES
('fresh', 'rerun', 'resume') (58-59; RAN). Creates the target folder with all parents (63); failure -> `ValueError("Cannot create target
folder: ...")` (64-65). RAN: a non-existent nested target is created. Caller: 51.

**`create_client_structure(self, client_info) -> Dict[str, Path]` (67-144)** - READ + RAN.
1. `state_code = get_state_code(client_info['state'])` (78).
2. `client_state_key = create_client_state_key(client, state, max_length=self.client_name_max_length)` (79-83) - the level-2 folder name
   (Private->Pvt, Limited->Ltd whole-word any case, state code, length limit; see helpers).
3. `client = sanitize_filename(client_info['client'])` (86) - used only inside category folder names and logs.
4. Level 1 from `_get_level1_folder()` (90).
5. Level 2 = level1 / client_state_key, created (93-96); version = level2 / `Version-{timestamp}`, created (99-104).
6. For each of `gstr3b`, `itc`, `sales` (113-135): name from the template ("GSTR-3B Exports ({client})", "Other ITC related files ({client})",
   "Sales related files ({client})"); include-name decision: if the global flag is on -> include (121-123), else look up
   `client_folder_settings.get(f"{client}-{state_code}", False)` (118, 126 - always False in B); if not included, remove the text
   `" ({client})"` (129-130); create folder (132-135).
7. Returns dict with keys `level1, level2, version, gstr3b, itc, sales` (107-111, 134).
- Exceptions: any error is logged, appended to `self.errors`, and **re-raised** (140-144) - the caller marks the client failed
  (processing_controller.py:173-181). Folders created before the error remain (no clean-up; `cleanup_on_error` is never called).
- Side effects: creates up to 6 folders; appends each to `created_folders` (96, 104, 135, plus 155/178).
- RAN folder tree (fresh mode, 3 clients):
  `Annual Statement-211026 0334/ClientA-DL/Version-211026 0334/{GSTR-3B Exports, Other ITC related files, Sales related files}`;
  for "ClientB Private Limited"/"Tamil Nadu" the level-2 folder is `ClientB Pvt Ltd-TN`.
- Inconsistency (HIGH, RAN): level-2 folder uses `create_client_state_key` (spaces kept, whole-word any-case abbreviations), category
  folder suffix and report names use `sanitize_filename` (spaces -> underscores, case-sensitive substring abbreviations), and the
  validation tab's "folder" column shows the parser key `"{client}-{state_code}"` (gui/views/validation_view.py:379 shows `client_key`)
  - three different spellings of the same client (e.g. parser key `ClientA Private Limited-DL`, folder `ClientA Pvt Ltd-DL`, report part `ClientA_Pvt_Ltd`).
- Collision (HIGH, RAN): two different clients whose keys become equal after abbreviation or truncation share one level-2 folder and, in
  one run, one version folder: e.g. "ClientA Private Limited" and "ClientA Pvt Ltd" (same state) -> both `ClientA Pvt Ltd-DL`, and their
  report names are also identical (`ITC_Report_ClientA_Pvt_Ltd_DL_...`), so the second client's report overwrites the first's; two long
  names that differ only after character 32 (with a 2-letter code and limit 35) -> same folder.
- Caller: processing_controller.py:99.

**`_get_level1_folder(self) -> Path` (146-179)** - READ + RAN.
- 'fresh': `target / "Annual Statement-{timestamp}"`, created with `exist_ok` (148-156). A second fresh run in the **same minute reuses
  the same folder** (RAN).
- 'rerun' / 'resume': lists sub-folders of the target whose name starts with "Annual Statement-" (160-163) and picks the one with the
  newest **folder modification time** (167); if none, creates a new one like fresh (171-179, logs warning). Folder mtime changes when
  something is added directly inside that folder (a new client folder, the summary workbook), or by any manual change, so "latest" means
  "most recently touched", not "newest date in the name" (HIGH, RAN with `os.utime`).
- The Extract tab instead picks the level-1 folder by **reverse alphabetical name** (gui/pyqt_main_window.py:432-436). With `DDMMYY`
  names that is not date order: RAN - names "Annual Statement-311225 1000" (31-12-2025), "Annual Statement-211026 0335" (21-10-2026),
  "Annual Statement-010126 0900": reverse name order puts the 31-12-2025 folder first, although 21-10-2026 is newer. So organiser and
  Extract tab can disagree on which run is "latest" (HIGH).
- Large target folders: one `stat()` per matching sub-folder (167); fine.
- If the target holds a file (not folder) starting with the prefix, `is_dir()` filters it (162).
- Caller: 90.

**`organize_files(self, client_info, folders, progress_callback=None) -> List[dict]` (181-231)** - READ + RAN.
- Flattens `client_info['files']` (type -> list) into `(type, file_data)` pairs (199-202), counts them (204).
- For each (207-229): optional progress callback "Copying <name>" (208-210) - **the only caller never passes one**
  (processing_controller.py:106), so the parameter is unused in practice; calls `_organize_single_file` (212); appends result to the local
  list and to `self.copied_files` (213-214).
- On 'Success' (217-229): takes the destination file name and **writes it back into the caller's `client_info`** (`name` and new key
  `sanitized_name`) for the matching `full_path` (224-229). This mutates the scan result held by the window (`ScanResult.client_data`,
  passed by reference via pyqt_main_window.py:368 and processing_controller.py:30, 85) from the worker thread; the window may read the
  same dicts (e.g. Dry Run at pyqt_main_window.py:335-339). No lock. A second processing run without re-scanning starts from the already
  sanitised names (MEDIUM - the race window is small; consequences cosmetic).
- Returns the list of result dicts (231). Docstring line 186 "FIXED: Updates client_info with sanitized filenames after copying." agrees.
- Caller: processing_controller.py:106.

**`_organize_single_file(self, file_type, file_data, folders) -> dict` (233-283)** - READ + RAN.
- Result skeleton (237-244): `filename` (original name), `file_type`, `source` (`full_path`), `destination` '', `status` 'Failed', `error` None.
- Destination folder from `_get_destination_folder` (248); none -> error "No destination folder for type: ..." and return (249-251).
- Destination path = folder / `sanitize_filename(name)` (254-257); also stores `sanitized_filename`.
- If the destination already exists (260-270):
  - 'fresh': makes a backup copy next to it, `"{stem}_backup_{DDMMYY_HHMM}{suffix}"` (263-265, via `create_backup`), then continues to
    overwrite. If the backup fails, `create_backup` returns None and **the overwrite still happens** (HIGH, READ).
  - 'resume': returns status 'Skipped', error "File already exists" (266-270).
  - 'rerun': **no branch** - the existing file is overwritten with no backup (HIGH, RAN).
- Copies with `safe_copy_file(source, dest, verify=True)` (273): 'Success' (274-275) or error "Copy operation failed" (277).
- Any exception -> `error = str(e)`, logged (279-281). Returns result (283).
- **Originals are never moved, renamed or deleted** - only read by `shutil.copy2` (helpers.py:169). The only deletions in my files are of
  the *destination* copy when sizes differ (helpers.py:177) and `shutil.rmtree` of an empty version folder in the never-called
  `cleanup_on_error` (375) (HIGH, READ; grep for remove/unlink/rmtree/move/rename).
- RAN results:
  - fresh, then fresh again in the same minute: second run found existing files -> backups `GSTR3B-ClientA-Delhi-April_backup_211026_0335.xlsx`
    etc. created in the category folders, then overwritten with the new content.
  - rerun in the same minute: overwritten, no backup.
  - resume in the same minute: both files 'Skipped'.
  - resume one minute later (timestamp forced): new `Version-...` folder, nothing skipped - **'resume' only skips files that are already in
    the *current* version folder, i.e. only within the same clock minute; across minutes it behaves exactly like 'rerun'** (HIGH, RAN).
    This does not match the mode's description "Continue previously interrupted processing" (constants.py:170, gui/views/setup_view.py:346).
  - resume skip + template data: the skipped file's `name` is not updated, so the template is given the original stem
    `GSTR-2B-Reco-ClientB Private Limited-Tamil Nadu-Apr` while the folder holds `GSTR-2B-Reco-ClientB_Pvt_Ltd-Tamil_Nadu-Apr.xlsx`
    (RAN, t_resume_name.py) - the report would point at a file name that does not exist (HIGH). Same for any file whose copy failed.
- Backup files stay inside the category/version folders forever and are not recognised by anything (they would be picked up by a
  Power Query that reads a whole folder, e.g. the "GSTR-3B Exports" folder whose path is written to the ITC template (excel_handler.py:555)
  - whether the template's query reads every file in that folder is not visible in the repository; LOW-MEDIUM, check the template's query).
- Two different source names that sanitise to the same destination name inside one client and one folder: the second copy overwrites
  the first (fresh: after backing it up) (MEDIUM, READ; not RAN).
- Caller: 212.

**`_get_destination_folder(self, file_type, folders) -> Optional[Path]` (285-298)** - fixed table (288-295): 'GSTR-3B Export' -> gstr3b;
'GSTR-2B Reco' -> itc; 'IMS Reco' -> itc; 'Sales' -> sales; 'Sales Reco' -> sales; 'Annual Report' -> version (comment "Goes in version
folder"). Unknown types -> None. Duplicates (and disagrees in wording with) the `'folder'` values in FILE_PATTERNS (constants.py:18, 26, 34, 42, 50, 58),
where Annual Report says `'root'`; those `'folder'` values are copied into each file record by the parser (file_parser.py:90) but never read
(HIGH, grep). Caller: 248.

**`create_organization_report(self, folders, client_info) -> Optional[Path]` (300-362)** - READ + RAN.
- File `"_Organization_Report_{client}_{state}.txt"` in the version folder (313-314) using the **raw** client and full state name (not
  sanitised, not abbreviated, with spaces; RAN: `_Organization_Report_ClientB Private Limited_Tamil Nadu.txt`). Written in UTF-8 with mode
  'w' (354-355) - overwrites one from the same minute.
- Content (317-351): title, "=" line, "Generated: YYYY-MM-DD HH:MM:SS", "Client:", "State:", "Processing Mode:"; "FOLDER STRUCTURE:" with each
  folder except level1/level2 as a path relative to level1 (329-333); "FILES ORGANIZED:" listing every file of the client by type using
  `sanitized_name` if present else `name` (338-343) - it lists all files, including ones that failed or were skipped; "STATISTICS:" with
  "Total Files: {file_count}", "Successfully Copied: {count}", "Folders Created: {count}".
- **Bug (HIGH, RAN)**: "Successfully Copied" counts `self.copied_files`, which accumulates over **all clients processed so far in the run**
  (346-347), and "Folders Created" is `len(self.created_folders)`, also cumulative and with the level-1 folder counted once per client
  (351). RAN: second client with 2 files showed "Successfully Copied: 10" and "Folders Created: 12"; third client with 1 file showed 11 and 18.
- Exceptions: logged, returns None (360-362); the caller ignores the return value (processing_controller.py:133).
- The leading "_" makes `_is_folder_empty` ignore this file (386).
- Caller: processing_controller.py:133.

**`cleanup_on_error(self, folders)` (364-379)** - only for 'fresh' (366-367): if the version folder has no files other than names starting
with "_" (373, `_is_folder_empty`), deletes it with `shutil.rmtree(..., ignore_errors=True)` (374-376). Errors logged (378-379).
**NO CALLERS FOUND (dead)** in B (grep; it had none in c258047 either). RAN directly: version folder with only the organisation report -> deleted.

**`_is_folder_empty(self, folder) -> bool` (381-390)** - walks the folder recursively (`rglob('*')`, 385); returns False at the first file
whose name does not start with "_" (386-387); True otherwise (388); any exception -> False (389-390). Caller: 373 only (so dead too).

**`get_summary(self) -> dict` (392-409)** - counts Success/Failed in `copied_files`, `folders_created` = number of distinct created folders
(403), `files_processed`, `skipped_files` = rest (407), `errors` = len(self.errors) (408). **NO CALLERS FOUND (dead)** (grep). RAN: works.

#### Business rules in this file
- Folder hierarchy: target / "Annual Statement-DDMMYY HHMM" / "<client key>" / "Version-DDMMYY HHMM" / three category folders; Annual Report
  sits directly in the version folder (99-135, 288-295; templates in constants.py:81-90). Not a GST rule. Timestamp is the run start, to the minute.
- [GST RULE] (borderline, which return file goes into which group) GSTR-3B exports -> "GSTR-3B Exports"; GSTR-2B Reco and IMS Reco ->
  "Other ITC related files"; Sales and Sales Reco -> "Sales related files"; Annual Report -> version folder (288-295).
- Client name in category folder names only when the global option is on (121-130) (per-client option dead in B).
- Existing destination handling per mode: fresh = backup then overwrite; rerun = overwrite; resume = skip (260-270).
- Rerun/resume reuse the most-recently-modified "Annual Statement-" folder (160-169).
- Client name max length default 35 and "0 or negative means 35" (30, 42).

#### Edge cases explicitly handled
- Invalid mode (58-59); target folder missing -> created (63); no existing level-1 folder in rerun/resume -> new one (171-179); unknown
  file type -> per-file error (249-251); existing destination per mode (260-270); copy failure -> per-file error (277, 279-281);
  report-writing failure -> None (360-362); `client_name_max_length <= 0` -> 35 (42).

#### Gaps, likely failures and silent bugs
1. 'resume' is effectively 'rerun' unless re-run within the same minute (RAN) (HIGH).
2. 'rerun' overwrites without backup when the destination exists (RAN) (HIGH).
3. Backup failure does not stop the overwrite (263-265 then 273) (HIGH).
4. Cumulative statistics in each client's organisation report (RAN) (HIGH).
5. Two clients can share one level-2 folder and one report name after abbreviation/truncation (RAN) (HIGH).
6. Skipped/failed copies leave the template pointing at the unsanitised name (RAN) (HIGH).
7. Per-client include-name setting never populated in B (READ) (HIGH); `progress_callback` never passed (READ) (HIGH);
   `cleanup_on_error`, `_is_folder_empty`, `get_summary` dead (HIGH); `Tuple` import unused.
8. Level-1 "latest" by mtime here vs by name in the Extract tab (RAN) (HIGH).
9. Partial folders stay behind when a client fails midway (140-144; nothing calls `cleanup_on_error`) (HIGH).
10. Network drive drop: each copy fails individually ("Copy operation failed"), a folder-creation failure fails the whole client; a copy
    interrupted midway may leave a partial destination file because `shutil.copy2` writes directly to the final name and the verify step
    (helpers.py:172-178) itself raises when `stat()` fails, so the clean-up `unlink` is not reached (MEDIUM, READ).
11. Two app copies or two runs in the same minute against the same target share folders and file names; no locking (MEDIUM, READ).
12. A relative target path (if typed) would be resolved against the program's current folder (38, 63) (LOW; depends on the GUI field).
13. Windows long paths: target + "Annual Statement-DDMMYY HHMM" + up-to-35-char key + "Version-DDMMYY HHMM" + category (+ client name) +
    file name (up to 190 chars after sanitising) can pass 260 characters; nothing here shortens or checks it (MEDIUM, READ).
- Closures/lambda: `key=lambda f: f.stat().st_mtime` (167) used immediately - fine. No mutable defaults, no `is`/`==` issues.

#### Excel workbook touchpoints
- None directly; Excel files are copied as plain files (`safe_copy_file`, `create_backup`).

#### Files, folders, settings, caches and logs
- Creates folders: `target/Annual Statement-{DDMMYY HHMM}` (150-155, 173-178), `.../{client key}` (94-95), `.../Version-{DDMMYY HHMM}` (99-103),
  `.../GSTR-3B Exports[ (<sanitised client>)]`, `.../Other ITC related files[ (...)]`, `.../Sales related files[ (...)]` (113-135).
- Copies each source file to `<destination folder>/<sanitize_filename(original name)>` (254-273); backups `"{stem}_backup_{DDMMYY_HHMM}{suffix}"`
  beside the existing destination (263; helpers.py:194-198).
- Writes `<version folder>/_Organization_Report_{client}_{state}.txt` (313-355), UTF-8 text.
- Reads: target folder listing (160-163) and folder mtimes (167).
- Logs: module logger `core.file_organizer`; INFO lines include target path, client names and file names (53-54, 137, 168, 275).
- Readers elsewhere: Extract tab scans `Annual Statement-*` / any sub-folder of level 2 / `ITC_Report_*.xlsx`, `Sales_Report_*.xlsx`,
  `*_Refreshed*` (pyqt_main_window.py:432-470); report_processor globs `{ITC|Sales}_Report_*.xlsx` in the version folder (report_processor.py:76-77).

#### Comments, docstrings, version strings
- Module docstring "File Organizer for GST File Organizer v3.0" (2) - stale version.
- `__init__` docstring lists 2 of 5 parameters (31-37). `create_organization_report` docstring "Create a text report of the organization process" - agrees.
- Comments 120 "Global setting overrides individual settings" and 122/125 describe the logic correctly, but the individual settings are never set in B.
- Comment 216 "IMPORTANT: Update the file data with the sanitized filename" and 223 "This ensures Excel handler gets the correct filename" -
  true only for successful copies.
- No TODO/FIXME or commented-out code.

#### Pattern notes
- Error handling: per-file errors are captured in result dicts (never raised); per-client folder errors are raised to the caller.
- Uses `safe_path_join`/`ensure_path_exists` helpers consistently; the level-1 prefix literal (162) duplicates the constant (82).
- Threading: runs in the processing QThread; mutates shared `client_info` dicts without locks (224-229).

---

### core/file_parser.py (399 lines) - [CORE]

- **Purpose**: scans the top level of the source folder for Excel files, checks each file's size and first bytes, recognises its type,
  client and state from the file name with fixed regular expressions, groups files per client, and works out which expected file types
  are missing or duplicated.

#### Imports
- Standard library: `logging` (7), `pathlib.Path` (8), `datetime` (9), `typing Dict, List, Tuple, Optional, Any, Union` (10; `Optional` unused),
  `collections.defaultdict` (11).
- Project: `utils.constants FILE_PATTERNS, EXPECTED_FILE_TYPES` (13); `utils.helpers find_excel_files, validate_client_name, format_size,
  get_state_code` (14-16). Importing this module also runs core/__init__.py (see that section).
- No OS-specific modules.

#### Module-level constants and globals
- `logger` (18). Inline literals: minimum size `1024` bytes (184); signatures `b'PK\x03\x04'` (xlsx/xlsm, a ZIP file) and
  `b'\xd0\xcf\x11\xe0'` (old .xls) (197-198) - duplicates of constants.py:182-185 (`EXCEL_SIGNATURES`, unused) and helpers.py:240-241;
  header read length 8 bytes (196); critical types list (387); suggestion keywords (321-332).

#### File-name recognition rules (exact) - [GST RULE] (decides which return/report a file is treated as)
Tried **in this order** (dict order of FILE_PATTERNS, constants.py:13-62); the **first** regex that matches wins (file_parser.py:78-118, `break` at 118).
All use `re.IGNORECASE` and `.match` (anchored at the start; each pattern also ends with `$`). `[^-]+` means "one or more characters that are not a hyphen".

| # | Key | Regex (exact) | type | category | 'folder' value | groups |
|---|---|---|---|---|---|---|
| 1 | GSTR-2B-Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | GSTR-2B Reco | ITC | itc | client, state, period |
| 2 | ImsReco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | IMS Reco | ITC | itc | client, state, date |
| 3 | GSTR3B | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | GSTR-3B Export | GSTR3B | gstr3b | client, state, month |
| 4 | Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | Sales | Sales | sales | client, state, start_month, end_month |
| 5 | SalesReco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Sales Reco | Sales | sales | client, state, period |
| 6 | AnnualReport | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Annual Report | Annual | root | client, state, year |

- Extensions: `\.xlsx?$` = ".xlsx" or ".xls" only (any case). **".xlsm" files never match any pattern** although `find_excel_files`
  collects them (helpers.py:255) - they always end up as variations "No matching pattern found" (HIGH, RAN).
- Pattern 3 accepts an optional leading browser-duplicate counter like "(1) " or "(12)" (RAN: both parsed). A trailing " (1)" before
  the extension is accepted as part of the month (RAN: month = "April (1)").
- Patterns 1, 5, 6 end with `(.+)` (anything, hyphens allowed); patterns 2, 3, 4 need an exact number of hyphen-separated parts.
- **Hyphen in a client name** (HIGH, RAN): pattern 1 with "GSTR-2B-Reco-Client-A-Delhi-Apr.xlsx" parses as client "Client", state "A",
  period "Delhi-Apr" - silently wrong grouping (only an unseen warning, see below); patterns 3/4/2 simply fail to match ("GSTR3B-ClientA-Delhi-Apr-2024.xlsx",
  "Sales-ClientA-Delhi-Apr-Mar-2025.xlsx" -> variations).
- Exact spellings required (RAN): "GSTR-3B-..." (with hyphen) fails; "Sales Reco-..." and "Annual Report-..." (with space) fail; "IMSReco"
  works (case-insensitive); IMS date must be exactly 8 digits ("310325" and "31-03-2025" fail).
- An all-space group: "Sales- -Delhi-Apr-Mar.xlsx" matches but the client strips to "" -> not grouped, becomes a variation with issue
  "Could not parse filename" (RAN).
- "GSTR3B-ClientA-Delhi-April.xlsx.xlsx" parses with month "April.xlsx" (RAN).
- Letter case of client/state is kept as typed, so "ClientA" and "clienta" become two different clients (RAN shows lower-case kept) (HIGH).
- Metadata groups (`period`, `date`, `month`, `start_month`, `end_month`, `year`) are stored in `result['metadata']` (102) but **never read
  anywhere** in the project (grep for `metadata`: only file_parser.py:74, 102) (HIGH). There is **no financial-year or tax-period logic**
  in B's core/utils: periods/months/years in file names are neither checked nor used (HIGH, READ). [GST RULE] (TIME-VARYING: the expected
  naming of periods/years is whatever the author's download/rename practice is).
- The `category` and `folder` values are copied into each file record (89-90) and never read (grep) (HIGH).

#### class FileParser (21-399)

**`__init__(self)` (26-30)** - `self.patterns = FILE_PATTERNS`, `self.expected_types = EXPECTED_FILE_TYPES`, `progress_callback = None`, then `reset()`.
Caller: gui/controllers/scan_controller.py:26 (new parser per scan, inside the scan worker thread).

**`reset(self)` (32-46)** - `scanned_files = {}`; `client_data = defaultdict(lambda: {'client': '', 'state': '', 'files': defaultdict(list),
'missing_files': [], 'extra_files': [], 'status': 'Unknown', 'total_size': 0, 'file_count': 0})` (the lambda builds a fresh dict for each
new client - no shared-mutable bug, READ); `variations = []`; `errors = []`. Callers: 30, 138.

**`_update_progress(self, current, total, message)` (48-54)** - calls the callback if set; any exception from it is logged as a warning and
ignored. Callers: 147, 151, 155, 158. The scan worker's callback emits a Qt signal (scan_controller.py:29), which is safe from a worker thread.

**`parse_filename(self, filename) -> dict` (56-124)** - READ + RAN.
- Default result (66-75): `parsed False`, `type/category/folder 'Unknown'`, `client ''`, `state ''`, `metadata {}`.
- Loops patterns in order (78); on the first match (81): sets `parsed True`, `pattern` (key), `type`, `category`, `folder` (85-91); maps
  groups by name - 'client' and 'state' stripped into their fields, others into `metadata` (94-102); skips empty groups (96).
- If both client and state are non-empty, calls `validate_client_name` on **both** (105-115) and adds `warnings` texts "Client name issue: ..." /
  "State name issue: ..." - the state check reuses the client function, so its messages say "Client name ..." (RAN: "State name issue:
  Client name too short"). **`warnings` is never read anywhere** (grep), so these warnings are invisible (HIGH).
- No match: logs WARNING "Could not parse filename: <real file name>" (121) and sets `error 'No matching pattern found'` (122).
- Returns the dict (124). Caller: 229 only.

**`scan_folder(self, folder_path, progress_callback=None) -> (scanned_files, client_data, variations)` (126-169)** - READ + RAN.
- `reset()`, store callback (138-139). Missing folder -> `FileNotFoundError` (142-143); not a folder -> `ValueError` (144-145).
- Progress "Finding Excel files..." (147); `find_excel_files(folder)` (148) = top level only, `*.xlsx`, `*.xls`, `*.xlsm`, minus `~$*`, sorted.
- None found: warning + progress "No Excel files found" + returns empty results (149-152).
- Progress "Found N files, processing..." (155); per file progress "[i/N] <name>" (158) and `_process_file` (159).
- `_analyze_client_completeness()` (161); logs summary (163-164); returns `(scanned_files, dict(client_data), variations)` (165) - only the
  outer dict is converted; each client's `files` stays a defaultdict.
- Any exception: logged, added to `self.errors`, **re-raised** (166-169) -> scan worker emits the error text to the window (scan_controller.py:39-41).
- RAN: 14 files found -> progress messages as listed (16 callback calls).
- Caller: scan_controller.py:27-30.

**`_process_file(self, file_path, base_folder)` (171-256)** - READ + RAN.
1. `stat()`; on `OSError`, log warning and **return without recording anything** (175-179) - the file vanishes from all results.
2. Size `< 1024` bytes -> variation "File too small to be valid Excel" (184-191).
3. Reads first 8 bytes (195-196); not ZIP/OLE signature -> variation "Invalid or corrupted Excel file" (199-206); read error -> variation
   "Cannot read file: <error>" (207-215).
4. Builds `file_info` from the same stat: name, path, size, size_formatted, modified, created (`st_ctime` - on Windows the creation time,
   on Linux the last metadata change), extension (218-226).
5. `parse_filename(name)` (229); merges into `file_data` with `full_path` and `relative_path` (232-237).
6. Stores `scanned_files[name] = file_data` (240) - keyed by file name (unique because the scan is not recursive).
7. If parsed and client and state non-empty -> `_add_to_client_data` (243-244); else variation with issue = `parsed['error']` or
   "Could not parse filename", size, and `pattern_expected = _suggest_pattern(name)` (245-252).
- Outer `except Exception` (254-256): logs and appends to `self.errors`; **`FileParser.errors` is never read** outside this class (grep) - those
  errors are only in the log (HIGH).
- Consequence (HIGH, RAN): files rejected in steps 1-3 are not in `scanned_files`, so the window's "Files found" (= `total_files`) is lower
  than the number of Excel files actually present (RAN: 14 found, `total_files` 12, `variations` 3, `unparsed_files` 1).
- Caller: 159.

**`_add_to_client_data(self, file_data)` (258-278)** - READ + RAN.
- `state_code = get_state_code(state)` (261); client key `f"{client}-{state_code}"` (262).
- Sets `client`, `state` (**last file wins**, comment "Keep original state name") and `state_code` (266-268); appends the file under its
  type (271-272); adds size and count (275-276).
- RAN: files written with "Delhi" and with "DL" for the same client merged into one client `ClientA-DL` (both map to "DL"), and the
  stored state became whichever was processed last ("Delhi"). Files with "New Delhi" or "NCT of Delhi" would instead get codes "ND"/"NO"
  and form separate clients (RAN via `get_state_code`) (HIGH). [GST RULE] (state-code logic decides grouping).
- Caller: 244.

**`_analyze_client_completeness(self)` (280-314)** - READ + RAN. [GST RULE] (which returns/reports a client must have)
- For each client: missing = expected types (6, constants.py:68-75) minus types found; stored sorted (283-288).
- Extras: any type with more than one file, **except 'GSTR-3B Export'** (comment "GSTR-3B can have multiple files (monthly filings - unlimited)")
  -> text "Multiple <type> (<n> files)" (291-299). [GST RULE] TIME-VARYING (filing frequency assumptions).
- `completeness` = found types / 6 x 100, rounded to 1 decimal (302-303) (types, not files).
- Status (306-314): no missing and no extras -> 'Complete' (icon ✅); any missing -> f'Missing {n} files' where n is the number of missing
  **types** (icon ⚠️) even if there are also duplicates; else 'Has duplicates' (icon ⚠️). `status_icon` is never read by B's GUI (grep: only
  this file) (HIGH).
- RAN: client with all 6 types and 3 GSTR-3B files -> 'Complete', 100.0; client with two GSTR-2B files only -> 'Missing 5 files', 16.7,
  extras "Multiple GSTR-2B Reco (2 files)".
- No check on how many GSTR-3B months exist (READ). Caller: 161.

**`_suggest_pattern(self, filename) -> str` (316-334)** - lower-cases the name and checks, **in this order, stopping at the first hit** (if/elif):
1. contains "gstr" and "2b" -> "GSTR-2B-Reco-ClientName-State-Period.xlsx"
2. contains "ims" -> "ImsReco-ClientName-State-DDMMYYYY.xlsx"
3. contains "gstr" and "3b" -> "GSTR3B-ClientName-State-Month.xlsx"
4. contains "sales" and "reco" -> "SalesReco-ClientName-State-Period.xlsx"
5. contains "sales" -> "Sales-ClientName-State-StartMonth-EndMonth.xlsx"
6. contains "annual" -> "AnnualReport-ClientName-State-Year.xlsx"
otherwise "Check file naming convention" (334). Because of `elif`, at most one suggestion exists, so `" or ".join` (334) never joins two (HIGH).
Plain substring tests give odd hints (RAN): "claims summary.xlsx" -> the IMS hint ("claims" contains "ims"); a correctly named `.xlsm` file gets
the `.xlsx` pattern as its hint, which does not tell the user the real reason (the extension). Caller: 251 only (shown in the summary's
Variations sheet only if the caller passed variations, which it never does - see excel_handler).

**`get_statistics(self) -> dict` (336-364)** - total_files, parsed_files (files with `parsed` True - includes a parsed file whose client was
empty), unparsed_files, total_clients, complete_clients (status 'Complete'), incomplete_clients, variations (count), total_size(+formatted),
file_type_distribution, parsing_rate %, completion_rate % (zero-division guarded at 361-363). Caller: scan_controller.py:31; shown by
pyqt_main_window.py:250-311.

**`validate_results(self) -> List[dict]` (366-399)** - per client: empty client name -> 'Empty Client Name' (Error) (372-377); empty state ->
'Empty State Name' (Error) (379-384); missing any of `critical_types = ['Annual Report', 'GSTR-3B Export']` -> 'Missing Critical Files'
(Warning) with details (387-397). **NO CALLERS FOUND (dead)** (grep; also none in c258047). RAN directly: works. [GST RULE] (which files are "critical").

#### Business rules in this file
- [GST RULE] File-type recognition table above (constants.py:13-62 used at 78-118). TIME-VARYING (return/report names and portal download names can change).
- [GST RULE] Six expected file types per client (EXPECTED_FILE_TYPES; 283-288).
- [GST RULE] GSTR-3B may have any number of files; every other type should have exactly one (291-297). TIME-VARYING.
- [GST RULE] Client grouping key = client text + state code (261-262); state code from STATE_CODE_MAPPING or fallback (helpers.py:485-511).
- [GST RULE] (dead) "Critical" types: Annual Report and GSTR-3B Export (387).
- Minimum Excel size 1 KB (184); ZIP/OLE signature check (197-199).
- Client/state names 2-100 characters, no `<>:"/\|?*` (only as unseen warnings; helpers.py:321-342).
- Completeness = share of the 6 types present, 1 decimal (302-303) - not a tax figure.

#### Edge cases explicitly handled
- Folder missing / not a folder (142-145); no Excel files (149-152); file cannot be stat-ed (175-179, silently skipped); tiny files (184-191);
  wrong signature (199-206); unreadable file (207-215); unparseable name (245-252); empty groups (96); progress callback failing (51-54);
  Excel lock files `~$...` (helpers.py:260); duplicate-download prefix "(n)" for GSTR3B (constants.py:31).

#### Gaps, likely failures and silent bugs
1. Not recursive: files in sub-folders of the source are ignored with no message (helpers.py:257 `glob`, not `rglob`) (HIGH, RAN).
2. ".xlsm" collected but never recognised (HIGH, RAN).
3. Extension case: on Linux `*.xlsx` does not match ".XLSX" (RAN: such a file was not even found); on Windows pathlib's glob is
   case-insensitive (MEDIUM - standard Python behaviour, not run on Windows).
4. Hyphenated client names mis-split silently for GSTR-2B/SalesReco/AnnualReport (RAN) (HIGH).
5. Different spellings of the same state/client create separate clients; the last file's state text is kept (RAN) (HIGH).
6. Warnings and per-file errors are collected but never shown (HIGH).
7. A file that cannot be stat-ed disappears without trace (175-179) (HIGH).
8. "Files found" in the window counts only files that passed the size/signature checks (RAN) (HIGH).
9. `status_icon`, `category`, `folder`, `metadata`, `pattern`, `warnings`, `modified`, `created`, `relative_path` are produced but unused by B (grep) (HIGH).
10. Very large folders: one stat + one 8-byte read + one progress signal (one log line in the window) per file (MEDIUM - the per-file log line could slow the window; B-R3/B-R4 to confirm).
11. Network drive drop during scan: per-file OSError -> warning/variation; folder vanishes -> glob error caught in `find_excel_files`
    (returns `[]`, logged) -> shown as "No Excel files found" rather than as a network error (HIGH, READ helpers.py:264-266).
12. Running twice: each `scan_folder` resets state (138) - safe. Two app copies: read-only, safe.
- No closure-in-loop, mutable-default, `is`/`==` problems (READ).

#### Excel workbook touchpoints
- None via openpyxl/COM; only raw `open(file, 'rb').read(8)` per file (195-196).

#### Files, folders, settings, caches and logs
- Reads the source folder listing (via helpers.find_excel_files), each file's metadata and first 8 bytes. Writes nothing to disk.
- Logs: WARNING "Could not parse filename: <name>" (121), DEBUG per parse/add (117, 278), INFO counts (154, 163-164); real client file names go to the log.

#### Comments, docstrings, version strings
- Module docstring "File Parser for GST File Organizer v3.0 / Robust pattern matching and file categorization. / Fixed with proper Union
  import." (1-5) - version stale; "Fixed with proper Union import" is a leftover change note.
- `_process_file` docstring "Process a single file with minimal I/O (one stat + one open)" (172) - agrees (7e09e5c).
- Comment 146 "Find all Excel files (fast glob, no I/O per file)"; comment 183 "Skip files too small to be valid Excel (<1KB)" - they are
  not skipped silently, they are listed as variations.
- Comment 294 "GSTR-3B can have multiple files (monthly filings - unlimited)" - agrees with 295-296.
- No TODO/FIXME, no commented-out code.

#### Pattern notes
- Errors: per-file problems become "variations" or `errors` entries; folder-level problems are raised.
- Duplicated logic: the size/signature check exists here (184-206) and in helpers.validate_excel_file (232-243, now unused) and as unused constants (constants.py:181-185).
- Runs in the scan QThread; only signals cross threads (scan_controller.py:29).

---

### utils/__init__.py (38 lines) - [UTIL]

- **Purpose**: package marker for `utils` that re-exports most constants and helper functions.
- **Imports**: project only - from `.constants`: FILE_PATTERNS, EXPECTED_FILE_TYPES, FOLDER_STRUCTURE, EXCEL_TEMPLATE_MAPPING, GUI_CONFIG,
  PROCESSING_MODES, MESSAGES, EXCEL_EXTENSIONS, EXCEL_SAFETY (6-10); from `.helpers`: get_timestamp, get_date_only, get_safe_timestamp,
  ensure_path_exists, safe_path_join, sanitize_filename, get_file_info, safe_copy_file, create_backup, validate_excel_file, find_excel_files,
  format_size, truncate_path, clean_windows_path, extract_filename_without_extension, validate_client_name, get_unique_path, ProgressTracker,
  format_duration, calculate_file_hash, get_relative_excel_path (12-20).
- **Module-level**: `__all__` lists the same 30 names (22-36); `__version__ = '3.0.0'` (38).
- **Not re-exported** (HIGH, RAN): helpers `get_short_path`, `get_state_code`, `create_client_state_key` and constants `STATE_CODE_MAPPING`,
  `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES` - the three helpers that are actually used are imported from
  `utils.helpers` directly by callers. Nobody imports from the `utils` package itself (all callers use `utils.helpers` / `utils.constants`), so the
  re-export list is unused, but the file still runs (and imports both modules) whenever `utils.helpers` or `utils.constants` is imported (RAN).
- Edge: `helpers.get_state_code` imports `.constants` lazily inside the function (helpers.py:491) "to avoid circular import"; with this
  `__init__` importing constants first, no circular import exists at present (RAN: imports fine) (MEDIUM that the lazy import is still needed).
- Comments: docstring "Utils package for GST File Organizer v3.0" - stale version.
- Pattern: same facade pattern as core/__init__.py; must be maintained by hand (already out of date: 3 used helpers missing).

---

### utils/constants.py (275 lines) - [CONFIG]

- **Purpose**: central table of file-name patterns, expected file types, folder-name templates, template cell mappings, GUI/message texts,
  safety settings and a state-name-to-code table.
- **Imports**: `re` (6); `typing Dict, List, Tuple` (7; `Tuple` unused). No OS-specific modules.
- Type hint at line 13 uses lowercase `any` (Python's built-in function) instead of `typing.Any` - harmless typo (HIGH).
- Unchanged since the initial commit edac6a3 (HIGH, git log).

#### Every constant (name = value, line, what it controls, who uses it)

**FILE_PATTERNS (13-62)** - [GST RULE] TIME-VARYING. Six entries, each with `pattern`, `type`, `category`, `folder`, `description`, `groups`
(exact regexes in the file_parser table above). Descriptions (user-facing examples, never displayed in B): 'GSTR-2B-Reco-ClientName-State-Period.xlsx' (19),
'ImsReco-ClientName-State-DDMMYYYY.xlsx' (27), 'GSTR3B-ClientName-State-Month.xlsx' (35), 'Sales-ClientName-State-StartMonth-EndMonth.xlsx' (43),
'SalesReco-ClientName-State-Period.xlsx' (51), 'AnnualReport-ClientName-State-Year.xlsx' (59). Used by file_parser.py:27/78 (patterns, type, groups;
category/folder/description copied or ignored). `description` unused (the same texts are hard-coded again in file_parser.py:322-332).

**EXPECTED_FILE_TYPES (68-75)** = ['GSTR-2B Reco', 'IMS Reco', 'GSTR-3B Export', 'Sales', 'Sales Reco', 'Annual Report'] - [GST RULE] TIME-VARYING
(the set of files a client needs). Used by file_parser.py:28, 284.

**FOLDER_STRUCTURE (81-90)** = `level1: 'Annual Statement-{timestamp}'` (82, used file_organizer.py:150, 173), `level2: '{client}-{state}'`
(83, **unused** - level 2 is built by `create_client_state_key`), `version: 'Version-{timestamp}'` (84, used 99), `categories`: `gstr3b:
'GSTR-3B Exports ({client})'`, `itc: 'Other ITC related files ({client})'`, `sales: 'Sales related files ({client})'` (85-89, used 113).
Not a GST rule (folder naming). The prefix "Annual Statement-" is also hard-coded at file_organizer.py:162 and gui/pyqt_main_window.py:434.

**EXCEL_TEMPLATE_MAPPING (96-120)** - [GST RULE] (borderline) TIME-VARYING (depends on template version). Which value is written into which cell
of sheet 'Links':
- ITC (97-108): B2 = gstr3b_folder; B4 = annual_folder; B5 = annual_filename; B7 = gstr2b_folder; B8 = gstr2b_filename; B10 = ims_folder; B11 = ims_filename.
- Sales (109-119): B2 = sales_folder; B3 = sales_filename; B5 = annual_folder; B6 = annual_filename; B8 = sales_reco_folder; B9 = sales_reco_filename.
- Note: ITC has no file-name cell next to the GSTR-3B folder (B3 unused) - the template presumably reads the whole folder (LOW - check the template).
Used by excel_handler.py:132, 234.

**GUI_CONFIG (126-151)** - window title 'GST File Organizer v3.0', size '1200x800' (a Tkinter-style geometry string), min (1000, 700); fonts
Segoe UI 10 / 16 bold / 12 bold / 9, Consolas 9; colours primary #0078D4, success #107C10, warning #FF8C00, danger #D83B01, info #0099BC,
light #F8F9FA, dark #323130, white #FFFFFF, complete #E8F5E8, incomplete #FFF4E5. **Unused in B** (grep: only utils/__init__.py) - it was
used by the deleted Tkinter GUI (c258047 tree: gui/main_window.py:58-63, gui/handlers/client_handler.py:58-59, etc.). Several of the same colour
values are re-typed as literals in B's GUI (e.g. '#107C10', '#FF8C00', '#D83B01' in gui/views/validation_view.py:389-393, gui/pyqt_main_window.py:585-587) (HIGH).

**PROCESSING_MODES (157-173)** - keys 'fresh' (name '🆕 Fresh Run', description 'Create new folder structure (recommended for first time)',
icon '🆕'), 'rerun' ('🔄 Re-Run', 'Add new version folders to existing structure', '🔄'), 'resume' ('▶️ Resume', 'Continue previously interrupted
processing', '▶️'). In B only the **keys** are used (file_organizer.py:58 validation); names/descriptions/icons unused - the setup view re-types
similar texts (gui/views/setup_view.py:343-347). The 'resume' description does not match behaviour (see file_organizer) (HIGH).

**EXCEL_EXTENSIONS (179)** = {'.xlsx', '.xls', '.xlsm'} - unused (re-exported only). `find_excel_files` hard-codes the same three (helpers.py:255);
the GUI re-types `{'.xlsx', '.xls', '.xlsm'}` (pyqt_main_window.py:469).
**TEMPLATE_EXTENSIONS (180)** = {'.xlsx', '.xltx', '.xltm'} - unused; disagrees with the template picker filter `*.xlsx *.xlsm` (setup_view.py:612, 619).
**MAX_FILE_SIZE_MB (181)** = 100 - unused (no maximum size is enforced anywhere in my files).
**EXCEL_SIGNATURES (182-185)** = {b'PK\x03\x04': 'xlsx', b'\xd0\xcf\x11\xe0': 'xls'} - unused (re-typed in file_parser.py:197-198 and helpers.py:240-241).

**MESSAGES (191-211)** - unused in B (were used by the deleted Tk handlers, c258047 gui/handlers/file_handler.py:123-139, processing_handler.py:30, 61).
Texts: errors - no_source 'Please select a source folder containing GST files'; no_templates 'Please select both ITC and Sales template files';
no_target 'Please select a target folder for organized files'; no_selection 'Please select at least one client to process'; invalid_excel
'Invalid Excel file: {file}'; template_error 'Error loading template: {error}'; processing_error 'Error processing {client}: {error}'.
success - scan_complete '✅ Found {count} clients with {files} files'; client_complete '✅ Processed {client} successfully'; all_complete
'🎉 All processing completed! {count}/{total} successful'. info - scanning '🔍 Scanning folder for GST files...'; processing '⚙️ Processing {client}...';
creating_report '📊 Creating {report} report...'.

**EXCEL_SAFETY (217-228)** - max_retries 3, retry_delay 1 (seconds), temp_prefix 'temp_gst_', backup_prefix 'backup_', verification_size_min
10240 (comment "10KB minimum"), save_options {data_only False, keep_vba False, keep_links False}. Only assigned to `ExcelHandler.safety_config`
(excel_handler.py:54) and **never read** - so there are no retries, no temp-file prefix, no size verification; actual backups use the infix
"_backup_" (helpers.py:195), not this prefix; the openpyxl load uses `keep_links=True` (excel_handler.py:231), the opposite of `keep_links False` here (HIGH).

**STATE_CODE_MAPPING (234-275)** - [GST RULE] TIME-VARYING (state/UT list and abbreviations). 39 lower-case keys -> 36 distinct 2-letter codes
(RAN). Entries exactly: andhra pradesh AP; arunachal pradesh AR; assam AS; bihar BR; chhattisgarh CG; goa GA; gujarat GJ; haryana HR;
himachal pradesh HP; jharkhand JH; karnataka KA; kerala KL; madhya pradesh MP; maharashtra MH; manipur MN; meghalaya ML; mizoram MZ;
nagaland NL; odisha OD; orissa OD (comment "Alternative name"); punjab PB; rajasthan RJ; sikkim SK; tamil nadu TN; telangana TS; tripura TR;
uttar pradesh UP; uttarakhand UK; west bengal WB; andaman and nicobar islands AN; chandigarh CH; dadra and nagar haveli and daman and diu DD;
delhi DL; national capital territory of delhi DL; jammu and kashmir JK; ladakh LA; lakshadweep LD; puducherry PY; pondicherry PY (comment
"Alternative name"). These are letter abbreviations, **not** the two-digit numeric GST state codes; nothing in my files uses numeric codes or
GSTINs (HIGH, READ). Used only by helpers.get_state_code (helpers.py:491, 497). Not re-exported by utils/__init__.py.

#### Business rules in this file (summary)
- [GST RULE] TIME-VARYING: FILE_PATTERNS (13-62), EXPECTED_FILE_TYPES (68-75), EXCEL_TEMPLATE_MAPPING (96-120), STATE_CODE_MAPPING (234-275).
- IMS file date format DDMMYYYY (23, 27) - [GST RULE] (a date in the file name; its meaning - download date or period end - is not stated in code) TIME-VARYING candidate.
- Folder naming templates (81-90). Processing mode keys (157-173). Unused limits: 100 MB max (181), 10 KB min (222), 3 retries / 1 s (218-219).

#### Edge cases / gaps
- Many constants are dead in B (GUI_CONFIG, MESSAGES, PROCESSING_MODES texts, EXCEL_EXTENSIONS, TEMPLATE_EXTENSIONS, MAX_FILE_SIZE_MB,
  EXCEL_SIGNATURES, EXCEL_SAFETY, FOLDER_STRUCTURE['level2'], FILE_PATTERNS descriptions/category/folder) while the same values are re-typed
  elsewhere - risk that a future edit changes one copy only (HIGH).
- The state table has no entries for older separate names (e.g. "dadra and nagar haveli", "daman and diu", "uttaranchal") or for variants with
  "&"; those fall to the fallback in helpers.get_state_code (RAN: "DA", "DA", "UTT", "J&", "A&") - described only, not judged.

#### Comments, docstrings, version strings
- Docstring "Constants and configuration for GST File Organizer v3.0 / Contains all application constants, patterns, and settings." (1-4) - "v3.0"
  stale; "all application constants" is not true for B (many literals live elsewhere).
- Section banner comments only; comments "Alternative name" (255, 274), "seconds" (219), "10KB minimum" (222), "ZIP-based/OLE-based Excel files" (183-184).

---

### utils/helpers.py (537 lines) - [UTIL]

- **Purpose**: shared small functions - time stamps, path/folder helpers, file-name cleaning, safe copy and backup, Excel file checks and
  listing, display formatting, name validation, progress/ETA tracking, hashing, Windows short paths, state codes and the client-state folder key.

#### Imports
- Standard library (all unconditional): `os` (6), `re` (7; imported again inside `create_client_state_key` at 516), `time` (8), `shutil` (9),
  `hashlib` (10), `logging` (11), `pathlib.Path` (12), `datetime` (13), `typing Union, List, Optional, Dict, Any, Tuple` (14), `platform` (15).
- OS-specific: `win32api` imported **inside** `get_short_path` (477), **guarded** by `platform.system() == 'Windows'` (476) and `try/except Exception` (475, 481-483).
- Project: `.constants.STATE_CODE_MAPPING` imported inside `get_state_code` (491) - comment "Import here to avoid circular import".
- Module-level: `logger` (17). No other globals.

#### Functions (file order)

**`get_timestamp() -> str` (23-25)** - `datetime.now().strftime("%d%m%y %H%M")`, e.g. "211026 0334" (RAN). Local PC time, to the minute.
Callers: file_organizer.py:43; processing_controller.py:44.

**`get_date_only() -> str` (27-29)** - "%d%m%y". **NO CALLERS FOUND (dead)** (only re-exported).

**`get_safe_timestamp() -> str` (31-33)** - "%d%m%y_%H%M", e.g. "211026_0334". Caller: create_backup (194). (Imported but unused by excel_handler.py:37.)

**`ensure_path_exists(path) -> Path` (39-47)** - `mkdir(parents=True, exist_ok=True)`; returns the Path; on error logs and **re-raises**.
Callers: file_organizer.py:63, 95, 103, 133, 154, 177; excel_handler.py:53, 81, 315.

**`safe_path_join(*parts) -> Path` (49-64)** - drops empty/None parts (53), raises `ValueError("No valid path components provided")` if none
remain (54-55), joins the rest (57-59); errors logged and re-raised (62-64). Silently dropping an empty middle part would put a file one level
higher instead of failing (LOW; no caller passes an empty part today). RAN: `('/a','', 'b', None)` -> `/a/b`. Callers: file_organizer.py:94, 102, 132, 153, 176, 255, 314.

**`sanitize_filename(filename, max_length=200) -> str` (66-112)** - READ + RAN. Business naming rule.
1. Empty -> "unnamed" (70-71).
2. Splits into `Path(filename).stem` and `.suffix` (74-75). Note: for text that is not a file name (a client name), anything after the last dot
   counts as the "extension" and escapes all cleaning; a "/" (or "\" on Windows) is treated as a folder separator and everything before it is lost.
3. Abbreviations on the stem, **case-sensitive, anywhere in the text** (78-79): "Private" -> "Pvt", "Limited" -> "Ltd".
4. Control characters (codes 0-31 and 127-159) -> "_" (85-88).
5. Each of these characters -> "_" (82, 91-92): `< > : " / \ | ? * [ ] { } + = ! @ # $ % ^ , ; ' " ` ~` (the double quote is listed twice).
   Hyphen, "&", "(", ")" and "." are kept.
6. Runs of spaces/underscores -> one "_" (95). 7. Strips leading/trailing "_", "." and spaces (98). 8. Empty -> "unnamed" (101-102).
9. Truncates the stem to `max_length - len(extension) - 10` characters (105-106) (185 for ".xlsx" with the default 200). 10. Re-appends the extension unchanged (109-110).
- RAN examples: 'ClientA Private Limited' -> 'ClientA_Pvt_Ltd'; 'ClientA PRIVATE LIMITED' -> 'ClientA_PRIVATE_LIMITED' (no abbreviation);
  'Privateer Unlimited Co' -> 'Pvter_Unlimited_Co' (substring replaced); 'A.B. Traders' -> 'A.B. Traders' (unchanged - ". Traders" taken as the
  extension); 'Client A. Ltd' -> 'Client_A. Ltd'; 'ClientA Pvt. Ltd.' -> 'ClientA_Pvt._Ltd.'; 'M/s ClientA & Sons (Unit-1)' -> 's_ClientA_&_Sons_(Unit-1)'
  ("M/" lost); '  lead.trail.  ' -> 'lead.trail.  ' (trailing spaces kept in the "extension"); '___' -> 'unnamed'; 250 x 'a' + '.xlsx' -> 185 + '.xlsx';
  'Client+A=B!@#$%^,;...' -> 'Client_A_B'; 'GSTR3B-Client A Pvt Ltd-Tamil Nadu-April.xlsx' -> 'GSTR3B-Client_A_Pvt_Ltd-Tamil_Nadu-April.xlsx';
  extension case kept ('.XLSX'). Results for names with dots depend on Python's `pathlib` suffix rules (RAN on 3.14.6, e.g. 'Client.' has suffix '.';
  older Pythons may differ - not checked) (MEDIUM).
- Callers: file_organizer.py:86 (client name -> category folder suffix), file_organizer.py:254 (destination file name);
  processing_controller.py:112 (client part of report names). (Imported but unused by excel_handler.py:37.)
- Inconsistent with `create_client_state_key` (whole-word, any case, spaces kept) (HIGH, RAN).

**`get_file_info(file_path) -> dict` (115-136)** - name, absolute path, size, formatted size, modified, created, extension, is_file, is_dir; missing
-> {'error': 'File not found'}; exceptions -> {'error': str(e)} (logged). **NO CALLERS FOUND (dead)** since 7e09e5c (file_parser used it before). RAN: works.

**`safe_copy_file(source, destination, verify=True) -> bool` (142-185)** - READ + RAN.
- Source must exist and be a file, else logs and returns False (161-163; RAN). Creates the destination folder (166). `shutil.copy2` (169) - copies
  content and timestamps; **overwrites** an existing destination. If `verify`: compares byte sizes; on mismatch logs, **deletes the destination copy**
  and returns False (172-178). Logs INFO success (180) and returns True. Any exception -> logged, False (183-185).
- Never touches the source. No temporary name + rename, so an interrupted copy can leave a partial destination (see file_organizer gap 10) (MEDIUM).
- On some network shares `copy2`'s copying of timestamps/permissions can raise even though the content copied (LOW - not tested); the result would be
  "Copy operation failed" with a complete copy left in place.
- Caller: file_organizer.py:273.

**`create_backup(file_path) -> Optional[Path]` (187-204)** - if the file exists, copies it (copy2) to `"{stem}_backup_{DDMMYY_HHMM}{suffix}"` in the same
folder (194-198) and returns that path; missing -> None; errors -> logged, None. A second backup in the same minute **overwrites the first backup** (RAN). Caller: file_organizer.py:263.

**`validate_excel_file(file_path) -> bool` (210-247)** - exists and is a file; extension in {.xlsx, .xls, .xlsm, .xltx, .xltm} (228); size >= 1024 (232);
first 8 bytes start with ZIP or OLE signature (236-243); exceptions -> logged, False. **NO CALLERS FOUND (dead)** since 7e09e5c (imported, unused, in
excel_handler.py:38). RAN: works.

**`find_excel_files(folder) -> List[Path]` (249-266)** - globs `*.xlsx`, `*.xls`, `*.xlsm` in the folder itself only (255-257; **not sub-folders**);
removes names starting "~$" (260); returns sorted list (262) (sorted by full path; on Linux upper-case letters sort before lower-case; on Windows
paths compare case-insensitively - MEDIUM, standard Python behaviour). Any exception -> logged, returns `[]` (264-266) - so an unreachable network
folder looks like "no files" (HIGH, READ). Caller: file_parser.py:148.

**`format_size(size_bytes) -> str` (272-285)** - negative -> "Unknown"; otherwise divides by 1024 up to TB, one decimal (RAN: 1023 -> '1023.0 B',
1024 -> '1.0 KB', 10^16 -> '9094.9 TB'). Callers: file_parser.py:222, 359; helpers.py:127.

**`truncate_path(path, max_length=60) -> str` (287-309)** - short enough -> unchanged; 2 or fewer parts -> "..." + last characters; else
`first + sep + "..." + sep + last` (or "..." + sep + last). `middle_length` (304) is computed but no middle text is ever kept. On Linux the first part
"/" gives "//.../file.xlsx" (RAN). **NO CALLERS FOUND (dead)** (was used by the deleted Tk file_handler).

**`clean_windows_path(path) -> str` (311-315)** - empty/None -> ""; else `str(Path(path))` with "/" replaced by "\" (RAN). Used to write folder paths into
template cells: excel_handler.py:555, 560, 571, 582, 594, 605, 616; and helpers.py:460, 463, 467.

**`validate_client_name(name) -> (bool, str)` (321-342)** - empty/blank -> (False, "Client name cannot be empty"); length < 2 -> "Client name too short";
> 100 -> "Client name too long"; contains any of `<>:"/\|?*` -> "Client name contains invalid characters"; else (True, "") (RAN). Callers:
file_parser.py:106-107 (client and state; results go to unseen warnings).

**`get_unique_path(base_path) -> Path` (344-363)** - if free, returns it; else appends "_1", "_2", ... (before the suffix for files, after the name for
folders) up to 100, then raises `ValueError` (RAN: "..._1.xlsx", "src_1"). **NO CALLERS FOUND (dead)** - notably not used for reports/backups, which overwrite instead.

**`class ProgressTracker` (369-403)** - `__init__(total, callback=None)` (372-376) stores total, current 0, callback, start time.
`update(increment=1, message="")` (378-384) adds to current, computes percent (0 if total is 0) and calls `callback(current, total, percent, message)`
if set. `get_elapsed_time()` (386-388). `get_eta()` (390-403) - "Unknown" when current is 0, else remaining / (current / elapsed) formatted by
`format_duration`. Caller: processing_controller.py:56 (no callback), 86 (`update(0 if idx == 0 else 1, ...)` - so after client i, current = i),
87 (`get_eta()` when idx > 0), 196 (`get_elapsed_time()`). The ETA counts clients, so it is meaningful only when clients take similar time (MEDIUM).

**`format_duration(seconds) -> str` (405-416)** - "<60" -> "Ns"; "<3600" -> "Mm Ss"; else "Hh Mm" (RAN; negative input -> "-5s"). Caller: 401.

**`calculate_file_hash(file_path, algorithm='md5') -> Optional[str]` (422-440)** - reads in 8192-byte chunks; missing -> None; errors -> logged, None
(RAN). **NO CALLERS FOUND (dead)**.

**`extract_filename_without_extension(filename) -> str` (446-448)** - `Path(filename).stem` (RAN: 'A.B.xlsx' -> 'A.B'). Callers: excel_handler.py:561, 572, 583, 595, 606, 617.

**`get_relative_excel_path(from_path, to_path) -> str` (450-467)** - `to` relative to `from`'s folder if possible, else absolute; Windows-style;
errors -> logged, Windows-style `to`. **NO CALLERS FOUND (dead)** (reports get absolute paths instead).

**`get_short_path(file_path) -> str` (473-483)** - Windows: `win32api.GetShortPathName(str(path))`; other systems: unchanged string; any error
(including missing pywin32) -> logged at DEBUG, unchanged string (RAN on Linux: unchanged). Caller: excel_handler.py:298. Not re-exported.

**`get_state_code(state_name) -> str` (485-511)** - READ + RAN. [GST RULE] TIME-VARYING.
- Empty/None -> returned as is (487-488). Lower-cases and strips (494); looks up STATE_CODE_MAPPING (497-499).
- Not found: if 2+ words -> first letter of the first two words, upper-case (502-506); if one word -> **first 3 characters of the original text**,
  upper-case (508-511; comment says "first 2-3 characters" but it is always up to 3). Both log a WARNING "Unknown state '<name>', using fallback: <code>".
- RAN: 'Delhi'/'delhi'/' Delhi ' -> DL; 'DL' -> DL and 'MH' -> MH (via fallback, with a warning each time); 'Tamil Nadu' -> TN; 'TamilNadu'/'Tamilnadu' -> TAM;
  'Jammu & Kashmir' -> 'J&'; 'Jammu and Kashmir' -> JK; 'Andaman & Nicobar' -> 'A&'; 'Dadra and Nagar Haveli' -> DA; 'Daman and Diu' -> DA;
  'Uttaranchal' -> UTT; 'Telangana' -> TS; 'Orissa' -> OD; 'NCT of Delhi' -> NO; 'Delhi NCR' -> DN; 'New Delhi' -> ND; 'Chattisgarh' -> CHA; 'Gujrat' -> GUJ;
  'UP' -> UP; 'West  Bengal' (two spaces) -> WB (fallback, coincidentally equal); 'Other Territory' -> OT; 'Centre Jurisdiction' -> CJ; 'Karnataka State' -> KS; 'A' -> 'A'; '' -> ''.
- Consequences: two spellings of one state can give different codes (separate clients/folders), and two different unknown states can share a
  fallback code (e.g. "Dadra and Nagar Haveli" and "Daman and Diu" both "DA") (HIGH, RAN). Described only.
- Callers: file_parser.py:261; file_organizer.py:78; excel_handler.py:550 (unused result); processing_controller.py:113; helpers.py:521.

**`create_client_state_key(client, state, max_length=35) -> str` (513-537)** - READ + RAN. Business naming rule (level-2 folder name).
- Whole-word, any-case abbreviations (517-518): "Private" -> "Pvt", "Limited" -> "Ltd" (regex `\bPrivate\b`, `\bLimited\b`, IGNORECASE).
- `key = f"{client}-{get_state_code(state)}"` (521-524). If longer than `max_length` (527): if more than 5 characters remain for the client
  (`max_length - len(code) - 1`), cut the client to that length (530-532); otherwise cut the whole key to `max_length`, which can drop the state code
  entirely (533-535; RAN max_length 8 -> 'A Very L'; the GUI spin box allows only 15-100, gui/views/setup_view.py:375, so this branch needs an
  out-of-range value) (HIGH).
- No character cleaning (relies on names coming from file names). Cutting can leave a trailing space before the hyphen (RAN: 'A Very Long Client Name Pvt Ltd -TN').
- RAN: 'ClientA Private Limited'+'Delhi' -> 'ClientA Pvt Ltd-DL'; 'ClientA private limited'+'Maharashtra' -> 'ClientA Pvt Ltd-MH'; 'ClientA Pvt. Ltd.' -> 'ClientA Pvt. Ltd.-DL';
  'PrivateLimited Co' -> unchanged (no word boundary); two long names differing after char 32 -> identical keys 'A Very Long Client Name Pvt Ltd -TN'.
- The limit is applied whether or not the "include client name" checkbox is ticked, although the spin box is enabled only with that checkbox
  (setup_view.py:377, 382; value always read at pyqt_main_window.py:203) (HIGH, READ).
- Caller: file_organizer.py:79. Not re-exported.

#### Business rules in this file
- [GST RULE] TIME-VARYING: state-name -> code lookup and fallback (485-511).
- Name rules: Private->Pvt / Limited->Ltd (78-79 case-sensitive substring; 517-518 whole-word any case); invalid characters list (82); 200-char
  file-name limit with 10-char reserve (66, 105); 35-char client-state key (513, 527-535); client/state names 2-100 chars, no `<>:"/\|?*` (331-340).
- Time stamps DDMMYY HHMM / DDMMYY_HHMM / DDMMYY (23-33) - day-first, two-digit year, minute resolution (not FY logic).
- Excel recognition: extensions .xlsx/.xls/.xlsm (255), lock files "~$" skipped (260), 1 KB minimum and signatures (232-243, dead copy).
- Backup naming "{stem}_backup_{DDMMYY_HHMM}{suffix}" (195).
- Operational: hash chunk 8192 bytes (433); unique-path limit 100 (362); ETA by average rate (395-401).

#### Edge cases explicitly handled
- Empty names (70-71, 101-102, 328-329, 487-488); control characters (85-88); over-long names (105-106, 527-535); missing source (161-163);
  size mismatch after copy (175-178); missing file for backup/hash/info (191-192, 427-428, 119-120); Excel lock files (260); zero totals in
  progress (381) and ETA at zero (392-393); negative sizes (274-275); short path unavailable (480, 481-483); unknown states (501-511).

#### Gaps, likely failures and silent bugs
1. `sanitize_filename` treats dotted client names' tails as "extensions" and leaves them uncleaned; loses text before "/" (RAN) (HIGH for the behaviour; impact depends on real client names).
2. Two different abbreviation rules (78-79 vs 517-518) for the same client (RAN) (HIGH).
3. State fallback codes can collide or split one state into several codes (RAN) (HIGH).
4. `create_backup` overwrites a same-minute backup (RAN) (HIGH).
5. `find_excel_files` turns any error (e.g. network folder gone) into "no files" (264-266) (HIGH).
6. Dead functions: get_date_only, get_file_info, validate_excel_file, truncate_path, get_unique_path, calculate_file_hash, get_relative_excel_path (HIGH, grep).
7. `truncate_path` never shows a middle part; doubles the separator after a root (RAN) (dead anyway).
8. Duplicate `import re` (7, 516) - harmless.
- No closure-in-loop issue (`iter(lambda: f.read(8192), b'')` at 433 is the standard chunked-read idiom), no mutable defaults, no `is`/`==` issues (READ).

#### Files, folders, settings, caches and logs
- Writes: folders (43, 166); copies (169) and backups (198); deletes a destination copy on size mismatch (177).
- Reads: folder listings (257), file headers (236-237), file contents for hashing (432-434, dead).
- Logs via `utils.helpers` logger; messages include full paths/file names (180, 199, 505, 510).
- No settings or caches.

#### Comments, docstrings, version strings
- Docstring "Helper utilities for GST File Organizer v3.0 / Robust utility functions with comprehensive error handling." (1-4) - stale version.
- Banner "UPDATED: Time and Date Functions" (20) and "NEW: Short Path and State Code Functions" (470) - leftover change notes from before the initial commit.
- `get_timestamp` docstring "(DDMMYY HHMM)" agrees; `get_short_path` docstring "(Windows only)" agrees; `get_state_code` comment "take first 2-3 characters"
  (508) - always 3; `sanitize_filename` docstring "with business abbreviations and length limit" agrees; comment 81 "Define invalid characters WITHOUT including hyphen" agrees.
- No TODO/FIXME or commented-out code.

#### Pattern notes
- Two error styles: some helpers log and re-raise (ensure_path_exists, safe_path_join), most log and return a neutral value (False/None/[]/{'error':...}).
- Windows import done the guarded way here (476-483), like excel_handler.py, unlike report_processor.py:11-15.

---

## Cross-file observations

### Shared files/folders: write path vs read path
| Shared thing | Written by | Read by | Consistent? |
|---|---|---|---|
| Level-1 folder `Annual Statement-{DDMMYY HHMM}` | file_organizer.py:150-155, 173-178 (from constants.py:82) | file_organizer.py:160-169 (prefix literal 162, newest by **mtime**); gui/pyqt_main_window.py:432-436 (prefix literal, newest by **reverse name**) | **No** - two different "latest" rules; name order is not date order for DDMMYY (RAN) (HIGH) |
| Level-2 client folder | file_organizer.py:93-96 via helpers.create_client_state_key | pyqt_main_window.py:452-489 (any sub-folder; displays folder name); validation tab shows parser key instead (validation_view.py:379) | Display mismatch: folder column shows `{client}-{code}` (parser key), real folder may be abbreviated/truncated (HIGH) |
| Version folder `Version-{DDMMYY HHMM}` | file_organizer.py:99-104 | pyqt_main_window.py:456-487 (any sub-folder of level 2, sorted by reverse name; kept only if it contains a report) | Not checked for the "Version-" prefix; reverse name order not date order (HIGH) |
| Report files `ITC_Report_{client}_{code}_{DDMMYY_HHMM}.xlsx` / `Sales_Report_...` | names: processing_controller.py:114-128, 220; file: excel_handler.py:124/173 or 228/255 | pyqt_main_window.py:461-462 (`ITC_Report_*.xlsx`); report_processor.py:76-77 (`{type}_Report_*.xlsx`, takes `report_files[0]`) | Pattern matches. But (a) a failed creation leaves the template copy under this name (HIGH); (b) refreshed copies `<stem>_Refreshed_<ts>.xlsx` (report_processor.py:94) also match `ITC_Report_*.xlsx`, so on a second extraction `report_files[0]` (unordered glob) may pick a refreshed copy instead of the original (MEDIUM, READ - B-R2 to confirm) |
| Template "Links" cells | excel_handler.py:159-167 / 246-252 with constants.py:96-120 | the templates' own Power Queries (not in repo) | Cannot verify; empty values leave template text (RAN) |
| Category folder paths written into cells | excel_handler.py:555-616 (absolute Windows-style paths) | template Power Queries | Paths become stale if the output tree is moved/copied/renamed (MEDIUM) |
| File names given to templates | file_organizer.py:224-227 updates `name`; excel_handler.py:561-618 reads `name` | template Power Queries | Mismatch for skipped/failed copies (RAN) (HIGH) |
| Backups `{stem}_backup_{DDMMYY_HHMM}{suffix}` | helpers.py:194-198 | nothing (stay in category folders) | Could be read by a folder-wide Power Query (LOW-MEDIUM) |
| `_Organization_Report_{client}_{state}.txt` | file_organizer.py:313-355 | nothing in code (for people) | Statistics are cumulative across clients (RAN) |
| Summary workbook `GST_Processing_Summary_{DDMMYY_HHMM}.xlsx` | path processing_controller.py:187-188; content excel_handler.py:308-346 | nothing in code | Variations sheet always empty (processing_controller.py:73) (HIGH); failed clients absent from Client Status (HIGH) |
| `%TEMP%\gst_excel_temp` | excel_handler.py:52-53 | nothing | Dead folder (HIGH) |
| Scan results dict (`client_data`) | file_parser.py (scan thread) -> ScanResult | validation view, dry run, processing worker (which mutates `name`) | Shared mutable dicts across threads, no lock (MEDIUM) |
| `client_folder_settings` | never written in B (models/app_state.py:83 default) | file_organizer.py:126 | Always empty - feature lost in the PyQt rewrite (c258047/a62d3cc) (HIGH) |
| File-type names ('GSTR-2B Reco', 'IMS Reco', 'GSTR-3B Export', 'Sales', 'Sales Reco', 'Annual Report') | constants.py:16-56, 69-74 | file_parser.py:295, 387; file_organizer.py:289-294; excel_handler.py:558-614 | Consistent spelling everywhere (grep) (HIGH) |
| Folder keys 'gstr3b', 'itc', 'sales', 'version' | constants.py:86-88; file_organizer.py:107-111, 134 | file_organizer.py:288-298; excel_handler.py:555-616 | Consistent (HIGH) |

### Inconsistent patterns
- Three spellings of a client's name (parser key, level-2 folder key, sanitised report/category name) with two different "Private/Limited" rules (HIGH, RAN).
- OS-specific imports: guarded in core/excel_handler.py:21-28 and utils/helpers.py:475-483, unguarded in power_query_extractor/core/report_processor.py:11-15 (and line 710 local) - which makes the app unable to start without pywin32 and makes the openpyxl fallback practically unreachable (HIGH, RAN).
- Excel instance handling: excel_handler.py attaches to the user's Excel via `GetActiveObject` and quits it; report_processor.py uses `Dispatch` (174, 487) and keeps Excel visible during refresh (181) - neither uses `DispatchEx` (READ; B-R2 owns report_processor).
- Two different `_safe_open_workbook` implementations (excel_handler.py:264-306 vs report_processor.py:638-747) (HIGH).
- Signature/size checks implemented three times (file_parser.py:184-206, helpers.py:228-243, constants.py:181-185 unused) (HIGH).
- Processing-mode texts and colours re-typed in the GUI instead of using PROCESSING_MODES/GUI_CONFIG (HIGH).
- Version labels: core/utils say v3.0/3.0.0, app says v4.0 (HIGH).
- Logging: module loggers everywhere except root-logger calls at excel_handler.py:26, 28 (HIGH).

### Dead code found in my files (grep over all .py files of b0e7d0e)
- Functions/methods never called: `FileParser.validate_results`; `FileOrganizer.cleanup_on_error`, `_is_folder_empty` (only from cleanup_on_error),
  `get_summary`; `ExcelHandler.__del__` body (condition never true); helpers `get_date_only`, `get_file_info`, `validate_excel_file`, `truncate_path`,
  `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`; `_create_report_openpyxl` practically unreachable through the app.
- Parameters never used in practice: `organize_files(progress_callback)`; `FileOrganizer(client_folder_settings)` always `{}`.
- Attributes never used: `ExcelHandler.temp_dir`, `safety_config`, `excel_app`; `FileParser.errors`; parsed fields `metadata`, `warnings`,
  `category`, `folder`, `pattern`, `status_icon`, `modified`, `created`, `relative_path`.
- Constants never used: GUI_CONFIG, MESSAGES, PROCESSING_MODES (except keys), EXCEL_EXTENSIONS, TEMPLATE_EXTENSIONS, MAX_FILE_SIZE_MB,
  EXCEL_SIGNATURES, EXCEL_SAFETY (all values), FOLDER_STRUCTURE['level2'], FILE_PATTERNS 'description'.
- Imports never used: excel_handler `os`, `List`, `Optional`, `sanitize_filename`, `get_safe_timestamp`, `validate_excel_file`; file_organizer `Tuple`;
  file_parser `Optional`; constants `Tuple`. Package metadata `__all__/__version__/__author__` in both `__init__` files unused.
- Local variable computed and unused: `state_code` in `prepare_template_data` (excel_handler.py:550); `state` (file_organizer.py:87) only in a log.
- Most of these were used by the Tkinter GUI deleted in a62d3cc (07-02-2026) (READ, `git grep` on c258047): GUI_CONFIG, MESSAGES, PROCESSING_MODES texts,
  truncate_path, client_folder_settings; others (validate_results, cleanup_on_error, get_summary, get_unique_path, calculate_file_hash,
  get_relative_excel_path, get_date_only, EXCEL_SAFETY) had no callers even then.

### Behaviour summary for the merge (what happens to files)
- Original source files: only read (8-byte header + full copy); never moved, renamed, deleted or modified (HIGH, READ + RAN).
- Destination: copies with cleaned names; existing copies are backed up (fresh), overwritten (rerun) or skipped (resume, same minute only).
- Reports: template copied then filled; existing report of the same name silently overwritten; failed fills leave the raw copy.
- Temporary files: none used (the `gst_excel_temp` folder is created and left empty; EXCEL_SAFETY `temp_prefix` unused).
- Network shares: no special handling (no retries despite `max_retries: 3`; folder-listing errors become "no files"; Excel open falls back to 8.3 short paths only).

## Possible client names

None found in my seven files or in their git history on B/main (searched the full `git log -p B/main` of these files for business-name words such as
Pvt/Ltd/Limited/Traders/Enterprises/Industries/LLP/Sons/Associates; only the generic code abbreviations matched). The test names used in these notes are my own placeholders.

## Questions for the author

1. [GST RULE] Are the six file types in EXPECTED_FILE_TYPES (constants.py:68-75) still the complete set a client needs, and is "any number of GSTR-3B files, exactly one of everything else" (file_parser.py:291-297) the intended rule?
2. [GST RULE] State abbreviations (constants.py:234-275, helpers.py:485-511): are these only short labels for folder names, or are they meant to be official state codes? What should happen for spellings not in the table (today: first letters, e.g. "New Delhi" -> "ND", "Jammu & Kashmir" -> "J&")?
3. Do client names in your file names ever contain hyphens or dots (e.g. "A.B. Traders", "Pvt. Ltd.")? Hyphens break recognition (file_parser) and dots escape name cleaning (helpers.sanitize_filename).
4. [GST RULE] When a client has two GSTR-2B (or IMS, Sales, Sales Reco, Annual) files, which one should feed the report? Today: the first alphabetically (excel_handler.py:561-618).
5. [GST RULE] When a file type is missing, should its cells in the template be cleared, or left as the template has them (today: left)? (excel_handler.py:162, 249)
6. What should "Resume" do? Today it only skips files already copied in the same clock minute; otherwise it behaves like "Re-Run" (file_organizer.py:260-270, 148-179).
7. Was the per-client "include client name in folder" toggle (old Tkinter GUI) meant to survive in the PyQt version? In B it is gone; only the global checkbox works.
8. Do you ever download sources as ".xlsm", put them in sub-folders, or name them "GSTR-3B-..." / "Sales Reco-..." / "Annual Report-..."? None of these are recognised today.
9. [GST RULE] The IMS file name carries an 8-digit date (DDMMYYYY): is that the download date or a period date? (constants.py:23) The other files' period/month/year parts are not checked at all - should they be?
10. Should the app ever use an Excel window you already have open (today it attaches to it, hides it and closes it when done - excel_handler.py:110-120, 203-211)?

## INCOMPLETE

Nothing in my assigned scope is unread. Limits of what could be verified here (not omissions of reading):
- Real Excel/COM behaviour (excel_handler.py:95-306) could not be run (Linux, no Excel, no pywin32); it was traced by reading and with a fake-object simulation that only proves Python control flow. To check on a Windows test PC: open Excel with an unsaved workbook, run one client through Processing, and watch whether that Excel window disappears and closes.
- Windows-specific file behaviour (case-insensitive glob, 8.3 short paths, MAX_PATH limits, network-share copy2 behaviour) was reasoned from standard Python/Windows behaviour, not run.
- Template workbooks and their Power Queries are not in the repository, so what the Links cells feed (and whether backups/stale values are picked up) cannot be checked.
