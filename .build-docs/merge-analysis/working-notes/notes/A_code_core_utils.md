# A-R1 working notes: project A (AnnualStatement_oldv1, snapshot bf2f1dd, 27-07-2026)

Reader: A-R1. Files: `core/__init__.py`, `core/excel_handler.py`, `core/file_organizer.py`, `core/file_parser.py`, `utils/__init__.py`, `utils/constants.py`, `utils/helpers.py`. I read every line of all seven, plus the callers I needed in other files.

## How to read these notes

- **[READ]** means I read it in the code. **[RAN]** means I ran it, always on copies under `/root/merge-scratch/run/A-R1/`. Each inference carries **HIGH / MEDIUM / LOW** confidence.
- Line numbers are repository lines at bf2f1dd. Paths are repository-relative.
- Line counts: the first number is the number of lines the file has. `wc -l` gives one less for every one of these seven files, because none of them ends with a newline character.
- **[GST RULE]** marks anything that decides which return, table, sheet or cell a file or figure belongs to, a state code, a tax period, and so on. It is described only, never judged. **TIME-VARYING** marks values that can change over time.
- Examples use placeholders (`ClientA`, `<Client A>`, `ClientB` …). No client data was used or opened. I did not open `gst_organizer.log`.

### Short glossary (terms used below)
- **regex / pattern**: a text-matching rule, for example "starts with `Sales-`, then four parts separated by hyphens".
- **group**: the part of a regex match that is captured and kept, for example the client name.
- **COM**: Windows' mechanism that lets one program drive another. Here Python drives Microsoft Excel through the `pywin32` package (`win32com`, `pythoncom`).
- **openpyxl**: a Python library that reads and writes `.xlsx` files without Excel. It drops Power Query when it saves.
- **Power Query**: Excel's built-in import/refresh feature. Inside an `.xlsx` it is stored in a `customXml/item*.xml` part containing `DataMashup`, plus `xl/connections.xml`.
- **Links sheet**: the sheet in the ITC/Sales templates where the tool writes the folder paths and file names that the template's queries read.
- **stat**: a file's metadata (size, dates, whether it is a file or a folder). **mtime** is the last-modified time.
- **os.scandir / DirEntry**: one directory listing that returns every entry with its metadata. On Windows this costs no extra network round-trip.
- **thread / worker thread**: a second line of work running at the same time as the window. A **ThreadPoolExecutor** is a pool of helper threads.
- **UI thread**: the thread that runs the Tk window. While it is busy, the window does not redraw.
- **callback**: a function handed to other code so that code can "call back", for example to report progress.
- **exception**: an error raised while the program runs. **Swallowed** means caught and ignored, with or without a log line.
- **bare `except:`**: catches every possible error, including Ctrl+C.
- **8.3 short path**: old DOS-style short names such as `PROGRA~1`.
- **UNC path**: a network path written as `\\server\share\...`.
- **dict / defaultdict**: a lookup table. A defaultdict creates an empty entry automatically when a missing key is used.
- **Treeview**: the Tk table widget that shows the client list. **IntVar**: a Tk variable holding a whole number.
- **content type**: the label inside an `.xlsx` package that says whether it is a workbook, a template or a macro workbook.
- **re-export / `__all__`**: a package file that lets `from package import X` work. The names must be listed by hand.

### Experiments I ran (scripts kept in `/root/merge-scratch/run/A-R1/`)
| Script | What it exercised |
|---|---|
| exp1_parse.py | `FileParser.parse_filename` on 33 synthetic file names |
| exp2_scan.py | `FileParser.scan_folder` on a synthetic folder: valid files, invalid files, duplicate types, Excel lock file, subfolder, folder named like a workbook, state-name and abbreviation variants |
| exp3_helpers.py, exp3b_pyver.py | every helper in `utils/helpers.py`; `sanitize_filename` under Python 3.12.3, 3.13.16 and 3.14.6 |
| exp4_organize.py | `FileOrganizer` in fresh, rerun and resume modes, collisions, the include-client-name option, the organisation report |
| exp5_excel_openpyxl.py | `ExcelHandler` report creation through the openpyxl path, the Power Query parts, a `.xltx` template, the summary workbook |
| exp6_fake_com.py, exp6b_reraise.py | `ExcelHandler`'s Excel-through-COM path driven with fake COM objects. This proves only the Python control flow, not what real Excel does |
| exp7_tree.py (xvfb) | Tk Treeview round-trip of numeric-looking client and state values |
| exp8_resume_names.py | file-name mismatch in resume mode |
| project tests | copied `tests/conftest.py`, `test_helpers.py`, `test_file_ops.py`, `test_file_patterns.py`, `test_scan.py`, `test_excel_com.py` and ran them: **98 passed** [RAN]. No test covers `core/file_organizer.py` or `core/excel_handler.py`. The grep found only a mention of excel_handler in a docstring of `tests/test_optional_pywin32.py:4` [READ] |

Git: every commit that touched these files, on A/main (all main branch):
- `core/__init__.py`: edac6a3 only.
- `core/excel_handler.py`: edac6a3, 12d505f, 064c634, ba2a753.
- `core/file_organizer.py`: edac6a3, 275d9a4.
- `core/file_parser.py`: edac6a3, 6237597, a603c02, 064c634, 56aaad9.
- `utils/__init__.py`: edac6a3, 56aaad9.
- `utils/constants.py`: edac6a3, 5c8ce95, 7347383, 510f7b4.
- `utils/helpers.py`: edac6a3, 6237597, a603c02, 33d0baf, 4219ffb, 56aaad9, 4f84ba9.

So before the split (47a1492) only edac6a3 and 6237597 touched these files. Every later change is from 23–26 July 2026; the two March 2026 commits did not touch them. The July commit messages carry "Co-Authored-By: Claude Opus 4.8" [READ].

---

### core/__init__.py (13 lines) - [CORE]

- **Purpose:** marks `core` as a Python package and re-exports its three classes. Python runs this file whenever any `core.*` module is imported, so importing any one core module also imports all three.
- **Imports:** project only: `from .file_parser import FileParser` (6), `from .file_organizer import FileOrganizer` (7), `from .excel_handler import ExcelHandler` (8). No OS-specific modules directly. Importing `excel_handler` brings in its guarded `win32com`/`pythoncom` import and its *unconditional* `openpyxl` import (see that file) [READ].
- **Module-level constants:** `__all__ = ['FileParser', 'FileOrganizer', 'ExcelHandler']` (10). `__version__ = '3.0.0'` (12). `__author__ = 'GST File Organizer Team'` (13).
- **Classes/functions:** none.
- **Who uses it:** nobody writes `from core import …` (grep of all .py files: none) [READ]. The file still runs as a side effect of `from core.file_parser import FileParser` (gui/main_window.py:15-17, tests/test_scan.py:9, tests/test_file_patterns.py:10). Consequences [READ, HIGH]:
  - importing only the parser still imports openpyxl;
  - it runs `warnings.filterwarnings('ignore')` (core/excel_handler.py:17);
  - it runs the Windows/pywin32 probe (core/excel_handler.py:20-28).
- **Business rules:** none.
- **Edge cases handled:** none.
- **Gaps:**
  - The re-export list must be kept in step with the modules by hand.
  - Import side effects as above.
  - The version string `3.0.0` matches `utils/__init__.py:38` and the window title "GST File Organizer v3.0" (utils/constants.py:134). Project B calls itself v4.0.
- **Excel touchpoints / files written:** none.
- **Comments/docstrings:** "Core package for GST File Organizer v3.0 / Contains business logic modules." (1-4). This matches the code.
- **Pattern notes:** the same re-export style is used in `utils/__init__.py` and `gui/__init__.py`.

---

### core/excel_handler.py (713 lines) - [CORE]

- **Purpose:** for each client, makes the ITC or Sales report from the chosen template. It copies the template to a local temp file, writes folder paths and file names into the template's "Links" sheet (through Excel by COM on Windows, otherwise openpyxl), and copies the result to the output folder. It also builds the run's summary workbook with openpyxl.

#### Imports
- Standard library:
  - `os` (6): **unused**.
  - `time` (7), `shutil` (8), `logging` (9), `tempfile` (10), `pathlib.Path` (11).
  - `typing` Dict, List, Optional, Any, Union (12): List and Optional **unused**.
  - `warnings` (13), `platform` (14).
  - (An AST scan, i.e. a parse of the code's structure, was RAN.)
- **OS-specific:**
  - `win32com.client` (23) and `pythoncom` (24) are **guarded**: tried only when `platform.system() == 'Windows'` (21), inside `try/except ImportError` (22-28).
  - On success `USE_WIN32COM = True` (25) and the *root* logger logs INFO "Using win32com for Excel operations (Power Query support)" (26). On ImportError the root logger logs WARNING "win32com not available. Install with: pip install pywin32" (28).
  - These two lines use `logging.info/warning` directly instead of the module `logger` (43). This is a small pattern deviation.
  - `win32api` is imported lazily inside `utils/helpers.get_short_path` (helpers.py:584), guarded by a platform check and try/except.
- Third-party, **unconditional**: `from openpyxl import load_workbook, Workbook` (31), `from openpyxl.styles import Font, PatternFill, Alignment, Border, Side` (32), `from openpyxl.utils import get_column_letter` (33). The comment at 30, "Fallback to openpyxl if not on Windows", is only half true: openpyxl is always imported and is always used for the summary workbook.
- Project:
  - `utils.constants`: `EXCEL_TEMPLATE_MAPPING`, `EXCEL_SAFETY` (35).
  - `utils.helpers` (36-40): `sanitize_filename` (**unused**), `get_safe_timestamp` (**unused**), `ensure_path_exists`, `validate_excel_file` (**unused**), `clean_windows_path`, `extract_filename_without_extension`, `get_short_path`, `get_state_code`. The comment `# ADD THESE TWO` at 39 is a leftover pasting instruction.
  - `utils.excel_com.find_open_workbook` (41).

#### Module-level constants and globals
- `warnings.filterwarnings('ignore')` (17): turns off **all** Python warnings for the **whole program** from the moment this module is imported. It is process-wide, not limited to this file. That includes openpyxl's warnings about unsupported workbook features and any deprecation warnings [READ, HIGH].
- `USE_WIN32COM = False` (20); set to True at 25.
- `logger = logging.getLogger(__name__)` (43).
- Operational values written inline (no named constant):

| Value | Line(s) |
|---|---|
| temp folder name `"gst_excel_temp"` | 53 |
| `time.sleep(0.5)` | 155 and 210 (1 second per COM report) |
| path-length threshold `218` | 356 |
| COM error text `"-2147352567"` (the generic COM "exception occurred" code; MEDIUM on the meaning) | 355 |
| default sheet `'Links'` | 163, 286 |
| fallback sheet-name substring `'link'` | 175 |
| log preview length `value[:50]` | 195 |
| summary styles: header fill `"366092"` (dark blue), white font `"FFFFFF"`, sizes 12 and 16 | 414-423 |
| auto-fit width cap 50, padding +2 | 600 |

#### Class `ExcelHandler` (46-713)

**`__init__(self)` (51-56)**
1. `self.temp_dir = Path(tempfile.gettempdir()) / "gst_excel_temp"` (53). On Windows this is normally `%TEMP%\gst_excel_temp`, i.e. under `C:\Users\<user>\AppData\Local\Temp` (MEDIUM: it depends on the TEMP/TMP environment variables).
2. `ensure_path_exists(self.temp_dir)` (54) creates the folder. It is not caught, so if TEMP is unusable the app fails to start: the instance is created at start-up.
3. `self.safety_config = EXCEL_SAFETY` (55): **never read anywhere** (grep) → dead.
4. `self.excel_app = None` (56): **never assigned anything else anywhere** in the project (grep) [READ, HIGH]. See `__del__` and the `finally` block below.
- Caller: gui/main_window.py:116 `self.excel_handler = ExcelHandler()`. One instance per app session, created on the UI thread, used later from the processing worker thread.

**`__del__(self)` (58-64)**
- If `USE_WIN32COM and self.excel_app`, calls `Quit()` inside a bare `except: pass`. Because `excel_app` is always None, this never does anything [READ, HIGH].
- Called by Python when the object is discarded.

**`create_report_from_template(self, template_path, output_path, data_mappings, report_type) -> bool` (66-102)**
1. DEBUG logs (73-76). Line 76 would dump the whole mapping dict, with paths that contain client names. DEBUG is not written at the app's INFO level.
2. Converts both paths to Path (78-79).
3. If the template does not exist: ERROR log with a leftover "[DEBUG]" prefix (83), then **raises** `FileNotFoundError("Template not found: …")` (84). It raises rather than returning False.
4. DEBUG log of the template size (85). This still does a `stat()` of the template.
5. `ensure_path_exists(output_path.parent)` (87) creates the output folder at any depth.
6. If `USE_WIN32COM`: `return self._create_report_win32com(...)` (90-95).
7. Otherwise: WARNING "Not on Windows - Power Query may not be preserved!" (98), then `return self._create_report_openpyxl(...)` (100-102). The text says "Not on Windows" even when the real cause is Windows without pywin32.
- Returns True/False. Can raise FileNotFoundError, anything `ensure_path_exists` raises, and (COM path) an exception escaping the `finally` block (see below).
- Callers:
  - gui/handlers/processing_handler.py:309 (ITC). Template `app.itc_template`; output `folders['version'] / "ITC_Report_<sanitized client>_<state code>_<DDMMYY_HHMM>.xlsx"` (processing_handler.py:292-298).
  - gui/handlers/processing_handler.py:376 (Sales). Output `Sales_Report_…` (355-356).
  - Both run on the processing worker thread, inside try/except blocks that count failures in `summary_data['report_errors']`.

**`_create_report_win32com(self, template_path, output_path, data_mappings, report_type) -> bool` (104-260)**, step by step:
1. DEBUG logs (109-112).
2. `pythoncom.CoInitialize()` (120) sets COM up for the current (worker) thread.
3. Getting Excel (123-128):
   - `win32com.client.GetActiveObject("Excel.Application")` **attaches to an Excel the user already has open** (INFO "Using existing Excel instance").
   - A bare `except:` falls back to `win32com.client.Dispatch("Excel.Application")` (INFO "Created new Excel instance").
   - `Dispatch` (not `DispatchEx`) may itself connect to a running Excel (MEDIUM; needs real Excel to confirm).
4. `excel.Visible = False`, `excel.DisplayAlerts = False`, `excel.ScreenUpdating = False` (131-133). These are applied to the user's own Excel too, which hides the user's Excel window [RAN with fakes: `Visible` is still False after the run].
5. Local staging (135-155):
   - `local_temp = self.temp_dir / output_path.name` (147).
   - If it already exists, it is deleted and errors are swallowed (148-152).
   - `shutil.copy2(template → local_temp)` (154).
   - `time.sleep(0.5)` (155). The comment "Give Windows time to complete the copy" does not fit the code: `copy2` has already finished when it returns, so the wait is not needed for correctness (MEDIUM).
6. `wb = self._safe_open_workbook(excel, local_temp)` (159).
7. `mapping_config = EXCEL_TEMPLATE_MAPPING.get(report_type, {})` (162); `sheet_name = mapping_config.get('sheet', 'Links')` (163). An unknown report type gives `{}`: sheet 'Links' and no cells.
8. Finding the sheet (166-183):
   - (a) exact name, case-insensitive (167-170);
   - (b) else the **first** sheet whose name *contains* "link", case-insensitive (for example "Hyperlinks"), with WARNING "Using sheet '…' instead of 'Links'" (173-178);
   - (c) else ERROR "Links sheet not found! Creating one..." and `wb.Sheets.Add()`, renamed 'Links', and the run carries on (180-183). It later returns True.
   - [RAN with fakes: all three branches.]
9. Writing the cells (186-199): for each `cell_ref, data_key`, `value = data_mappings.get(data_key, '')`.
   - **Only if the value is non-empty**: `ws.Range(cell_ref).Value = str(value)`, with INFO "  B2 = <first 50 chars>..." (195).
   - A failure on one cell gives an ERROR log and the loop continues (196-197).
   - **Empty values leave the template's existing cell content in place** [READ HIGH; RAN on the openpyxl path, which uses the same `if value:` logic].
10. `wb.Save()` (203), then `wb.Close(SaveChanges=True)` (206): saved twice, which is harmless. `wb = None` (207). `time.sleep(0.5)` (210).
11. `shutil.copy2(local_temp → output_path)` (215) **overwrites** any existing file at the output path, with no check and no backup.
12. If `output_path.exists()`: INFO with the size, return True (218-221). Else ERROR "Output file not found!", return False (222-224).
13. `except Exception` (226-228): ERROR with traceback, return False.
14. `finally` (230-260):
   - If `wb` is still set (error path): `wb.Close(SaveChanges=False)` inside a bare `except: pass` (232-236).
   - If `excel`:
     - `excel.ScreenUpdating = True` (239) and `excel.DisplayAlerts = True` (240) are **not protected by try**. If Excel has died ("RPC server unavailable"), these raise and the exception **escapes the `finally`**. Then `Quit`, `CoUninitialize` and the temp-file clean-up are all skipped [RAN with fakes: result "RAISED", temp file left in `gst_excel_temp`].
     - Then `if self.excel_app is None:` (242) is **always true**, so `excel.Quit()` (244) runs every time. **Excel is always quit, including the user's own Excel obtained at 124**, despite the comment "Don't quit if Excel was already open" (241) [READ HIGH; RAN with fakes: `USER.Quit()` was called].
     - `DisplayAlerts` was switched **back on just before Quit** (240). Standard 12 (rebuild_pipeline_FINAL.md lines 851-855) warns that without `DisplayAlerts = False`, `Quit()` can wait forever on an invisible "save changes?" prompt.
     - `Visible` is never restored.
   - `pythoncom.CoUninitialize()` in a bare except (249-252). It runs while local variables (`excel`, `ws`, `sheet`) still hold COM objects; Standard 12 (lines 872-875) asks for the Excel work to sit in a separate function so nothing is held at that point.
   - The local temp file is removed, best effort (255-260).
- Returns True/False (or raises, see 14).
- Swallowed errors: bare `except:` at 126, 235, 245, 251; `except Exception: pass` at 151, 259.
- Side effects: creates and deletes a temp file; writes the output file; may start Excel and always quits it; changes Excel application settings.
- Runs on the processing worker thread (processing_handler.py:129-134).
- Each COM report starts Excel (if none is running), opens, saves and quits it. Two reports per client means two Excel start/quit cycles per client, plus a fixed 1 second of sleeps per report [READ, HIGH].
- Caller: `create_report_from_template` (93).

**`_create_report_openpyxl(self, template_path, output_path, data_mappings, report_type) -> bool` (262-322)**
1. WARNING "Using openpyxl - Power Query content may be lost!" (267).
2. Local staging with the same temp name; a stale copy is deleted, errors swallowed (273-279).
3. `load_workbook(str(local_temp), data_only=False, keep_links=True)` (282). `keep_vba` is not set, so macros are dropped from `.xlsm`/`.xltm` templates.
4. Mapping (285-286). The sheet is looked up by **exact case-insensitive name only** (289-293). There is **no** "contains 'link'" fallback and **no** sheet creation; this differs from the COM path.
5. If the sheet is found: each non-empty value is written as text (297-301). A per-cell error is swallowed **silently** by a bare `except: pass` (302-303).
6. `wb.save(local_temp)` (306), `wb.close()` (307), `shutil.copy2(local_temp → output_path)` (308), `return output_path.exists()` (310). **When no Links sheet is found, nothing is written and it still returns True** [RAN: a template without Links gives True; a template whose only sheet is "Hyperlinks" gives True with nothing written].
7. `except Exception` gives ERROR "Fallback method error: …" and returns False (312-314). `finally` removes the temp file (315-322).
- [RAN] facts:
  - (a) **Power Query is dropped.** A template with `customXml/item1.xml` (DataMashup) and `xl/connections.xml` came out with neither part.
  - (b) A **`.xltx` template gives an output named `.xlsx` whose internal content type is still "template"** (`application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml`). I could not check whether Excel opens such a file (no Excel here). MEDIUM that Excel refuses it with "file format or file extension is not valid". How to check: run once with an `.xltx` template on a machine where this path runs, and open the output in Excel.
  - (c) An existing output file is overwritten.
  - (d) Empty mapping values leave the template's old cell text (cells B7, B8, B10 and B11 kept their "OLD-…" values).
- When this path runs: only when `USE_WIN32COM` is False, i.e. not Windows, or Windows without pywin32. Since be6418c (26-07-2026), start-up requires pywin32 on Windows (A-R4's files), so in production this path should run only if that check is bypassed (MEDIUM).
- Caller: `create_report_from_template` (100).

**`_safe_open_workbook(self, excel, file_path)` (324-369)** (the "already-open workbook" detection lives here)
1. `file_path = Path(file_path).absolute()` (326).
2. If the file does not exist: ERROR, then raise FileNotFoundError (329-331).
3. Already-open check (337-340): `find_open_workbook(excel, file_path)` (utils/excel_com.py:19-29) loops over every workbook open in that Excel and returns the first whose `Path(wb.FullName).absolute()` equals `file_path`, else None.
   - If found: INFO "File already open, using existing workbook", and that workbook is returned, with whatever content it has in Excel's memory, **not** the fresh copy just written.
   - Practical reach: the file is the temp copy written at line 154. On Windows, Excel locks files it has open, so the delete at 150 and the copy at 154 would normally fail first if that temp file were open. So this branch is effectively unreachable for ExcelHandler (MEDIUM). It matters for ReportProcessor (A-R2).
   - Comparison caveat (READ, MEDIUM): the paths are compared as text after making them absolute. On Windows, `WindowsPath` compares case-insensitively, but an 8.3 short form (some TEMP variables look like `C:\Users\ABCDEF~1\...`) against Excel's long form, or a mapped drive against a UNC path, would not match.
4. Logs the path length (342-344).
5. `excel.Workbooks.Open(file_str)` (347-349). There is no `IgnoreReadOnlyRecommended`, no `UpdateLinks` and no `ReadOnly` check; Standard 12 (lines 857-859) asks for `IgnoreReadOnlyRecommended=True` and a ReadOnly check.
6. If Open fails (350-369):
   - `error_code = e.args[0]` is logged (351-352).
   - If the message contains "-2147352567", or contains "path" (any case), or the path is longer than 218 characters, it tries `get_short_path` (8.3 name via `win32api.GetShortPathName`, helpers.py:580-590). If the short form differs, it opens that and returns it (359-364). Failures there are logged (365-366).
   - Then a bare `raise` **re-raises the ORIGINAL first error** [RAN with fakes: "failure #1" was re-raised after the short-path attempt also failed].
- Returns a COM workbook object or raises.
- Caller: `_create_report_win32com` (159). The calls at power_query_extractor/core/report_processor.py:241 and :542 are to ReportProcessor's own method of the same name (A-R2).

**`create_summary_report(self, output_path, report_data) -> bool` (371-409)**
1. `Workbook()` (381); removes the default "Sheet" (384-385); `styles` (388).
2. Creates sheets in this order: Summary, Client Status, File Mapping, Errors, Variations (391-395).
3. Auto-fits every sheet (398-399).
4. `wb.save(str(output_path))` (402) **directly on the target**, which is usually the network share. There is no local staging, unlike the reports. Then `wb.close()` and `return output_path.exists()`.
5. Any exception: ERROR with traceback, return False (407-409).
- [RAN] Sheets and contents as described below.
- Caller: processing_handler.py:469, with path `level1_folder / "GST_Processing_Summary_<DDMMYY_HHMM>.xlsx"` (processing_handler.py:462-464), and only if at least one client succeeded (456).

**`_create_styles(self)` (411-430)**: returns:
- `header`: bold 12-point white font, solid fill "366092", centred horizontally and vertically;
- `title`: bold 16-point font, and an alignment that is never applied;
- `border`: thin on all four sides.
- Caller: 388.

**`_create_summary_sheet(self, wb, data, styles)` (432-462)**
- Sheet "Summary". A1 = "GST FILE PROCESSING SUMMARY", bold 16, merged A1:B1.
- From row 3, a bold label in A and its value in B (443-454):

| Label | Value |
|---|---|
| "Processing Date" | `timestamp` |
| "Total Clients" | `total_clients` |
| "Successful" | `successful_clients` |
| "Failed" | `failed_clients` |
| "Total Files" | `total_files` |
| "Reports Generated" | `reports_generated` |
| "ITC Reports Created" | `itc_reports_created` |
| "Sales Reports Created" | `sales_reports_created` |
| "Report Creation Errors" | `report_errors` |
| "Processing Mode" | `processing_mode` |
| "Target Folder" | `target_folder` |
| "Include Client Name in Folders" | "Yes" / "No" |

- The `# NEW` markers at 449-451 are leftovers.

**`_create_client_status_sheet` (464-491)**
- Headers "Client", "State", "Status", "Files", "Missing", "Complete %", styled with borders.
- One row per entry of `data['clients']`: client, state, status, file_count, the *number* of missing files, and completeness as **text** `f"{…:.1f}%"`, for example "100.0%" (485). Because it is text, Excel cannot sort or add it as a number.
- A non-numeric completeness makes the whole summary fail and return False [RAN g2]. Completeness is always numeric in practice.
- Only successful clients are passed in (processing_handler.py:420-429), so failed clients appear only on the Errors sheet.

**`_create_file_mapping_sheet` (493-519)**: headers "Original File", "Client", "Type", "Destination", "Status". Rows come from `data['file_mappings']`, where `client` is the client *key* "client-STATECODE" (processing_handler.py:432-438).

**`_create_errors_sheet` (521-552)**: headers "Time", "Client", "Type", "Error", "Action". With no errors, A2 = "No errors encountered", merged A2:E2. Otherwise one row per error, with "Action" defaulting to "Review".

**`_create_variations_sheet` (554-585)**: headers "Filename", "Issue", "Expected Pattern", "Size", "Action". With none, A2 = "No file variations found", merged. Otherwise "Size" is raw bytes and "Action" is always "Rename file".

**`_autofit_columns(self, ws)` (587-601)**: column width = min(longest text in the column + 2, 50). A bare `except: pass` covers each cell (597-598). Caller: 399. The call at power_query_extractor/core/data_consolidator.py:169 is to that class's own method.

**`prepare_template_data(self, client_info, folders, report_type) -> Dict[str,str]` (603-713)**
1. DEBUG logs (609-612).
2. `state_code = get_state_code(client_info['state'])` (616) is logged (617) and **never used** (dead). It sits *outside* the `try`, so a missing 'state' key raises KeyError to the caller [RAN].
3. ITC (620-656):
   - `gstr3b_folder` = the GSTR3B folder, **always** (622), even if the client has no GSTR-3B files.
   - Annual Report present: `annual_folder` = the **Version** folder and `annual_filename` = the stem (name without extension) of the **first** Annual Report file. Otherwise both are '' (625-633).
   - GSTR-2B Reco: the ITC folder plus the first file's stem (636-644).
   - IMS Reco: the ITC folder plus the first file's stem (647-655).
4. Sales (657-708):
   - Sales: the Sales folder plus the first stem;
   - Annual Report: the Version folder plus the first stem;
   - Sales Reco: the Sales folder plus the first stem;
   - '' when a file type is absent.
5. Any other `report_type` gives `{}` [RAN].
6. `except Exception` (710-711): ERROR log, then returns whatever was filled in so far.
- What "first file" means: the order in `client_info['files'][type]` is the order files were added during the scan, which is sorted path order (helpers.py:361), i.e. alphabetical. **`…-FY (1).xlsx` sorts before `…-FY.xlsx`**, because a space sorts before a dot, so a "(1)" duplicate is the one linked [RAN exp2]. **[GST RULE]**: this decides which return file's figures feed the report.
- Which name is used: `name` is the name *after* `organize_files` replaced it with the cleaned copy name (file_organizer.py:243), but only for files copied with status "Success". In **resume mode a "Skipped" file keeps its original name**, so the Links cell gets a name that differs from the actual copy whenever cleaning changed it [RAN exp8: Links B8 would be `GSTR-2B-Reco-Client A-Delhi-FY` while the folder holds `GSTR-2B-Reco-Client_A-Delhi-FY.xlsx`].
- The values are absolute folder paths built from the target-folder text chosen in the GUI, as either a mapped drive letter or a UNC path. A report refreshed on another PC needs the same path to work (MEDIUM; depends on how the template's queries use it).
- File names are written **without extension** (`extract_filename_without_extension`), so the template must know the extension. The patterns accept `.xls` as well as `.xlsx` for every type (template-dependent; see Questions).
- Callers: processing_handler.py:305 (ITC) and :368 (Sales).

#### Business rules in this file
- **[GST RULE]** Which Links cell receives which folder or file name: `EXCEL_TEMPLATE_MAPPING` (constants.py:102-126), applied at 162/189 (COM) and 285/297 (openpyxl), with keys built at 620-706. TIME-VARYING (it depends on the template layout).
- **[GST RULE]** The **first** file of each type is the one linked (629, 640, 651, 667, 683, 699).
- **[GST RULE]** GSTR-3B is linked **by folder only** (622, cell B2), never by file name. The ITC report never receives a GSTR-3B file name.
- **[GST RULE]** The Annual Report folder written is the Version folder (627, 681), because the organiser puts Annual Reports there (file_organizer.py:311).
- Sheet name 'Links', case-insensitive. The COM path alone falls back to any sheet containing "link", or creates one (166-183).
- An empty value leaves the cell as it is in the template (192, 300).
- Summary-workbook layout (432-585).
- 218-character path threshold for the 8.3 retry (356): technical.

#### Edge cases explicitly handled
- Template missing: 82-84.
- Output folder missing: 87.
- Stale temp file: 148-152, 274-278.
- Sheet-name case: 168, 291.
- No Links sheet: the COM path creates one (180-183); the openpyxl path skips writing (295).
- Per-cell failures: 196-197, 302-303.
- Open failure with a long or odd path: 8.3 retry (355-366).
- Empty error and variation lists: 537-539, 570-572.
- Default sheet removed: 384-385.

#### Gaps, likely failures and silent bugs
1. **Excel is always quit, including the user's own Excel** (241-246; `excel_app` never set, 56) [READ HIGH, RAN with fakes]. This has been the case since edac6a3 (initial commit). What it does to unsaved user work in real Excel (closed with a prompt, or a hang on a hidden prompt because `DisplayAlerts` is True at 240): MEDIUM, needs a Windows test. How to check: open Excel with an unsaved new workbook, process one client in Step 3, and watch that window.
2. **It takes over and hides the user's Excel** (124, 131), and `Visible` is never restored [HIGH; RAN with fakes]. Standard 12 (line 853) says "Never take over the user's own open Excel".
3. **The `finally` block can raise** (239-240). Quit, CoUninitialize and temp clean-up are then skipped, a stray hidden Excel can be left running (MEDIUM), and the caller counts "ITC/Sales Report Error" even though the report may already exist (the copy at 215 happened before the return) [HIGH, RAN with fakes].
4. **Template cells are not cleared** when the client lacks that file (192/300) [HIGH, RAN]. Whether the real templates contain stale sample paths in those cells is unknown; the templates are not in the repository.
5. A **missing Links sheet still counts as success** (183, 295) [HIGH]. The processing log hint "Check if template has 'Links' sheet" (processing_handler.py:323, 397) only appears when False is returned, which never happens for this cause.
6. The COM path and the openpyxl path behave differently on sheet fallback (174-183 vs 289-295) [HIGH].
7. The openpyxl path drops Power Query and macros; `.xltx`/`.xltm` content types survive into `.xlsx`-named files [HIGH, RAN]. Whether Excel refuses those files: MEDIUM.
8. On the COM path with a `.xltx`/`.xltm`/`.xlsm` template (all accepted by gui/utils/setup_validation.py:21), the template bytes are copied into a file named `.xlsx` (147, 154) and opened by Excel. Not verified: Excel may refuse a template or macro package under an `.xlsx` name, and `wb.Save()` keeps whatever format Excel opened (LOW/MEDIUM). How to check: run one client with an `.xltx` template on Windows.
9. The output is **overwritten silently** (215, 308) [HIGH].
10. The temp-file name is just the output file name (147, 273). Two copies of the app building the same report name (same client, state and minute) would collide (LOW). Two copies may also attach to the same Excel through `GetActiveObject`, and one may quit it while the other is using it (MEDIUM).
11. **Partial output on a network failure**: if `copy2` at 215/308 fails half-way, a truncated file stays at the output path. False is returned but the file is not deleted [READ HIGH].
12. The summary is written straight to the network (402); a failure mid-save may leave a partial file (MEDIUM).
13. `warnings.filterwarnings('ignore')` is process-wide (17) [HIGH].
14. Dead or unused code [HIGH, AST scan RAN]:
    - unused imports `os`, `List`, `Optional`, `sanitize_filename`, `get_safe_timestamp`, `validate_excel_file`;
    - `self.safety_config` (55);
    - `state_code` (616);
    - `__del__` (58-64), which never acts;
    - `self.excel_app` (56).
15. **Client data in the log at INFO level**: 153 (temp path with client name), 158, 195 (cell values with client names in paths and file names), 214, 220. `main.py:20-27` sets INFO and writes `gst_organizer.log` in the current folder [HIGH].
16. A leftover "[DEBUG]" prefix in an ERROR message (83); the others were converted in 064c634 [HIGH].
17. Standard 12 deviations (rebuild_pipeline_FINAL.md 846-875) [READ HIGH]:
    - `GetActiveObject`/`Dispatch` instead of `DispatchEx`;
    - `DisplayAlerts` set back to True before `Quit()`;
    - no `IgnoreReadOnlyRecommended` and no ReadOnly check;
    - `CoUninitialize` while COM objects are still held;
    - the openpyxl fallback saves over a template that may contain Power Query with only a warning, instead of checking for DataMashup and stopping.
18. `ws.Range(...).Value = str(value)` on the COM path: Excel may turn text that looks like a number or date into one. That only matters if a folder path or file stem looks numeric (LOW). How to check: write a digits-only stem through COM and look at the cell.
19. Closure late-binding: none. Mutable defaults: none. `== / is None` usage: correct (242, 340).
20. COM truthiness: `if not ws:` (172, 180) and `if wb:` (232) test COM objects for "truthiness" (whether Python treats them as true). With pywin32 objects this is normally fine (LOW concern).

#### Excel workbook touchpoints
| Line | Action | Workbook |
|---|---|---|
| 124 / 127 | GetActiveObject / Dispatch Excel.Application | the user's running Excel, or a new one |
| 154 | file copy (not open) | template → `%TEMP%\gst_excel_temp\<output name>` |
| 349 / 364 | `Workbooks.Open` (normal, then 8.3 path) | local temp copy |
| 337 | iterate `excel.Workbooks` (read `FullName`) | every workbook open in that Excel |
| 182 | `Sheets.Add` | local temp copy |
| 193 | write cells | Links sheet of the temp copy |
| 203 / 206 | Save, then Close(SaveChanges=True) | temp copy |
| 234 | Close(SaveChanges=False) on the error path | temp copy |
| 244 | `Quit()` | that Excel instance (always) |
| 215 | file copy | temp copy → output path (ITC_Report_… / Sales_Report_… in the Version folder) |
| 282 / 306 / 307 | openpyxl load / save / close | temp copy (openpyxl path) |
| 308 | file copy | temp → output (openpyxl path) |
| 381 / 402 / 403 | openpyxl create / save / close | `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` in the Level-1 folder |

#### Files and folders read or written
- `Path(tempfile.gettempdir()) / "gst_excel_temp"`: created at app start (54) and never removed.
- `…/gst_excel_temp/<output file name>`: written at 154/279, deleted at 258/320. It is left behind if the `finally` raises.
- Output report at the path given by the caller (215/308).
- Summary workbook at the path given by the caller (402).
- Logging through `logger` (root handlers: `gst_organizer.log`, relative to the current folder, set in main.py:24).

#### Comments, docstrings, version strings
- Module docstring: "Excel Handler for GST File Organizer v3.0 / WINDOWS VERSION - Uses Excel COM to preserve Power Query and all Excel features" (1-4). Mostly true; the openpyxl fallback does not preserve Power Query.
- 135-146: the staging rationale. "Power Query is preserved: the report's data sources are ABSOLUTE paths written into the Links sheet". This cannot be verified without the templates. "Saving locally then copying yields a byte-identical file at the destination" is plausible (MEDIUM).
- 155: "Give Windows time to complete the copy". The wait is unnecessary (see above).
- 241: "Don't quit if Excel was already open". **Not implemented.**
- 333-336: NOTE that the open-with-fallback logic deliberately differs from ReportProcessor's. This agrees with utils/excel_com.py:1-15.
- 380: "Always use openpyxl for summary (no Power Query needed)". True.
- 449-451: `# NEW` markers. 39: `# ADD THESE TWO`.

#### Pattern notes
- Errors: "log and return False" plus many bare excepts. A template-not-found error *raises* while everything else returns False, which is inconsistent.
- Logging: the module logger, except the root-logger calls at 26/28.
- Config comes from `utils.constants`.
- Threading: called from the worker thread; `CoInitialize`/`CoUninitialize` per call.
- Deviation from the extractor: ReportProcessor uses `Dispatch` only and sets `Visible = True` for refresh (report_processor.py:218, 225), while this file tries `GetActiveObject` first (124). These are two different Excel-acquisition patterns in the same project.

---

### core/file_organizer.py (426 lines) - [CORE]

- **Purpose:** for each selected client, creates the output folder tree under the target folder, **copies** (never moves) the client's source files into it under cleaned names, and writes a per-client text report.

#### Imports
- Standard library: `logging` (6), `Path` (7), typing Dict/List/Optional/Any/Tuple/Union (8; Tuple **unused**), `datetime` (9). `import shutil` sits inside `cleanup_on_error` (391).
- Project: `FOLDER_STRUCTURE`, `PROCESSING_MODES` (11). Helpers (12-16): `ensure_path_exists`, `safe_copy_file`, `create_backup`, `get_timestamp`, `sanitize_filename`, `safe_path_join`, `get_state_code`, `create_client_state_key` (comment `# ADD THESE` at 15 is a leftover).
- No OS-specific modules.

#### Module-level constants and inline values
- `logger` (18).
- Inline values:
  - default `client_name_max_length = 35` (30, 42);
  - the type-to-folder map (305-312);
  - report name prefix `"_Organization_Report_"` (330);
  - run-folder prefix `"Annual Statement-"` used in rerun/resume (179);
  - report separators `"=" * 50`.

#### The folder tree it builds [RAN exp4]
```
<target>/
  Annual Statement-<DDMMYY HHMM>/              (Level 1; FOLDER_STRUCTURE['level1'])
    <client-state key>/                         (Level 2; create_client_state_key, e.g. "ClientA Pvt Ltd-DL")
      Version-<DDMMYY HHMM>/                    (Level 3; FOLDER_STRUCTURE['version'])
        AnnualReport-…xlsx                      (Annual Report files go here)
        _Organization_Report_<client>_<state>.txt
        (ITC_Report_… / Sales_Report_… are written here by processing_handler)
        GSTR3B[ (<cleaned client>)]/            (Level 4; GSTR-3B Export files)
        ITC[ (<cleaned client>)]/               (GSTR-2B Reco + IMS Reco files)
        Sales[ (<cleaned client>)]/             (Sales + Sales Reco files)
  GST_Processing_Summary_<DDMMYY_HHMM>.xlsx     (written by processing_handler via ExcelHandler)
```

#### Class `FileOrganizer` (21-426)

**`__init__(self, target_folder, processing_mode='fresh', include_client_name=False, client_folder_settings=None, client_name_max_length=35)` (26-58)**
1. Stores the target as a Path (38), the mode (39) and the include flag (40). `client_folder_settings or {}` (41): no mutable-default bug.
2. `client_name_max_length if > 0 else 35` (42). A text value would raise TypeError [RAN]. The GUI passes an IntVar integer (processing_handler.py:194; main_window.py:93), so this is fine in practice, but an empty Spinbox makes `IntVar.get()` raise on the worker thread (A-R3's file).
3. `self.timestamp = get_timestamp()` (43): **one** value "DDMMYY HHMM" (minute resolution), used for both the Level-1 and Version names.
4. `_level1_folder = None` (47). Tracking lists `created_folders`, `copied_files`, `errors` (50-52).
5. `_validate_inputs()` (55). INFO log (57-58).
- The docstring (31-37) documents only two of the five parameters.
- Caller: processing_handler.py:189-195, on the worker thread, with a **new instance per run**.

**`_validate_inputs(self)` (60-69)**: the mode must be a key of `PROCESSING_MODES` ('fresh', 'rerun', 'resume'), else ValueError "Invalid processing mode: …" (62-63) [RAN]. It then **creates the target folder** if it is missing (67); any failure is re-raised as ValueError("Cannot create target folder: …") (68-69). The GUI only checks that the target text is non-empty before processing (gui/handlers/file_handler.py:138-140), so a mistyped target is created silently (MEDIUM).

**`create_client_structure(self, client_info) -> Dict[str, Path]` (71-148)**
1. `state_code = get_state_code(client_info['state'])` (82).
2. `client_state_key = create_client_state_key(raw client, raw state, max_length=self.client_name_max_length)` (83-87).
3. `client = sanitize_filename(raw client)` (90); `state = state_code` (91).
4. `level1 = self._get_level1_folder()` (94).
5. Level 2 = `level1 / client_state_key`, created, recorded (97-100). **Not passed through `sanitize_filename`.**
6. Version = `level2 / "Version-<self.timestamp>"`, created, recorded (103-108).
7. For each category (117-139): name = template.format(client=cleaned client) (119). Lookup key `f"{raw client}-{state_code}"` (122). If the global option is on, include (125-127); else `client_folder_settings.get(key, False)` (128-130). If not included, the `" (<client>)"` suffix is stripped (133-134). The folder is created and recorded (136-139).
8. Returns `{'level1', 'level2', 'version', 'gstr3b', 'itc', 'sales'}` (111-115, 138, 142).
9. Exception: ERROR log, appended to `self.errors`, **re-raised** (144-148). The caller's per-client `except` logs "❌ Failed <key>".
- Observations [RAN exp4 unless noted]:
  - The **Level-2 name** is abbreviated (Private→Pvt, Limited→Ltd, as whole words) and truncated to `max_length`. **Underscore-joined names are not abbreviated here** (`ClientA_Private_Limited-DL`), unlike `sanitize_filename`. Commit 4219ffb fixed only `sanitize_filename`.
  - The **Level-4 suffix** uses the full cleaned name, with underscores and up to 190 characters, e.g. `ITC (ClientA_Pvt_Ltd)`. The Setup help text (gui/tabs/setup_tab.py:112-117) shows "'ITC (ABC Pvt Ltd)'" with spaces. The "Max client folder name length" setting **does not shorten the Level-4 suffix**: with max 15, Level 2 became `ClientA Pvt -DL` but Level 4 stayed `ITC (ClientA_Pvt_Ltd)`.
  - **Two different scanned "clients" can share one Level-2 folder and one Version folder**, mixing their files:
    - abbreviation variants: `ClientC Private Limited` and `ClientC Pvt Ltd` both → `ClientC Pvt Ltd-DL`. Both clients' GSTR-3B files ended up in the same GSTR3B folder, with two organisation reports side by side;
    - names that agree in their first (max_length − code length − 1) characters (`'A'*32+'X'` and `'A'*32+'Y'`);
    - names that differ only in letter case, on Windows' case-insensitive disks (READ, MEDIUM);
    - states sharing a fallback code, for example 'Daman and Diu' and 'Dadra and Nagar Haveli' both → `DA`, and an unknown 'Unknown Place' → `UP`, the same as Uttar Pradesh (RAN exp3). **[GST RULE]** (state-code logic).
  - With max length ≤ 8 (two-letter code) or ≤ 9 (three-letter code), **the state code is dropped completely** (max 6 gives `Client`). The Spinbox minimum is 15 (setup_tab.py:440), but a typed value is never range-checked: only "≤ 0 → 35" (setup_tab.py:435-436, main_window.py:386-387) (MEDIUM).
  - `created_folders` gathers duplicates across clients.
- Caller: processing_handler.py:253.

**`_get_level1_folder(self)` (150-161)**: computes once and caches in `self._level1_folder` (275d9a4, 23-07-2026). Caller: 94.

**`_compute_level1_folder(self)` (163-196)**
- `fresh`: `target / "Annual Statement-<timestamp>"`, created with exist_ok, so a second run **in the same minute reuses the same folder** (165-173) [RAN].
- `rerun`/`resume`:
  - lists the target (`iterdir`) and keeps folders whose name starts with "Annual Statement-" (177-180);
  - picks the one with the **latest mtime**, not the date in the name (184). A folder's mtime changes when an entry directly inside it is added, removed or renamed. Any folder with that prefix counts, including renamed copies;
  - if none is found: WARNING "No existing Annual Statement folder found, creating new" and it creates one (188-196).
- Listing or stat errors (for example a network drop) propagate. That client fails, the cache stays empty, and the next client retries.

**`organize_files(self, client_info, folders, progress_callback=None) -> List[Dict]` (198-248)**
1. Flattens `(type, file)` pairs (216-219). The order follows the types as first met in the scan, then the files.
2. For each pair:
   - `progress_callback(idx+1, total, "Copying <name>")` (225-227). The callback runs on the worker thread;
   - `_organize_single_file`, with the result appended to `results` and `self.copied_files` (229-231);
   - if the status is Success: the matching dict in `client_info['files'][type]` (found by `full_path`, which is redundant because it is the same object) gets `name` and `sanitized_name` set to the copy's name (233-246).
- This **mutates** the app's own `client_data`. These are the same dicts as `app.client_data` and `app.scanned_files`, so the original names are gone from memory for the rest of the session. The DEBUG message at 245 prints the new name twice, because `file_data` is the same object.
- Returns the list of results.
- Caller: processing_handler.py:271.

**`_organize_single_file(self, file_type, file_data, folders) -> Dict` (250-300)**
1. Result `{filename: original name, file_type, source: full_path, destination: '', status: 'Failed', error: None}` (254-261).
2. Destination folder by type (265). An unknown type gives error "No destination folder for type: …" and returns (266-268).
3. `safe_filename = sanitize_filename(name)` (271); `dest_path` (272); the result records `destination` and `sanitized_filename` (273-274).
4. **If the destination already exists** (277-287):
   - `fresh`: `create_backup(dest_path)` copies the *existing destination file* to `<stem>_backup_<DDMMYY_HHMM><ext>` **in the same folder**, then the copy overwrites it;
   - `resume`: status 'Skipped', error 'File already exists', no copy;
   - `rerun`: no branch, so it is **overwritten without a backup**.
5. `safe_copy_file(source, dest, verify=True)` (290): Success, or error 'Copy operation failed' (294).
6. `except Exception`: error text in the result, logged (296-298).
- What happens to the **original files**: they are only read. `shutil.copy2` at helpers.py:199 copies content and timestamps. **Never moved, renamed or deleted** (grep: no unlink/rename/move on source paths) [READ HIGH].
- What happens at the destination: overwritten (rerun), backed up then overwritten (fresh), skipped (resume). A copy whose size does not match is deleted (helpers.py:205-208).
- [RAN] Two source files that clean to the same name (`GSTR3B-ClientB-Delhi-Apr.xlsx` and `GSTR3B-ClientB-Delhi-Apr!.xlsx`):
  - in **fresh** mode the first copy was backed up as `GSTR3B-ClientB-Delhi-Apr_backup_<now>.xlsx` **inside the GSTR3B folder** and then overwritten by the second;
  - in **resume** mode the second was "Skipped – File already exists", i.e. silently not copied.
- The backup timestamp is *now* (`get_safe_timestamp`), not the run timestamp. A second backup in the same minute **overwrites the first** [RAN].
- If `create_backup` fails it returns None and the code **goes on to overwrite without a backup** (280-282) [READ HIGH].
- Caller: `organize_files` (229).

**`_get_destination_folder(self, file_type, folders)` (302-315)**: `{'GSTR-3B Export': 'gstr3b', 'GSTR-2B Reco': 'itc', 'IMS Reco': 'itc', 'Sales': 'sales', 'Sales Reco': 'sales', 'Annual Report': 'version'}`, then `folders.get(key)`. **[GST RULE]**, tagged conservatively: it decides which folder each return file is placed in, and the templates read those folders. This disagrees with the unused `FILE_PATTERNS[...]['folder']` value for Annual Report, which is `'root'` (constants.py:58). Caller: 265.

**`create_organization_report(self, folders, client_info) -> Optional[Path]` (317-379)**
- File `"_Organization_Report_<raw client>_<raw state>.txt"` in the Version folder (330-331).
- Content: header "GST FILE ORGANIZATION REPORT", "Generated: %Y-%m-%d %H:%M:%S", Client, State, Processing Mode.
- "FOLDER STRUCTURE:" lists every folder except level1/level2, relative to Level 1 (346-350).
- "FILES ORGANIZED:" lists each type as "<type> (<n> files):", then the names (cleaned after copying) (355-360).
- "STATISTICS:" (363-368):
  - "Total Files" = `client_info['file_count']` for *this client*;
  - "**Successfully Copied**" and "**Folders Created**" are counted over the **whole run so far** (`self.copied_files`, `self.created_folders`), so they are **wrong for every client after the first**. [RAN: the third client's report said "Successfully Copied: 10, Folders Created: 16" and the fourth "11 / 21", for one file each.]
- Written as UTF-8 (371). Errors return None (377-379). The report is written even if copies failed.
- Caller: processing_handler.py:414, which ignores the return value.

**`cleanup_on_error(self, folders)` (381-396)**: **NO CALLERS FOUND (dead)**. In fresh mode it would remove the Version folder with `rmtree(..., ignore_errors=True)` when `_is_folder_empty` is True, and that would delete the `_Organization_Report` too [RAN].

**`_is_folder_empty(self, folder)` (398-407)**: True if no file anywhere below has a name not starting with "_". Errors return False. Only caller: `cleanup_on_error` (390), so it is dead too.

**`get_summary(self)` (409-426)**: **NO CALLERS FOUND (dead)**. It returns counts. `folders_created` uses `len(set(...))` here but not in `create_organization_report` (368), which is inconsistent.

#### Business rules in this file
- **[GST RULE] (conservative)** Type-to-folder placement (305-312).
- The folder naming scheme: Level 1 / Level 2 / Version / category (103-105, 117-139, 167-169; templates in constants.py:81-96). The time stamp is `DDMMYY HHMM` (helpers.py:27). This is a run timestamp, not a financial year.
- Processing modes (165-196, 277-287): fresh = new Level-1 folder; rerun = the latest Level-1 folder by mtime plus a new Version folder, overwriting on a name clash; resume = the same as rerun, but skipping existing files.
- **Resume does not do what its labels say.** The Version folder is always new (103-105, minute timestamp), so "skip files already in place" only works if the resumed run starts **in the same minute** as the interrupted one. Otherwise everything is copied again into a new Version folder [RAN exp4 RUN 4]. Labels: "Continue previously interrupted processing" (constants.py:187), and setup_tab.py:108-109 "Continue a run that was interrupted, skipping files that are already in place." [HIGH]
- Name cleaning (`sanitize_filename`) is applied to every copied file name (271) and to the Level-4 suffix (90). The abbreviation-and-truncation key (`create_client_state_key`) applies to Level 2 (83-87).
- The global "include client name" option overrides the per-client setting (124-130).

#### Edge cases explicitly handled
- Invalid mode: 62-63.
- Target creation failure: 66-69.
- Unknown file type: 266-268.
- Existing destination: 277-287.
- Copy failure: 290-294.
- Folder-creation failure: 144-148.
- No existing run folder in rerun/resume: 188-196.
- Report-writing failure: 377-379.

#### Gaps, likely failures and silent bugs
1. Resume mode (above) [HIGH].
2. **Resume mode feeds the wrong file name to the report** when a file is skipped, because `name` is not updated (233-246; see excel_handler `prepare_template_data`) [RAN exp8].
3. Rerun in the same minute overwrites files without a backup (283-287). Fresh mode puts **backup `.xlsx` files inside the category folders** (helpers.py:226). If a template reads a whole folder (GSTR-3B is linked by folder only), those backups and any "(1)" duplicates would be read too (MEDIUM; template-dependent).
4. Level-2 collisions mix two clients' files (see `create_client_structure`) [RAN].
5. Cumulative statistics in the per-client report (363-368) [RAN].
6. **No path-length guard**. Commit 7347383 says so ("a path-length guard remains the durable fix"), and nothing in this file measures paths [READ HIGH]. Fixed name lengths: "Annual Statement-DDMMYY HHMM" = 28 characters, "Version-DDMMYY HHMM" = 19 characters.
7. Network drop: listing errors in `_compute_level1_folder` fail that client. A partly copied file can remain at the destination (helpers.py:213-215).
8. Shared state without locks: `organize_files` mutates dicts the UI thread also reads (dry run, tree, export). The risk is low.
9. `client_key` is recomputed on each loop pass (122) with the same value, which is harmless. No late-binding, mutable-default or off-by-one problems found.
10. Two copies of the app at once in fresh mode within the same minute would share the same Level-1 and Version folders, with backups and overwrites interleaving (LOW likelihood).

#### Excel touchpoints
None open a workbook. Excel files are copied as plain files (`safe_copy_file`, `create_backup`).

#### Files and folders written
- The folders above.
- Copied files at `<category or Version>/<sanitize_filename(original name)>`.
- Backups `<stem>_backup_<DDMMYY_HHMM><ext>` next to the replaced file.
- `_Organization_Report_<client>_<state>.txt` (UTF-8 text) in the Version folder.
- Reads: `target.iterdir()` and folder mtimes (178, 184).

#### Comments and docstrings
- Module docstring "Handles folder creation and file organization." (1-4).
- 45-46 and 150-158: the caching notes match the code.
- 203: "FIXED: Updates client_info with sanitized filenames after copying." This is true only for Success.
- 311: "# Goes in version folder" is true.
- 383-384: "Only cleanup for fresh runs" belongs to dead code.

#### Pattern notes
- Errors: create-structure re-raises; file operations record the error in the result dict; reports return None.
- Logging: the module logger, at INFO for every copied file (292, client file names into the log).
- No threading of its own; it is used from the worker thread.

---

### core/file_parser.py (371 lines) - [CORE]

- **Purpose:** lists the valid Excel files in **one** folder (subfolders are not scanned) and matches each name against six name patterns. Recognised files are grouped by "client + state code", with missing and duplicate types and a status worked out per client. Valid Excel files with unrecognised names become "variations".

#### Imports
- Standard library: `logging` (7), `Path` (8), typing Dict/List/Tuple/Optional/Any/Union (9; Optional **unused**), `defaultdict` (10).
- Project: `FILE_PATTERNS`, `EXPECTED_FILE_TYPES` (12); helpers `get_file_info`, `scan_excel_files`, `validate_client_name`, `format_size`, `get_state_code` (13-16).
- No OS-specific modules.

#### Module-level
`logger` (18).

#### File-name recognition rules (exact; tried in this order, first match wins at 128 `break`)
The order is the dict order in constants.py:13-62. All six patterns are **case-insensitive**. All accept the extension **`.xlsx` or `.xls` only** (`\.xlsx?$`). Every captured group is trimmed of surrounding spaces (102, 104, 106). The client and state groups are `[^-]+`, so **neither can contain a hyphen**. The six literal prefixes do not overlap, so in practice the order does not matter (HIGH). **[GST RULE]** for all six, because they decide which return or report a file is. TIME-VARYING: the return names and file formats.

| # | Key | Regex | type / category / folder | Groups |
|---|---|---|---|---|
| 1 | `GSTR-2B-Reco` | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | 'GSTR-2B Reco' / 'ITC' / 'itc' | client, state, period (may contain hyphens) |
| 2 | `ImsReco` | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` | 'IMS Reco' / 'ITC' / 'itc' | client, state, date = exactly 8 digits; optional "(N)" or " (N)" after it (5c8ce95) |
| 3 | `GSTR3B` | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | 'GSTR-3B Export' / 'GSTR3B' / 'gstr3b' | client, state, month (no hyphen); optional leading "(N)" or "(N) " |
| 4 | `Sales` | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | 'Sales' / 'Sales' / 'sales' | client, state, start_month, end_month: exactly 4 parts |
| 5 | `SalesReco` | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | 'Sales Reco' / 'Sales' / 'sales' | client, state, period (may contain hyphens) |
| 6 | `AnnualReport` | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | 'Annual Report' / 'Annual' / 'root' | client, state, year (may contain hyphens) |

[RAN exp1] behaviour:
- Recognised: the canonical forms; `ImsReco-…-01012024 (1).xlsx` and `…01012024(2).xlsx` (date still "01012024"); `(1) GSTR3B-…` and `(1)GSTR3B-…`; `GSTR3B-…-Apr (1).xlsx` (month = "Apr (1)"); `.xls`; an upper-case `.XLSX`; all-lower-case names; spaces around parts (trimmed).
- **Not** recognised (they go to "variations"):
  - `ImsReco-…-0101202.xlsx` (7 digits) and `…-Jan2024.xlsx`;
  - `GSTR3B-…-Apr-2024.xlsx` (an extra hyphen part); `GSTR-3B-…` (hyphenated keyword);
  - `Sales-…-Apr2024.xlsx` (one period part) and `Sales-…-Apr-Mar-2025.xlsx` (three);
  - any `.xlsm` (the scanner lists `.xlsm`, but no pattern accepts it);
  - `Annual Report-…` (with a space);
  - `(1) Sales-…` and `Copy of Sales-…` (only GSTR3B accepts a prefix).
- **Silent mis-split of hyphenated client names** for patterns 1, 5 and 6: `GSTR-2B-Reco-Client-East-Delhi-Apr-Mar.xlsx` gives client "Client", state "East", period "Delhi-Apr-Mar". The file is **accepted under the wrong client and a made-up state** (state code fallback "EAS"). Patterns 3 and 4 reject the same names. So one client's files can end up split between a wrong "client" and the variations list [HIGH].

#### Class `FileParser` (21-371)

**`__init__` (26-30)**: `patterns = FILE_PATTERNS`, `expected_types = EXPECTED_FILE_TYPES`, `progress_callback = None`, then `reset()`. Callers: gui/main_window.py:115 (one instance per session); tests.

**`reset` (32-46)**: `scanned_files = {}`; `client_data = defaultdict(<factory>)`. The factory makes a fresh record each time (no shared-state bug): `{'client': '', 'state': '', 'files': defaultdict(list), 'missing_files': [], 'extra_files': [], 'status': 'Unknown', 'total_size': 0, 'file_count': 0}`. Also `variations = []`, `errors = []`. Callers: 30, 148.

**`_update_progress(current, total, message)` (48-54)**: calls the callback with three arguments if one is set. Exceptions give a WARNING "Progress callback failed: …" and are swallowed (added in 6237597). Caller: 166.

**`parse_filename(self, filename) -> Dict` (56-134)**
1. Result `{filename, parsed: False, type: 'Unknown', category: 'Unknown', folder: 'Unknown', client: '', state: '', metadata: {}}` (66-75).
2. DEBUG if the name contains "sales" (78-79).
3. Tries each pattern (82-128). On a match it sets parsed, pattern, type, category and folder (89-95), then maps the groups: client and state are trimmed into their fields, others into `metadata` (98-106). Empty groups are skipped (100).
4. If both client and state are non-empty, `validate_client_name` runs on **both** (110-111). Failures are added to `result['warnings']` as "Client name issue: …" / "State name issue: …". The state messages still use the client wording, e.g. "State name issue: Client name too short" [RAN]. **`warnings` is never read anywhere** (grep), so it is dead information that never reaches the user.
5. DEBUG logs (121, 124-126); `break`.
6. No match: **WARNING "Could not parse filename: <name>"** (131) puts client file names into the log; `error = 'No matching pattern found'` (132).
- Returns the dict.
- **`metadata` (period/date/month/start_month/end_month/year) is never used by any app code** (grep: only tests/test_file_patterns.py:67). **[GST RULE], an absence:** **no financial-year or tax-period check exists anywhere** in these files. Files of different periods or years for the same client and state are grouped together and treated as duplicates or as monthly GSTR-3B files.
- Callers: 196; tests/test_file_patterns.py:35, 58, 65, 72, 91.

**`scan_folder(self, folder_path, progress_callback=None) -> (scanned_files, client_data, variations)` (136-177)**
1. `reset()`, store the callback (148-149).
2. The folder must exist, else FileNotFoundError "Folder not found: …" (152-153), and must be a folder, else ValueError "Not a directory: …" (154-155) [RAN].
3. `excel_files = scan_excel_files(folder)` (158).
4. If none: WARNING, return empty results (159-161).
5. INFO count (163). For each file: progress (idx+1, total, "Processing <name>") (166), then `_process_file(path, folder, stat)` (167).
6. `_analyze_client_completeness()` (169). INFO summary (171-172).
7. Returns `self.scanned_files` (the live dict, **keyed by file name only**), `dict(self.client_data)` (a shallow copy whose inner records are the same objects), and `self.variations` (the live list) (173).
8. Any exception: ERROR, appended to `errors`, re-raised (174-177).
- Caller: gui/handlers/file_handler.py:72, **on the Tk UI thread**. The "🔍 Scan Files" button (setup_tab.py:459-460) calls `main_window.scan_files` (222-223), which calls `file_handler.scan_files`. The window cannot redraw during the scan; the status-bar label set by the callback is only drawn when the scan ends [READ HIGH]. Commit 4f84ba9 measured about 41 s for 2508 files on the office share.
- (Calls to `scan_folder` at processing_handler.py:169 and in power_query_extractor/... are to `ExtractorPanel.scan_folder`, a different method.)
- **Silent drops** [RAN exp2]: these never appear in `scanned_files` or in `variations`, with no message to the user:
  - files whose extension is not `.xlsx`/`.xls`/`.xlsm`, so `.xltx` and `.xlsb` are ignored;
  - files under 1024 bytes;
  - files whose first bytes are not an Excel signature (a wrong-signature file is not even logged);
  - unreadable files (an ERROR log only);
  - files in subfolders;
  - Excel lock files `~$…` (165 bytes);
  - a folder named `X.xlsx`.
- A **network failure while listing** makes `scan_excel_files` return `[]` (helpers.py:363-365), and the GUI says "No valid GST files found", which looks the same as an empty folder [READ HIGH].

**`_process_file(self, file_path, base_folder, file_stat=None)` (179-224)**
1. `file_info = get_file_info(file_path, file_stat)` (188) reuses the directory-listing metadata.
2. No re-validation (190-193; a603c02 removed it).
3. `parsed = parse_filename(name)` (196).
4. `file_data = {**file_info, **parsed, 'full_path', 'relative_path'}` (199-204). [RAN] keys: category, client, created, extension, filename, folder, full_path, is_dir, is_file, metadata, modified, name, parsed, path, pattern, relative_path, size, size_formatted, state, type.
5. `scanned_files[name] = file_data` (207).
6. If parsed and client and state: `_add_to_client_data` (210-211). Else a variations entry `{filename, path, issue: parsed.get('error', 'Could not parse filename'), size, pattern_expected: _suggest_pattern(name)}` (212-220). An all-space client or state comes out as '' after trimming, so it becomes a variation with issue "Could not parse filename".
7. Exceptions: ERROR, appended to `self.errors` (222-224).
- **`self.errors` is never read outside the class** (grep), so per-file errors never reach the GUI.
- Caller: 167.

**`_add_to_client_data(self, file_data)` (226-250)**
- Key = `f"{raw client}-{get_state_code(state)}"` (229-230). **[GST RULE]**, state-code logic.
- `client` and `state` are overwritten by each file (234-235), so **the last file's state spelling wins**: 'Orissa' and 'Odisha' merge into one client shown as 'Odisha' [RAN].
- `state_code` is stored (236). The file is appended to `files[type]` (239-240). `total_size` and `file_count` are updated (243-244). DEBUG logs (246, 249-250).
- **Case and abbreviation variants of the same client are separate clients**: 'ClientC Private Limited', 'ClientC Pvt Ltd' and 'clientc pvt ltd' gave three clients [RAN].
- Caller: 211.

**`_analyze_client_completeness(self)` (252-286)**
- For each client: `missing = expected − found`, sorted (255-260).
- `extras`: every type with more than one file, **except 'GSTR-3B Export'**, which is always allowed to repeat; each gives "Multiple <type> (<n> files)" (263-271). **[GST RULE]**: GSTR-3B is treated as possibly many files with no upper limit (comment "monthly filings - unlimited", 266). TIME-VARYING (filing frequency).
- `completeness = found/expected*100`, rounded to 1 decimal (274-275).
- Status (278-286): 'Complete' ✅ if nothing is missing and there are no extras; otherwise 'Missing N files' ⚠️ (**missing takes priority over duplicates**); otherwise 'Has duplicates' ⚠️.
- Caller: 169.

**`_suggest_pattern(self, filename) -> str` (288-306)**: keyword checks on the lower-cased name, in this order:

| Keywords found | Suggestion returned |
|---|---|
| 'gstr' and '2b' | "GSTR-2B-Reco-ClientName-State-Period.xlsx" |
| 'ims' | "ImsReco-ClientName-State-DDMMYYYY.xlsx" |
| 'gstr' and '3b' | "GSTR3B-ClientName-State-Month.xlsx" |
| 'sales' and 'reco' | "SalesReco-ClientName-State-Period.xlsx" |
| 'sales' | "Sales-ClientName-State-StartMonth-EndMonth.xlsx" |
| 'annual' | "AnnualReport-ClientName-State-Year.xlsx" |
| none of these | "Check file naming convention" |

These are plain substring checks, so any name containing "ims" (for example "claims") gets the IMS suggestion. Only one suggestion is ever produced (elif), so the `" or ".join` is pointless. Caller: 219.

**`get_statistics(self)` (308-336)**: returns total_files, parsed_files, unparsed_files, total_clients, complete_clients, incomplete_clients, variations, total_size (of all listed files), total_size_formatted, file_type_distribution, parsing_rate %, completion_rate %. Caller: gui/main_window.py:299 (`update_summary`).

**`validate_results(self)` (338-371)**: **NO CALLERS FOUND (dead)**. It would flag empty client or state names, and "Missing Critical Files" for `critical_types = ['Annual Report', 'GSTR-3B Export']` (359). **[GST RULE]** (dead).

#### Business rules in this file
- The six recognition patterns, via FILE_PATTERNS [GST RULE, TIME-VARYING].
- Client key = raw client + state code (230) [GST RULE].
- The expected six types and the completeness percentage (256-275) [GST RULE].
- GSTR-3B may repeat without limit (266-268) [GST RULE, TIME-VARYING].
- Status priority: missing over duplicates (278-286).
- Client and state validation rules, from helpers: 2-100 characters, no `<>:"/\|?*` (110-119).
- Critical types (359, dead) [GST RULE].
- No financial-year or period check (absence) [GST RULE].

#### Edge cases handled
- Folder missing or not a folder: 152-155.
- No Excel files: 159-161.
- Progress-callback failures: 50-54.
- Per-file errors: 222-224.
- Empty groups: 100.
- Empty client or state after trimming: 210.

#### Gaps and silent bugs
1. The scan runs on the UI thread [HIGH].
2. Silent drops of invalid, unreadable and subfolder files [RAN].
3. Hyphenated client names are mis-split [RAN].
4. `.xlsm` is listed but can never match [RAN].
5. The "(1)" duplicate sorts first and becomes the linked file [RAN].
6. Case and abbreviation variants become separate clients [RAN].
7. The last state spelling wins [RAN].
8. Validation warnings are never shown [READ HIGH].
9. Period and financial year are ignored [READ HIGH].
10. `scanned_files` is keyed by name only. That is unique because the scan is one folder deep; it would break if the scan were ever made recursive.
11. A client or state that looks numeric causes trouble later in the GUI (see Cross-file observations) [RAN exp7].

#### Excel touchpoints
None directly. The 8-byte signature reads happen in helpers.

#### Files read
- The source folder listing and metadata.
- The first 8 bytes of each candidate file (through helpers).
- Nothing written. Logs only.

#### Comments and docstrings
- "File Parser for GST File Organizer v3.0 / Robust pattern matching and file categorization. / Fixed with proper Union import." (1-5). The last line is a leftover change note.
- 190-193: the "no re-validation" note matches the code.

#### Pattern notes
Errors are logged and collected in `self.errors` but never shown; only folder-level errors are raised. The callback has three arguments, while `ProgressTracker` uses four (helpers).

---

### utils/__init__.py (38 lines) - [UTIL]

- **Purpose:** marks `utils` as a package and re-exports a chosen subset of `constants` and `helpers`.
- **Imports:**
  - from `.constants` (6-10): `FILE_PATTERNS`, `EXPECTED_FILE_TYPES`, `FOLDER_STRUCTURE`, `EXCEL_TEMPLATE_MAPPING`, `GUI_CONFIG`, `PROCESSING_MODES`, `MESSAGES`, `EXCEL_EXTENSIONS`, `EXCEL_SAFETY`.
  - from `.helpers` (12-20): `get_timestamp`, `get_date_only`, `get_safe_timestamp`, `ensure_path_exists`, `safe_path_join`, `sanitize_filename`, `get_file_info`, `safe_copy_file`, `create_backup`, `validate_excel_file`, `find_excel_files`, `scan_excel_files` (added in 56aaad9), `format_size`, `truncate_path`, `clean_windows_path`, `extract_filename_without_extension`, `validate_client_name`, `get_unique_path`, `ProgressTracker`, `format_duration`, `calculate_file_hash`, `get_relative_excel_path`.
  - No OS-specific modules.
- **Module-level:** `__all__` (22-36), the same list. `__version__ = '3.0.0'` (38).
- **Not re-exported:**
  - constants: `WORKFLOW_MODES`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `STATE_CODE_MAPPING`;
  - helpers: `get_short_path`, `get_state_code`, `create_client_state_key`, `_has_excel_signature`, `_passes_metadata_checks`.
  - This is harmless, because every caller imports from the submodules directly.
- **Who uses the re-exports:** nobody. The grep found no `from utils import <name>`, only `from utils import dependency_installer` in tests/test_dependency_installer.py:105, which is a submodule import [READ HIGH].
- **Side effect:** the file runs on **every** import of any `utils.*` module, so `constants` and `helpers` are always imported. Both use only the standard library, so the start-up dependency check (main.py:68 `from utils.requirements_check import …`) still works when openpyxl is missing [READ HIGH].
- **Business rules / edge cases / Excel / files:** none.
- **Gaps:** the hand-kept lists must follow any rename. Re-exporting `EXCEL_EXTENSIONS` and `EXCEL_SAFETY` gives an illusion that they are used; they are dead.
- **Comments:** "Utils package for GST File Organizer v3.0 / Contains utility functions and constants." (1-4).

---

### utils/constants.py (292 lines) - [CONFIG]

- **Purpose:** a single module of app-wide dictionaries: file-name patterns, expected types, folder-name templates, the Links-sheet mapping, GUI look, mode labels, messages, Excel "safety" values and the state-name→code table.
- **Imports:** `re` (6); `typing` Dict, List, Tuple (7; Tuple **unused**). No OS-specific modules.
- **Type-hint slip:** `Dict[str, Dict[str, any]]` (13) uses the built-in function `any` instead of `typing.Any`. It is harmless at run time.

#### Every constant (line → value → what it controls → who reads it)

**`FILE_PATTERNS` (13-62)**: six entries, each with keys `pattern`, `type`, `category`, `folder`, `description` and `groups`. The full regex table is under core/file_parser.py above. Descriptions (shown on the Setup tab's "📝 EXPECTED FILE NAMES" card, gui/tabs/setup_tab.py:494-503, with `type`):

| Type | Description shown |
|---|---|
| GSTR-2B Reco | "GSTR-2B-Reco-ClientName-State-Period.xlsx" |
| IMS Reco | "ImsReco-ClientName-State-DDMMYYYY.xlsx" |
| GSTR-3B Export | "GSTR3B-ClientName-State-Month.xlsx" |
| Sales | "Sales-ClientName-State-StartMonth-EndMonth.xlsx" |
| Sales Reco | "SalesReco-ClientName-State-Period.xlsx" |
| Annual Report | "AnnualReport-ClientName-State-Year.xlsx" |

- The card does not mention that `.xls`, the "(N)" prefix (GSTR3B) and the "(N)" suffix (IMS) are also accepted.
- `category` and `folder` are copied into each parse result (file_parser.py:93-94), but **nothing reads them afterwards** (grep), so they are dead. `folder` disagrees with the organiser for Annual Report ('root' here, the Version folder in file_organizer.py:311).
- Readers: core/file_parser.py:27; gui/tabs/setup_tab.py:494. **[GST RULE]**, TIME-VARYING.
- History: edac6a3; IMS "(N)" suffix added in 5c8ce95 (23-07-2026).

**`EXPECTED_FILE_TYPES` (68-75)**: `['GSTR-2B Reco', 'IMS Reco', 'GSTR-3B Export', 'Sales', 'Sales Reco', 'Annual Report']`. These are the six document types every client is expected to have. They drive "missing", "completeness" and "Complete". Readers: core/file_parser.py:28; tests/test_scan.py:36. **[GST RULE]**, TIME-VARYING.
- The same six type strings are repeated as **literal text** in file_parser.py:267 and 359, file_organizer.py:306-311 and excel_handler.py:625, 636, 647, 661, 677 and 693. They are consistent today (READ HIGH), but a rename would have to be made in about 15 places.

**`FOLDER_STRUCTURE` (81-96)**
- `'level1': 'Annual Statement-{timestamp}'` (82): used at file_organizer.py:167 and 190.
- `'level2': '{client}-{state}'` (83): **NOT used**. Level 2 is built by `create_client_state_key` instead, so this is dead.
- `'version': 'Version-{timestamp}'` (84): used at file_organizer.py:103.
- `'categories': {'gstr3b': 'GSTR3B ({client})', 'itc': 'ITC ({client})', 'sales': 'Sales ({client})'}` (91-95): used at file_organizer.py:117. Shortened in 7347383 (26-07-2026) from 'GSTR-3B Exports …', 'Other ITC related files …' and 'Sales related files …'.
- The comment at 85-90 records the measurement: worst path 235 characters with 27 files over the limit, against 215 with none. It agrees with the commit message.
- The Level-1 and Version timestamp is DDMMYY HHMM (helpers.py:27). Day-first text does not sort in date order, which is why the extractor picks folders by mtime (extractor_window.py:643-648).

**`EXCEL_TEMPLATE_MAPPING` (102-126)** **[GST RULE]** (which Links cell receives which source; TIME-VARYING with the templates):

| Report | Sheet | Cell → data key |
|---|---|---|
| ITC | 'Links' | B2 → gstr3b_folder, B4 → annual_folder, B5 → annual_filename, B7 → gstr2b_folder, B8 → gstr2b_filename, B10 → ims_folder, B11 → ims_filename |
| Sales | 'Links' | B2 → sales_folder, B3 → sales_filename, B5 → annual_folder, B6 → annual_filename, B8 → sales_reco_folder, B9 → sales_reco_filename |

- Not written: ITC B3, B6, B9; Sales B4, B7. Presumably these hold labels in the templates (LOW; the templates are not in the repository).
- Readers: core/excel_handler.py:162, 285.

**`GUI_CONFIG` (132-157)**
- `window`: title 'GST File Organizer v3.0', size '1200x800', min_size (1000, 700). Used at main_window.py:59-61.
- `fonts` (138-144): default ('Segoe UI', 10), title ('Segoe UI', 16, 'bold'), heading ('Segoe UI', 12, 'bold'), small ('Segoe UI', 9), code ('Consolas', 9). **Never read**: the grep for `'fonts'` found only this definition, and the widgets hard-code their fonts. These are Windows font names.
- `colors` (145-156): primary '#0078D4', success '#107C10', warning '#FF8C00', danger '#D83B01', info '#0099BC', light '#F8F9FA', dark '#323130', white '#FFFFFF', complete '#E8F5E8', incomplete '#FFF4E5'. All are used across the GUI **except 'white'**. Use counts from the grep: primary 22, success 13, dark 10, light 8, info 5, danger 4, warning 3, complete 2, incomplete 2.
- Readers: main_window, client_handler, the validation, processing, extract and setup tabs, dark_mode_manager, status_bar, title_bar (full list in the grep output above).

**`WORKFLOW_MODES` (163-172)** (added in 510f7b4, 26-07-2026):
- 'organize': name "📁 Organise files & create reports", description "Sort the files and build the ITC/Sales reports, then stop."
- 'full': name "🔄 Full pipeline", description "Also refresh Power Query and extract the data (slower)."
- Readers: gui/tabs/setup_tab.py:327; tests/test_workflow_mode.py:71-72.
- It sits under the "PROCESSING MODES" banner comment (159-161) and is not re-exported in utils/__init__.py.

**`PROCESSING_MODES` (174-190)**:
- 'fresh': name "🆕 Fresh Run", description "Create new folder structure (recommended for first time)", icon '🆕'.
- 'rerun': "🔄 Re-Run", "Add new version folders to existing structure", '🔄'.
- 'resume': "▶️ Resume", "Continue previously interrupted processing", '▶️'.
- The `icon` key is **never read** (grep).
- Readers: core/file_organizer.py:62 (validation), gui/tabs/setup_tab.py:401 (radio buttons).
- The 'resume' description does not match the behaviour (see file_organizer).

**`EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm'}` (196)**: only re-exported (utils/__init__.py:9, 26). **Never used** (dead).

**`TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` (197)**: **never used**. The Setup check uses its own `TEMPLATE_SUFFIXES = {'.xlsx', '.xltx', '.xltm', '.xlsm'}` (gui/utils/setup_validation.py:21), and the browse dialog offers `"*.xltx *.xlsx"` (gui/handlers/file_handler.py:31, 42). That is three different lists.

**`MAX_FILE_SIZE_MB = 100` (198)**: **never used**. No maximum file-size check exists anywhere in these files.

**`EXCEL_SIGNATURES = {b'PK\x03\x04': 'xlsx', b'\xd0\xcf\x11\xe0': 'xls'}` (199-202)**: **never used**. It is duplicated as `_XLSX_SIG` / `_XLS_SIG` at helpers.py:264-265.

**`MESSAGES` (208-228)**:
- errors:
  - 'no_source': "Please select a source folder containing GST files" (file_handler.py:123);
  - 'no_templates': "Please select both ITC and Sales template files" (file_handler.py:135);
  - 'no_target': "Please select a target folder for organized files" (file_handler.py:139);
  - 'no_selection': "Please select at least one client to process" (processing_handler.py:45, 83);
  - 'invalid_excel': "Invalid Excel file: {file}" (**unused**);
  - 'template_error': "Error loading template: {error}" (**unused**);
  - 'processing_error': "Error processing {client}: {error}" (**unused**).
- success (**all unused**): 'scan_complete' "✅ Found {count} clients with {files} files"; 'client_complete' "✅ Processed {client} successfully"; 'all_complete' "🎉 All processing completed! {count}/{total} successful".
- info (**all unused**): 'scanning' "🔍 Scanning folder for GST files..."; 'processing' "⚙️ Processing {client}..."; 'creating_report' "📊 Creating {report} report...".
- gui/main_window.py:20 imports MESSAGES but never uses it.

**`EXCEL_SAFETY` (234-245)**: `max_retries 3`, `retry_delay 1` (seconds), `temp_prefix 'temp_gst_'`, `backup_prefix 'backup_'`, `verification_size_min 10240` (10 KB), `save_options {data_only: False, keep_vba: False, keep_links: False}`.
- Stored as `ExcelHandler.safety_config` (excel_handler.py:55) and **never read**, so all of it is dead.
- The real behaviour differs from these values:
  - there are no retries;
  - temp files carry the output name with no prefix;
  - backups use the infix `_backup_` (helpers.py:225);
  - verification compares sizes only, with no 10 KB minimum (the scan minimum is 1024 bytes);
  - openpyxl loads with `keep_links=True` (excel_handler.py:282).

**`STATE_CODE_MAPPING` (251-292)** **[GST RULE]** (state-code logic), TIME-VARYING (states and union territories have been reorganised in the past, and abbreviations differ between systems).
- 39 keys, all lower case, mapping to **two-letter alphabetic** abbreviations. There are no two-digit numeric codes.
- States and aliases:

| Name | Code |
|---|---|
| andhra pradesh | AP |
| arunachal pradesh | AR |
| assam | AS |
| bihar | BR |
| chhattisgarh | CG |
| goa | GA |
| gujarat | GJ |
| haryana | HR |
| himachal pradesh | HP |
| jharkhand | JH |
| karnataka | KA |
| kerala | KL |
| madhya pradesh | MP |
| maharashtra | MH |
| manipur | MN |
| meghalaya | ML |
| mizoram | MZ |
| nagaland | NL |
| odisha | OD |
| orissa (alias, "# Alternative name") | OD |
| punjab | PB |
| rajasthan | RJ |
| sikkim | SK |
| tamil nadu | TN |
| telangana | TS |
| tripura | TR |
| uttar pradesh | UP |
| uttarakhand | UK |
| west bengal | WB |

- Union territories and aliases:

| Name | Code |
|---|---|
| andaman and nicobar islands | AN |
| chandigarh | CH |
| dadra and nagar haveli and daman and diu | DD |
| delhi | DL |
| national capital territory of delhi | DL |
| jammu and kashmir | JK |
| ladakh | LA |
| lakshadweep | LD |
| puducherry | PY |
| pondicherry (alias, "# Alternative name") | PY |

- Only reader: `utils/helpers.get_state_code` (helpers.py:598, 604).
- Names not in the table go to the fallback in `get_state_code`. RAN examples:
  - 'NCT of Delhi' → 'NO';
  - 'Jammu & Kashmir' → 'J&'; 'Andaman & Nicobar' → 'A&';
  - 'Daman and Diu' → 'DA' and 'Dadra and Nagar Haveli' → 'DA';
  - 'Other Territory' → 'OT';
  - 'TamilNadu' → 'TAM'; 'Maharastra' (misspelt) → 'MAH';
  - city names such as 'Mumbai' → 'MUM';
  - numeric '07' → '07'.

#### Business rules in this file
Listed per constant above. The [GST RULE] tags are on FILE_PATTERNS, EXPECTED_FILE_TYPES, EXCEL_TEMPLATE_MAPPING and STATE_CODE_MAPPING. Operational values: EXCEL_SAFETY (dead), MAX_FILE_SIZE_MB (dead), extension sets (dead) and the folder-name templates.

#### Edge cases / Excel / files
None. This is pure data.

#### Gaps
- Many dead constants and keys. The extension lists are inconsistent across the project. Literal type strings are repeated instead of referring to these constants.
- The docstring "Contains all application constants, patterns, and settings." (1-4) is not quite true: the Setup tab, the extractor and setup_validation keep their own constants.

#### Pattern notes
Plain module-level dicts, read directly by importers. There is no settings file.

---

### utils/helpers.py (644 lines) - [UTIL]

- **Purpose:** shared helpers for timestamps, folders, name cleaning, file metadata, safe copying and backups, Excel-file validation (with a parallel network scan), display formatting, progress and ETA, state codes and the client-state folder key.

#### Imports
- Standard library: `os` (6), `re` (7), `time` (8), `shutil` (9), `hashlib` (10), `logging` (11), `ThreadPoolExecutor` (12), `Path` (13), `stat.S_ISREG`/`S_ISDIR` (14), `datetime` (15), typing (16), `platform` (17; comment `# ADD THIS LINE` is a leftover).
- OS-specific: `win32api` is imported **inside** `get_short_path` (584), only when `platform.system() == 'Windows'`, inside try/except, so it is **guarded**.
- Project: `from .constants import STATE_CODE_MAPPING` **inside** `get_state_code` (598).
- No unused imports (AST scan RAN). `import re` is repeated inside `create_client_state_key` (623), which is redundant.

#### Module-level constants
- `_XLSX_SIG = b'PK\x03\x04'` (264): the first 4 bytes of ZIP-based files (.xlsx/.xlsm/.xltx/.xltm).
- `_XLS_SIG = b'\xd0\xcf\x11\xe0'` (265): the first 4 bytes of old OLE-format .xls files.
- `_SIGNATURE_READ_WORKERS = 16` (271): the most parallel signature reads at once.
- `_MIN_FILES_FOR_PARALLEL = 4` (272): with fewer candidates the reads run one after another.
- Inline values:
  - minimum Excel size `1024` bytes (303);
  - accepted extensions for validation `{'.xlsx', '.xls', '.xlsm', '.xltx', '.xltm'}` (299);
  - scan extensions `('.xlsx', '.xls', '.xlsm')` (327);
  - `sanitize_filename` default `max_length = 200` with a 10-character margin (119-120);
  - `validate_client_name` 2-100 characters (438-442);
  - `get_unique_path` gives up after 100 (469);
  - hash chunk 8192 bytes, default md5 (530, 540);
  - `truncate_path` default 60 (394);
  - `create_client_state_key` default 35 and the "available > 5" rule (620, 637).

#### Functions (file order)

**`get_timestamp()` (25-27)**
- Returns `datetime.now().strftime("%d%m%y %H%M")`, for example "211026 0336" [RAN]. This is the run stamp for folder names.
- Callers: file_organizer.py:43; processing_handler.py:186. gui/main_window.py:21 imports it but never uses it.
- processing_handler.py takes a **second, separate** timestamp (186) for the report and summary names. It can differ by one minute from the folder timestamp (file_organizer.py:43) if a minute boundary passes between the two calls (LOW).

**`get_date_only()` (29-31)**: "%d%m%y". **NO CALLERS FOUND (dead)**; it is only re-exported.

**`get_safe_timestamp()` (33-35)**: "%d%m%y_%H%M". Caller: `create_backup` (224). excel_handler imports it unused.

**`ensure_path_exists(path) -> Path` (41-49)**: `mkdir(parents=True, exist_ok=True)`. Errors are logged at ERROR and re-raised. Callers: file_organizer.py:67, 99, 107, 137, 171, 194; excel_handler.py:54, 87, 378.

**`safe_path_join(*parts) -> Path` (51-66)**: drops empty parts. With nothing left it raises ValueError "No valid path components provided" [RAN]. It then joins the parts with `/`. An absolute later part replaces everything before it (`'/a', '/etc', 'x'` → `/etc/x`) [RAN]; that cannot happen with the current callers. Errors are logged and re-raised. Callers: file_organizer.py:98, 106, 136, 170, 193, 272, 331.

**`sanitize_filename(filename, max_length=200) -> str` (68-126)** (name cleaning):
1. Empty → "unnamed" (72-73).
2. `base_name = Path(filename).stem`, `extension = Path(filename).suffix` (76-77). The text after the **last dot** counts as the "extension", even for client names that are not file names (see the findings below).
3. `re.sub(r'[\s_]+', ' ', base_name)` (91): runs of spaces and underscores become one space. Added in 4219ffb (25-07-2026).
4. **Abbreviations**: `\bPrivate\b` → "Pvt" and `\bLimited\b` → "Ltd", whole words, any letter case (92-93). Word-boundary matching came in 33d0baf (23-07-2026); before that it was a case-sensitive substring replace that changed "Privateer" to "Pvter".
5. Control characters (codes 0-31 and 127-159) → "_" (99-102).
6. Each of these characters → "_": `< > : " / \ | ? * [ ] { } + = ! @ # $ % ^ , ; ' " ` ~` (96, 105-106). The hyphen is deliberately kept (95). `( ) & .` are also kept. `"` is listed twice, which is harmless.
7. Runs of spaces and underscores → one "_" (109).
8. Leading and trailing `_`, `.` and spaces are stripped (112).
9. Empty → "unnamed" (115-116).
10. If longer than `max_length − len(extension) − 10`, the name is truncated to that length (119-120). For ".xlsx" the limit is 185 characters, giving 190 in total [RAN].
11. The extension is put back unchanged (123-124).
- [RAN] examples:

| Input | Output |
|---|---|
| 'ClientA Private Limited' | 'ClientA_Pvt_Ltd' |
| 'CLIENTA PRIVATE LIMITED' | 'CLIENTA_Pvt_Ltd' |
| 'ClientA_Private_Limited' | 'ClientA_Pvt_Ltd' |
| 'Privateer Unlimited' | 'Privateer_Unlimited' |
| 'ClientAPrivateLimited' | unchanged |
| 'ClientA & Sons' | 'ClientA_&_Sons' |
| "ClientA's Shop" | 'ClientA_s_Shop' |
| 'ClientA, Delhi' | 'ClientA_Delhi' |
| 'ClientA #1' | 'ClientA_1' |
| '(1) GSTR3B-…' | '(1)_GSTR3B-…' |
| 'GSTR-2B-Reco-…-FY (1).xlsx' | 'GSTR-2B-Reco-…-FY_(1).xlsx' |
| 'GSTR3B-ClientA-Delhi-Apr!.xlsx' | 'GSTR3B-ClientA-Delhi-Apr.xlsx' (the same as the file without "!", so they collide) |
| '___' or '...' | 'unnamed' |
| 'M/s ClientA' | 's_ClientA' (the slash is treated as a folder separator; not possible in real file names) |

- **Findings**:
  - (a) **Dots in client names**. `sanitize_filename` is also called on bare client names (file_organizer.py:90, processing_handler.py:292), so the part after a dot is not cleaned: 'A.B.C Traders' → 'A.B.C Traders' (the space is kept), and 'ClientA Pvt. Ltd' → 'ClientA_Pvt. Ltd' [RAN].
  - (b) **The result depends on the Python version** for names ending in a dot. 'ClientA Pvt. Ltd.' → 'ClientA_Pvt._Ltd' on Python 3.12/3.13 but 'ClientA_Pvt._Ltd.' on 3.14, because 3.14 treats a trailing "." as a suffix [RAN exp3b]. So the same client can get different Level-4 folder and report names under different Pythons. The `.pyc` files in the early history were built by Python 3.12 and 3.13 (git: `core/__pycache__/*.cpython-312.pyc` / `-313.pyc` deleted later). Standard 13 says all testing used Python 3.14.
  - (c) Cleaning **changes the copied file's name** (spaces → "_", abbreviations, punctuation removed). Two different originals can become one name and **overwrite or skip each other** (see file_organizer).
- Callers: file_organizer.py:90, 271; processing_handler.py:292; tests/test_helpers.py.
- Business rule: Private→Pvt and Limited→Ltd abbreviation (naming, not GST).

**`get_file_info(file_path, stat_result=None) -> Dict` (129-166)**
- If no stat is given: missing file → `{'error': 'File not found'}` [RAN], otherwise `path.stat()`.
- Returns name, path (absolute), size, size_formatted, modified (`datetime` from mtime), created (`datetime` from `st_ctime`: creation time on Windows, last metadata change on Linux), extension (lower case), is_file and is_dir (from the stat mode, with no extra disk call).
- Exceptions → `{'error': str}`.
- Callers: file_parser.py:188 (with the cached stat); tests/test_file_ops.py:149-169.

**`safe_copy_file(source, destination, verify=True) -> bool` (172-215)**
1. If the source is missing or not a file: ERROR, return False (191-193).
2. Creates the destination folder (196). `shutil.copy2` copies content plus timestamps (199).
3. If verify: compares sizes; on a mismatch, ERROR, **deletes the destination**, returns False (202-208).
4. INFO "Successfully copied <name> to <dest>" (210), which puts client names in the log.
5. Any exception → ERROR, return False (213-215). **A partly written destination is not removed** when `copy2` raises half-way, for example on a network drop [READ HIGH].
6. Copying a file onto itself → False (SameFileError caught) [RAN].
- On a size mismatch in rerun mode, the previous good copy at that path was already overwritten and is then deleted.
- Caller: file_organizer.py:290.

**`create_backup(file_path) -> Optional[Path]` (217-234)**: if the file exists, it is copied with `copy2` to `<stem>_backup_<DDMMYY_HHMM><suffix>` in the **same folder**, INFO-logged, and that path is returned; otherwise None. Errors → None. **A second backup in the same minute overwrites the first** [RAN]. Caller: file_organizer.py:280.

**`validate_excel_file(file_path, stat_result=None) -> bool` (240-260)**: `_passes_metadata_checks`, then `_has_excel_signature`. **Called only by tests** (tests/test_file_ops.py:19-43, 145). App code calls the two halves directly through `scan_excel_files`. excel_handler imports it unused. The behaviour has stayed the same since 4f84ba9.

**`_has_excel_signature(path) -> bool` (275-287)**: reads the first 8 bytes and checks whether they start with `PK\x03\x04` or `\xd0\xcf\x11\xe0`. Open errors are logged at ERROR and give False. **A wrong signature gives False with no log line.** Callers: 260, 354 (thread pool), 357.

**`_passes_metadata_checks(path, stat_result) -> bool` (290-306)**
- Without a stat: the file must exist and be a file. With a stat: `S_ISREG`, a regular file.
- The extension must be one of `.xlsx/.xls/.xlsm/.xltx/.xltm` (299).
- The size must be at least 1024 bytes (303).
- Errors → False with an ERROR log.
- Callers: 258, 344.

**`scan_excel_files(folder) -> List[(Path, stat)]` (309-365)** (network-share handling):
1. One `os.scandir` listing (328).
2. Keeps names ending in `.xlsx`, `.xls` or `.xlsm`, lower-cased (327, 330). So **.xltx/.xltm are never scanned**, although `_passes_metadata_checks` would accept them.
3. `entry.stat()`; an `OSError` skips that entry (332-338).
4. Metadata filter (343-344).
5. If there are 4 or more candidates, signature reads run in a `ThreadPoolExecutor` with `min(16, n)` workers through `pool.map`. `list()` collects every result; `_has_excel_signature` catches its own errors, so none escape. Otherwise the reads run in sequence (351-357).
6. Zip, then **sort by path** (359-361). The sort order decides which duplicate is "first" (see prepare_template_data).
7. **Any exception → ERROR and `[]`** (363-365). For example the folder is unreachable, or the listing breaks half-way: everything already listed is thrown away.
- Subfolders are not entered.
- The worker threads only read files; they touch no Tk or shared state, so they are safe.
- Commit history: a603c02 (single listing), 56aaad9 (cached stat), 4f84ba9 (parallel reads; the message reports 2508 files in 40.78 s on the office share, against an estimated 677 s).
- Callers: file_parser.py:158; helpers.py:373; tests/test_file_ops.py:90-131.

**`find_excel_files(folder) -> List[Path]` (368-373)**: the paths only, from `scan_excel_files`. **Called only by tests** (tests/test_file_ops.py:58-91), so it is dead in the app.

**`format_size(size_bytes) -> str` (379-392)**: negative → "Unknown"; otherwise B/KB/MB/GB/TB with 1 decimal, dividing by 1024 [RAN: 1023 → "1023.0 B", 1536 → "1.5 KB", 10^18 → "909494.7 TB"]. Callers: 153; file_parser.py:331; tests.

**`truncate_path(path, max_length=60) -> str` (394-416)**
- Paths up to `max_length` are returned as they are.
- With 2 parts or fewer: "..." plus the last `max_length − 3` characters.
- Otherwise "first + sep + ... + sep + last" when `max_length − len(first) − len(last) − 8 > 0`, else "...sep + last". All middle parts are dropped.
- [RAN] It can return text **longer than max_length** when the last part is long (74 characters for a limit of 60). It can also produce a doubled separator when the first part is a root such as "/", "C:\\" or a UNC root ("//.../bbb" seen).
- Display only. Callers: gui/handlers/file_handler.py:24, 54 (status bar).

**`clean_windows_path(path) -> str` (418-422)**: empty or None → ""; otherwise `str(Path(path))` with "/" replaced by "\\". On Windows this does nothing. On Linux it produces paths like `\root\...` that are meaningless there [RAN]. Callers: excel_handler.py:622-697 (Links values); 567, 570, 574; tests.

**`validate_client_name(name) -> (bool, str)` (428-449)**: empty or blank → (False, "Client name cannot be empty"); under 2 characters → "Client name too short"; over 100 → "Client name too long"; any of `< > : " / \ | ? *` → "Client name contains invalid characters"; otherwise (True, "") [RAN all branches]. Callers: file_parser.py:110-111 (on the client **and** on the state); tests.

**`get_unique_path(base_path) -> Path` (451-470)**: adds `_1`, `_2`, … (before the suffix for files, after the name for folders) until the path is free; after 100 it raises ValueError [RAN: f.xlsx → f_2.xlsx when f_1.xlsx exists]. **NO CALLERS FOUND (dead)**.

**`ProgressTracker` (476-510)**
- `__init__(total, callback=None)` (479-483). Caller: processing_handler.py:203 (no callback).
- `update(increment=1, message="")` (485-491): adds to `current`, computes the percentage (0 if `total` ≤ 0), and calls the callback with **4 arguments** `(current, total, progress, message)`. FileParser/FileOrganizer callbacks take 3. No callback is ever passed. Caller: processing_handler.py:235 (increment 0 for the first client, 1 after that).
- `get_elapsed_time()` (493-495). Caller: processing_handler.py:478.
- `get_eta()` (497-510): `current == 0` → "Unknown"; otherwise rate = current / elapsed and the remaining time goes through `format_duration`. Callers: processing_handler.py:236, 457.
- Behaviour with the caller's pattern [READ HIGH]:
  - at the end of a run, `current = total − 1`, because the last client is never counted, so the "Creating summary report | ETA: …" message shows a non-zero ETA after all clients are done;
  - with one client it shows "Unknown".
- `elapsed == 0.0` would raise ZeroDivisionError. That needs two clients to fail within one clock tick: possible on Python ≤ 3.12 on Windows, whose `time.time()` ticks about every 15.6 ms (MEDIUM on that fact), and **LOW** in practice. If it happened, the error is raised at processing_handler.py:236, *outside* the per-client try, and would abort the whole run.

**`format_duration(seconds) -> str` (512-523)**: under 60 → "Ns"; under 3600 → "Mm Ss"; otherwise "Hh Mm" [RAN: −5 → "-5s"; 172800 → "48h 0m"]. Callers: 508; tests.

**`calculate_file_hash(file_path, algorithm='md5') -> Optional[str]` (529-547)**: hashes in 8 KB chunks. A missing file → None. A bad algorithm name is caught → None [RAN]. **NO CALLERS FOUND (dead)**.

**`extract_filename_without_extension(filename) -> str` (553-555)**: `Path(filename).stem`. Callers: excel_handler.py:628, 639, 650, 666, 682, 698; tests.

**`get_relative_excel_path(from_path, to_path) -> str` (557-574)**: a relative path if possible, otherwise absolute, always with backslashes. **NO CALLERS FOUND (dead)**.

**`get_short_path(file_path) -> str` (580-590)**: on Windows, `win32api.GetShortPathName(str(path))`; otherwise the path unchanged. Errors → the path unchanged (DEBUG log). On volumes where short names are switched off, the result equals the input, so the caller's retry is skipped. Caller: excel_handler.py:361.

**`get_state_code(state_name) -> str` (592-618)** **[GST RULE]** (state-code logic):
1. Empty or None → returned as it is (594-595).
2. `from .constants import STATE_CODE_MAPPING` (598). The comment "Import here to avoid circular import" does not match the code: constants imports nothing from helpers, so there is no circle (MEDIUM). Commit 6237597 changed this line from `utils.constants` to `.constants`, citing a circular-import error.
3. Trim and lower-case, then look the name up (601-606).
4. **Fallback for unknown names** (608-618):
   - 2 or more words: the upper-cased first letters of the **first two words**, e.g. 'Unknown Place' → "UP", **the same as Uttar Pradesh** (also pinned by tests/test_helpers.py:97);
   - one word: `state_name[:3].upper()`, taken from the **original, untrimmed** text, so ' Xyz' → ' XY' with a leading space [RAN];
   - both log a WARNING "Unknown state '…', using fallback: …".
5. **Non-text input** (an `int` from the Tk Treeview) crashes with AttributeError at `.strip()` (601) [RAN exp7].
- Callers: file_parser.py:229; file_organizer.py:82; excel_handler.py:616; gui/handlers/client_handler.py:169, 181, 218; processing_handler.py:293; helpers.py:628; tests.
- TIME-VARYING (the mapping table).

**`create_client_state_key(client, state, max_length=35) -> str` (620-644)**
1. `\bPrivate\b` → Pvt and `\bLimited\b` → Ltd, any letter case (624-625). **No underscore normalisation** (unlike `sanitize_filename`), so 'ClientA_Private_Limited' stays as it is [RAN].
2. `state_code = get_state_code(state)` (628). `key = "<client>-<code>"` (631).
3. If longer than `max_length`: `available = max_length − len(code) − 1`. If `available > 5`, the client part is truncated to `available` (an ending space is possible, e.g. 'ClientA Pvt -DL'); otherwise the **whole key** is cut to `max_length`, which **drops the state code** (634-642) [RAN: max 8 → 'ClientA ', max 6 → 'Client'].
4. **No character cleaning.** It relies on names coming from valid file names. A client name starting with "_" gives a Level-2 folder like '_ClientA-DL' [RAN], which the extractor skips (see Cross-file).
- Empty state → 'ClientA-' [RAN].
- Callers: file_organizer.py:83; tests.

#### Business rules in this file
- **[GST RULE]** State-name → code mapping and its fallback (592-618). TIME-VARYING.
- Name cleaning and the Pvt/Ltd abbreviation (68-126, 620-644). This is a naming rule, not a GST rule.
- What counts as a valid Excel file: extension, at least 1024 bytes, signature (290-306, 275-287, 327).
- The timestamp formats (25-35). These are run stamps, not financial years.

#### Edge cases handled
- Empty names: 72-73, 115-116.
- Control characters: 99-102.
- Long names: 119-120.
- A file vanishing between the listing and the stat: 335-338.
- A folder named like a workbook: 296-297.
- Small or corrupt files: 303, 284.
- Copy-size mismatch: 202-208.
- Missing source: 191-193.
- Unique-name limit: 469-470.
- Non-Windows short path: 583-587.
- Unknown states: 608-618.
- Zero total in ProgressTracker: 488.

#### Gaps, likely failures and silent bugs (summary)
1. Python-version-dependent cleaning [RAN].
2. Dots in client names [RAN].
3. Clean-name collisions [RAN].
4. Fallback state codes can equal real codes, and two unknown names can collide [RAN].
5. Numeric input crashes `get_state_code` [RAN].
6. The two abbreviation routines disagree on underscores [RAN].
7. The state code is dropped at small max lengths [RAN].
8. A network failure in the scan returns `[]` silently [READ HIGH].
9. A partial copy is left behind on failure [READ HIGH].
10. A same-minute backup is overwritten [RAN].
11. `truncate_path` can exceed its limit [RAN].
12. ProgressTracker's final ETA is off by one client [READ HIGH].
13. Dead helpers: `get_date_only`, `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`, plus `find_excel_files` and `validate_excel_file`, which only tests use [READ HIGH, grep].
14. Thread safety: only the signature reads run in parallel and they share nothing. OK.

#### Excel touchpoints
No workbook is opened. It reads the first 8 bytes of each candidate (282), reads whole files for hashing (539, dead), and copies files (199, 228).

#### Files and folders
- Source folder: listed (328) and read (282).
- Destination files: written (199), deleted on a mismatch (207).
- Backups written next to the original copy (228).
- Folders created (45, 196).

#### Comments and docstrings vs code
- 26 "(DDMMYY HHMM)" ✓; 31 ✓; 34 ✓.
- 79-90: the abbreviation reasoning ✓ [RAN].
- 95 "WITHOUT including hyphen" ✓.
- 98 "ASCII 0-31 and 127-159" ✓.
- 111 "underscores and dots": spaces are also stripped.
- 135-139, 247-250, 263-272, 312-319 and 325-326, 333: the network round-trip notes. They are consistent with the code, and the measured figures come from the commit messages.
- 581 "(Windows only)" ✓.
- 597: the circular-import note is stale or unneeded.
- 615 "take first 2-3 characters": the code always takes up to 3.
- 622 ✓; 637 ✓.
- Module docstring: "Robust utility functions with comprehensive error handling." (1-4).

#### Pattern notes
- Errors: helpers return False, None, `[]` or an error dict, except `ensure_path_exists` and `safe_path_join`, which re-raise.
- Logging: the module logger. INFO lines include client file names (210, 229).
- Lazy imports are used for `win32api` and `STATE_CODE_MAPPING`.

---

## Cross-file observations

### Shared state: write path vs read path
| Shared thing | Written / built at | Read / used at | Consistent? |
|---|---|---|---|
| Client key `"<raw client>-<STATECODE>"` | core/file_parser.py:229-230 | gui/handlers/client_handler.py:169-170 (rebuilt from Treeview `values[0]`, `values[1]`); :181, :218; core/file_organizer.py:122; processing_handler.py:232 `client_data[client_key]` | Yes for ordinary text. **No when the client or state looks like a number**: Treeview hands back ints (RAN exp7: ('007', '07') → [7, 7]). Then `get_state_code(7)` raises AttributeError, so selecting clients crashes, and a client like '007' becomes key '7-…', which gives a KeyError at processing_handler.py:232, *outside* the per-client try, aborting the whole run. HIGH |
| `client_folder_settings` | client_handler.py:49 (keyed by the parser key), :185, :219 | file_organizer.py:130 | Yes since 33d0baf (23-07-2026; before that the GUI keyed by state *name*), except for numeric-looking names as above |
| Folder tree "Annual Statement-*" / client / "Version-*" / ITC_Report_*, Sales_Report_* | file_organizer.py:103, 167, 190; FOLDER_STRUCTURE (constants.py:82, 84); report names at processing_handler.py:297, 355 | power_query_extractor/gui/extractor_window.py:623 (glob "Annual Statement-*", latest by mtime, the same rule as file_organizer.py:184), :638-639 (**skips client folders whose name starts with "_"**), :641-648 (latest "Version-*" by mtime), :36-39 (lower-case prefixes 'itc_report_' / 'sales_report_') | Mostly. **Mismatch:** a client whose parsed name starts with "_" gets Level-2 folder "_<client>-XX" (`create_client_state_key` does not strip "_", RAN), and the extractor ignores it. HIGH (by reading both sides); rare |
| Links-sheet cells | excel_handler.py:189-195 / 297-301 using EXCEL_TEMPLATE_MAPPING | the templates' Power Queries (not in the repository) | **Cannot verify.** Not sure. How to check: open both templates and compare each query's source step with Links B2…B11 |
| Copied file names | file_organizer.py:271-274, 237-244 (`name` replaced on Success) | excel_handler.py:629, 640, 651, 667, 683, 699 (stems written to Links) | Yes on Success. **No for Skipped (resume) or Failed files**: the original name is written while the copy has the cleaned name (RAN exp8) |
| Temp staging file | excel_handler.py:154 / 279 | :159 / 282; deleted at 258 / 320 | Yes. It is left behind when the `finally` raises (RAN with fakes) |
| Backups `_backup_DDMMYY_HHMM` | helpers.py:225-228, inside the category folders | no code reads them; any template query reading a whole folder would see them | Possible double reading (MEDIUM, template-dependent) |
| `_Organization_Report_*.txt` | file_organizer.py:330-371 | nobody | — |
| `GST_Processing_Summary_*.xlsx` (Level 1) | excel_handler.py:402, via processing_handler.py:462-469 | nobody (the extractor iterates only folders, extractor_window.py:639) | — |
| Mode labels vs behaviour | constants.py:174-190; setup_tab.py:103-110 | file_organizer.py:165-196, 277-287 | **Resume does not behave as labelled** (RAN) |
| Workflow help text | setup_tab.py:96-97: 'organize' "Fast - no Excel automation involved." | processing_handler.py:309, 376 always call `create_report_from_template`, which on Windows with pywin32 drives Excel through COM (excel_handler.py:90-95) | **Text does not match the code**: organise mode does drive Excel. HIGH |
| Setup help texts | setup_tab.py:74-76 ("nothing here is moved, renamed or deleted"), 77-82 ("the template itself is never modified"), 83-85 ("Each run adds a new timestamped folder, so previous runs are left untouched"), 112-117 (example "'ITC (ABC Pvt Ltd)'") | file_organizer / excel_handler | The first two are **true** (only `copy2` from sources and templates; READ HIGH). The third is true only in fresh mode: rerun and resume reuse the latest run folder, and rerun in the same minute overwrites. The fourth: the real folder uses underscores, e.g. 'ITC (ABC_Pvt_Ltd)' (RAN) |
| Log file | my files log through `logging` (INFO includes client names and paths: excel_handler.py:153, 158, 195, 214, 220; file_organizer.py:57, 141, 185, 292; helpers.py:210, 229; file_parser.py:131) | main.py:20-27 (`FileHandler('gst_organizer.log')`, relative to the **current folder**, level INFO) | The file collects client names and file names. That is presumably how `gst_organizer.log` got committed in edac6a3 (A-R4 / history reader) |

### Inconsistent patterns found
- **Excel extension lists differ across the project**:
  - constants.py:196 `EXCEL_EXTENSIONS` {xlsx, xls, xlsm} (unused);
  - constants.py:197 `TEMPLATE_EXTENSIONS` {xlsx, xltx, xltm} (unused);
  - helpers.py:327, scan: (.xlsx, .xls, .xlsm);
  - helpers.py:299, validation: {.xlsx, .xls, .xlsm, .xltx, .xltm};
  - FILE_PATTERNS regexes: `\.xlsx?` (xlsx and xls only);
  - gui/utils/setup_validation.py:21 `TEMPLATE_SUFFIXES` {xlsx, xltx, xltm, xlsm};
  - gui/handlers/file_handler.py:31/42 dialog "*.xltx *.xlsx".
- **Signature bytes are defined twice**: constants.py:199-202 (unused) and helpers.py:264-265.
- **Two abbreviation routines**: helpers.py:91-93 normalises underscores; helpers.py:624-625 does not.
- **Three different "client name" forms**:
  - the raw parsed name (parser key, organisation report name);
  - the abbreviated and truncated key (Level 2);
  - the cleaned, underscored, full-length name (Level-4 suffix, copied files, report file names at processing_handler.py:292-297).
- **Two ways to get Excel**: excel_handler.py:124-127 (`GetActiveObject`, then `Dispatch`) against report_processor.py:218, 536 (`Dispatch` only). Neither uses `DispatchEx` as Standard 12 asks.
- **Progress callback signatures**: 3 arguments (file_parser.py:52, file_organizer.py:226) against 4 (helpers.py:491).
- **Logger style**: module loggers everywhere except excel_handler.py:26 and 28 (root `logging`).
- **Error style**: mostly "log and return False/None"; `ExcelHandler` raises for a missing template; FileParser collects errors that nobody reads.
- **Leftover AI pasting comments**: `# ADD THIS LINE` (helpers.py:17), `# ADD THESE` (file_organizer.py:15), `# ADD THESE TWO` (excel_handler.py:39), `# NEW` (excel_handler.py:449-451), "[DEBUG]" in an ERROR message (excel_handler.py:83).
- `warnings.filterwarnings('ignore')` (excel_handler.py:17) silences warnings for every module, including those of other readers' files.

### Threading and UI-thread notes touching my files
- The scan (`FileParser.scan_folder`, `scan_excel_files`) runs on the Tk UI thread (file_handler.py:72), so the window freezes during a long network scan [READ HIGH].
- `FileOrganizer` and `ExcelHandler` run on the processing worker thread (processing_handler.py:129-134). They mutate `app.client_data` records the UI thread also reads. The Scan and Re-scan buttons' enabled state depends only on Setup readiness (gui/tabs/setup_tab.py:585-591); nothing disables them while processing runs (grep for `scan_btn`/`rescan_btn`: no other references) [READ HIGH]. If Re-scan is pressed during a run, `rescan_files` clears `app.client_data` (file_handler.py:103), and the worker's next `self.app.client_data[client_key]` (processing_handler.py:232) raises KeyError outside the per-client try, so the whole run stops with "Fatal error" (HIGH by reading; not run; A-R3 to confirm the Setup tab stays reachable during processing).
- `organize_files` calls the progress callback from the worker thread; the callback marshals onto the UI thread through `root.after` (processing_handler.py:266-269, 33-36).

### Dead code found in my files (no callers in the app, verified by searching all .py files)
- core/excel_handler.py: `__del__` (it never acts), `safety_config`, `excel_app` (always None), the local `state_code` at 616, unused imports.
- core/file_organizer.py: `cleanup_on_error`, `_is_folder_empty`, `get_summary`.
- core/file_parser.py: `validate_results`; the parse-result keys `warnings`, `metadata`, `category`, `folder` are written but never read; `self.errors` is never read outside the class.
- utils/helpers.py: `get_date_only`, `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`. `find_excel_files` and `validate_excel_file` are used by tests only.
- utils/constants.py:
  - `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `EXCEL_SAFETY` (all of its values);
  - `FOLDER_STRUCTURE['level2']`;
  - `GUI_CONFIG['fonts']`, `GUI_CONFIG['colors']['white']`;
  - `PROCESSING_MODES[*]['icon']`;
  - MESSAGES `errors.invalid_excel`, `errors.template_error`, `errors.processing_error`, and all of `success.*` and `info.*`;
  - `FILE_PATTERNS[*]['category']` and `['folder']` (copied, never read).
- utils/__init__.py and core/__init__.py: every re-export is unused by app code.

### Other observations
- `/root/merge-scratch/A-copy` contains `__pycache__` folders timestamped 21-10-2026 03:21, which is before my run. They are **not** in the repository (`git ls-tree -r A/main` lists no `__pycache__`). They were made by whatever set up the copies. My runs wrote nothing into A-copy or B-copy (checked with `find -newer`).
- Commit messages 4f84ba9 and 7347383 name the office server and share path. That is infrastructure, not client data; I have not copied it.
- No templates (`.xlsx`/`.xltx`) are in the repository, so anything that depends on the Links sheet or the queries cannot be checked here.

## Possible client names

None found in my seven files, or in the commit messages and diffs of the commits that touched them. The only company-like strings are generic examples: "ABC_Private_Limited", "ABC PRIVATE LIMITED", "Privateer", "Unlimited" (helpers.py:82, 89-90), and "ABC" in the commit messages 33d0baf, 4219ffb and 5c8ce95. I treat these as obviously generic.

## Questions for the author

1. **Resume mode**: should "Resume" reuse the interrupted run's own Version folder? Today it always makes a new one, so it skips nothing unless it is started within the same minute (RAN).
2. **[GST RULE] Which file feeds the report** when a client has two files of the same type, for example an original and a "(1)" copy? Today the alphabetically first one is used, and "…(1).xlsx" sorts before the original (RAN). I am only describing this; you decide.
3. **Links cells for missing files**: when a client has no IMS Reco (for example), the IMS cells are left exactly as they are in the template. Do your templates keep sample paths or names in those cells? (RAN on the openpyxl path; the COM path uses the same logic.)
4. Do the ITC/Sales templates read the **whole GSTR3B folder** (all files in it)? If so, backup copies (`…_backup_DDMMYY_HHMM.xlsx`) and "(1)" duplicates placed in that folder would also be read.
5. Are your templates `.xlsx`, `.xltx`, `.xltm` or `.xlsm`? The output is always named `.xlsx`; with a template or macro file this may not open in Excel (not verified).
6. **[GST RULE] State text in file names**: what do real names contain in the state position: full names, short forms, two-digit numbers? A state written as digits (for example "07") crashes client selection in the GUI (RAN). Unknown names are turned into initials, which can equal a real state's code ('Unknown Place' → 'UP'). How should states be identified? (Asking only, not suggesting.)
7. **[GST RULE] Financial year and period**: the period, month, date and year parts of file names are captured but never checked or compared. Is mixing files of different periods for the same client intended to be impossible in your workflow?
8. **[GST RULE] GSTR-3B count**: any number of GSTR-3B files is accepted for a client with no warning. Is any check expected? (Asking only.)
9. Do you ever run Step 3 while your own Excel is open with work in it? The tool attaches to that Excel, hides it and then quits it (READ; RAN with fake COM; not tried with real Excel).
10. Is the target folder chosen as a mapped drive (e.g. `Z:\…`) or a `\\server\share` path, and is the same mapping on every PC that refreshes the reports? The Links cells store the full path as chosen.
11. Annual Report and other source files may be `.xls` or `.xlsx`, but the Links cells store the name without the extension. Do the templates assume `.xlsx`?
12. Should the "Max client folder name length" setting also shorten the client name in the Level-4 folders (`ITC (…)`) and in the report file names? Today it only shortens the Level-2 folder, and at values of 8 or 9 or below it drops the state code entirely.
13. Client names with hyphens (for example "<Client A>-East") are silently split into a wrong client and state for GSTR-2B Reco, Sales Reco and Annual Report files, and rejected for GSTR-3B and Sales. Do real client names ever contain hyphens?
14. The same client spelt two ways in file names ("… Private Limited" vs "… Pvt Ltd", or a different letter case) becomes two clients that share one output folder, and their ITC/Sales report names collide. Does this happen in your source folders?

## INCOMPLETE

Nothing in my assignment is unread: all 2,497 lines of the seven files (13 + 713 + 426 + 371 + 38 + 292 + 644) were read in full, and the callers were traced. Two limits (not gaps in reading):
- Real Excel behaviour on the COM path (what Excel does on `Quit` with unsaved user workbooks, whether a `.xltx`/`.xlsm` package renamed `.xlsx` opens, how `Range.Value` treats numeric-looking text) **could not be run** on this Linux machine. It was traced by reading and with fake COM objects only. How to check: on a Windows PC with Excel, process one test client with (a) Excel already open with an unsaved workbook, (b) a `.xltx` template, (c) a template without a Links sheet.
- Template-dependent behaviour (what the queries read from Links B2…B11) cannot be checked, because no template is in the repository.
