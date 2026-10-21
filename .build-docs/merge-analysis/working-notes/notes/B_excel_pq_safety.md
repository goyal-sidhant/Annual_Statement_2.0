# B-X-EXCEL - Project B (AnnualStatement_oldv2, snapshot b0e7d0e, 09-02-2026): Excel and Power Query safety

Working notes for the formal audit. Written 21-10-2026 by the B-X-EXCEL reader.
All paths are repository-relative to project B unless stated. Line numbers are those of snapshot b0e7d0e.

---

## 0. How to read these notes

### 0.1 Evidence marks
- **READ** = I read the code line by line.
- **RAN** = I executed B's own code on this Linux machine (Python 3.14.6, openpyxl 3.1.5, PyQt5 5.15.11, from `/root/merge-scratch/venv-B`), in a copy at `/root/merge-scratch/run/B-X-EXCEL/app`.
- **RAN (fake COM)** = I executed B's own code, but with stand-in versions of the Windows-only pywin32 modules (`/root/merge-scratch/run/B-X-EXCEL/fakewin/`) that do nothing except record, in order, every request B makes to "Excel". This proves **what B asks Excel to do and in what order**, and how B's own code reacts to the errors I simulated. It does **not** prove how real Excel behaves; where real-Excel behaviour matters I say so and give a confidence level and a way to check.
- **HIGH / MEDIUM / LOW** = my confidence in an inference (anything beyond what the lines literally say).
- **[GST RULE]** = the code decides something the author owns (which return/table/cell a figure belongs to, state codes, etc.). Described only, never judged. **TIME-VARYING** = a value that can change over time (table numbers, state list, etc.).

### 0.2 Plain-English glossary (first use)
- **COM** (Component Object Model): the Windows mechanism that lets one program (here B, written in Python) remote-control another program (here Microsoft Excel). Requires Windows and Excel.
- **pywin32 / win32com / pythoncom / win32gui / win32api / win32con**: the Python add-on package that gives Python access to COM and to Windows windows/keyboard functions. Windows-only.
- **Excel instance**: one running copy of EXCEL.EXE. A user can have their own instance open; a program can start a separate hidden one.
- **`Dispatch`, `DispatchEx`, `GetActiveObject`**: three ways to get an Excel instance. `GetActiveObject` = grab the Excel that is already running (normally the user's own). `Dispatch` = "give me an Excel" (normally starts one, but the author's Standard 12 forbids it because it is not guaranteed to be separate). `DispatchEx` = always start a brand-new separate Excel (what Standard 12 requires).
- **openpyxl**: a Python library that reads/writes .xlsx files directly, without Excel. When it saves, it rewrites the whole file and silently drops parts it does not understand, including Power Query.
- **Power Query (PQ)**: Excel's "Get & Transform" queries that pull data from other files/folders into a workbook. Stored inside the .xlsx in a part called **DataMashup** (`customXml/item1.xml`) plus **connections** (`xl/connections.xml`).
- **Refresh**: re-running the queries so the workbook picks up the latest data. **BackgroundQuery**: if on, a refresh runs in the background and Excel returns control before it finishes.
- **QThread / worker thread**: B runs long jobs in a background "thread" (a second line of execution) so the window does not freeze. COM work in such a thread needs **`pythoncom.CoInitialize()`** at the start and **`CoUninitialize()`** at the end.
- **`finally` block**: code that Python runs at the end of a `try`, whether or not an error happened. Used for clean-up.
- **SendKeys / AppActivate**: Windows Script Host functions. `AppActivate(title)` tries to bring a window to the front; `SendKeys(keys)` types keystrokes into **whatever window currently has keyboard focus**, exactly as if a person pressed them.
- **KeyTips**: Excel's ribbon keyboard shortcuts (Alt, then letters).
- **glob**: a filename pattern search such as `ITC_Report_*.xlsx` (`*` = anything).
- **`shutil.copy2`**: Python's file copy (copies contents and the file's dates).
- **UNC path**: a network path like `\\server\share\...`. **NTFS**: the Windows disk format.
- **Modal dialog**: a pop-up box that blocks the program until answered.

### 0.3 What I read
READ in full, every line: `core/excel_handler.py` (626 lines), `power_query_extractor/core/report_processor.py` (751), `power_query_extractor/core/data_consolidator.py` (416), `power_query_extractor/config/cell_mappings.py` (277), `power_query_extractor/extractor_main.py` (78), `gui/controllers/processing_controller.py` (298), `gui/controllers/extraction_controller.py` (184), `gui/controllers/scan_controller.py` (94), `gui/pyqt_main_window.py` (608), `gui/views/extraction_view.py` (456), `core/file_organizer.py` (408), `core/file_parser.py` (398), `utils/helpers.py` (536), `utils/constants.py` (274), `models/app_state.py`, `models/cache_manager.py`, `models/run_history.py`, `models/signals.py`, `main.py`, `launch_extractor.py`, `requirements.txt`, `.gitignore`, `README.md` (363), every `__init__.py`.
READ in part: `gui/views/setup_view.py` lines 200-240 and 600-647 (template browse dialogs, `set_values`); `gui/views/validation_view.py` and `gui/views/processing_view.py` only the signal lines (validation_view.py:18-19, 184, 217; processing_view.py:16, 95). A grep (search) of every .py file for the task's Excel/file keyword list, plus `os.remove|unlink|rename|shutil.move|rmtree|subprocess|os.system|startfile|open(`, found no Excel or file touchpoint in the parts I did not read (setup_view.py rest, validation_view.py, processing_view.py, gui/theme/theme_manager.py, gui/widgets/*.py). See INCOMPLETE.
Git history READ: `git log` of B/main; full Excel-relevant diffs of c258047, bd176bf, b0e7d0e, 7e09e5c; `git grep` of the shared initial commit edac6a3 and split point 47a1492 for the COM patterns.

### 0.4 Experiments I ran (all under `/root/merge-scratch/run/B-X-EXCEL/`)
| # | Script | What it shows |
|---|---|---|
| E1 | `app/main.py`, `app/launch_extractor.py`, `python -m power_query_extractor.extractor_main` (no pywin32 in this environment) | What happens when pywin32 is missing |
| E2 | `exp/make_pq_template.py` + `exp/run_openpyxl_fallback.py` | B's openpyxl fallback on a synthetic workbook that carries a DataMashup part and a connections part |
| E3 | `exp/trace_excel_handler.py` (3 scenarios), `exp/trace_eh_crash_after_copy.py` | Step 3 report creation through COM: call order with no Excel running, with the user's Excel running, and with Excel dying (fake COM) |
| E4 | `exp/trace_full_processing.py` (normal and `block_summary`) | Whole Step 3 for one placeholder client (<Client A>): files, COM call order, what the Processing tab shows (fake COM) |
| E5 | `exp/trace_report_processor.py` (scenarios normal, error_dialog, crash_on_save, open_rejected with NFAIL=3 and 12, already_open_user_excel, empty_suffix) | Step 4 refresh + extraction call order (fake COM) and B's real consolidated report built from the result (openpyxl, real) |
| E6 | `exp/trace_suffix_skip.py` | Custom file suffix + "Skip refresh", and which files the report pattern matches |
| E7 | one-line Python checks | How Level-1 folder names sort; template values left in the Links sheet |

The synthetic workbooks hold only placeholder text (no client data). The dummy figures used in E5 are 101, 102, ... chosen by me.

---

## 1. Most important findings (details and evidence in the sections below)

1. **The Power Query refresh is done by fake keystrokes, not by Excel's refresh command.** Step 4 makes Excel visible and maximised (report_processor.py:181-183), then types Alt+A, R, A with `SendKeys` (report_processor.py:216-248), waits a fixed time, and saves (report_processor.py:253-305). There is no `RefreshAll`, no `BackgroundQuery = False`, no `CalculateUntilAsyncQueriesDone`, no check that a connection is still refreshing. READ; call order RAN (fake COM). Standard 12 rules R7-R9 NOT MET.
2. **The keystrokes go to whatever window has keyboard focus.** `AppActivate(excel.Caption)` (report_processor.py:222) asks Windows for a window titled "Excel"; its true/false answer is ignored (activation is assumed when no error is raised, report_processor.py:221-224). B also types the keys when the workbook failed to open at all (RAN fake COM, E5 open_rejected NFAIL=12). If the user's own Excel, or any other program, has focus, those keystrokes act there. READ HIGH; real-Windows outcome MEDIUM.
3. **Step 3 takes over the user's own running Excel, hides it and closes it.** `GetActiveObject` first (excel_handler.py:111), then `Visible = False` on that instance (excel_handler.py:118), and `Quit()` always runs (excel_handler.py:207-211) because `self.excel_app` is never set to anything but `None` (excel_handler.py:55; only other uses 59, 61, 207). `DisplayAlerts` is switched back on just before `Quit()` (excel_handler.py:205). RAN (fake COM, E3 "user_excel_open"): the user's instance is hidden and told to quit while the user's workbook is still open in it. What real Excel then does (save prompt, possibly invisible, possibly blocking B forever) is MEDIUM/LOW - see Q2.
4. **An "error-dialog closer" scans every window on the desktop and clicks OK/Close/Yes/Continue.** report_processor.py:368-404 (any visible, enabled window whose title contains "microsoft excel", "error", "problem", "cannot", "failed", "unable", "could not", "exception", "invalid" or "missing", and whose window class is `#32770`, `NUIDialog` or `bosa_sdm_msword`), then report_processor.py:443-476 clicks the first button whose text contains "ok", "close", "yes" or "continue", or sends a close message. Not limited to B's own Excel. READ HIGH; RAN (fake COM, E5 error_dialog).
5. **"Refresh succeeded" is judged only by Info!BB2 being non-empty** (report_processor.py:589-636), and the refreshed copy is **saved even when an error dialog was seen or the BB2 check failed** (save at report_processor.py:305 happens before the result is decided at 310-330). BB2 can already hold a value from an earlier refresh or from the template, so it is not proof that this refresh ran (same weakness Standard 12 describes for `RefreshDate`). READ HIGH; RAN (fake COM, E5 error_dialog).
6. **The consolidated Extract report puts the wrong figures in three sheets.** In `PQ_Extraction_Report_*.xlsx`, the IGST/CGST/SGST columns of the sheets "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" receive the values read for "GSTR-3B Difference" (cells J15/K15/L15 of sheet "Diff GSTR-3B"), not the cells their own configuration names. Cause: a substring (part-of-text) match at data_consolidator.py:130-133 that is tried before the exact-key lookup at 136-138. RAN (E5 normal: IMS Reco row shows 101/102/103 instead of 501/502/503). [GST RULE] context - the mapping itself is the author's; the defect is that the code does not follow it. HIGH.
7. **When a refresh fails, every figure of that report is replaced by "Missing"** in the consolidated report even though B did read the cells (data_consolidator.py:116-117; extraction runs regardless, report_processor.py:143-144). RAN (E5 error_dialog).
8. **Without pywin32 the app does not start at all**, so the openpyxl fallback that would strip Power Query is unreachable in B's HEAD. `report_processor.py:11-15` imports pywin32 unconditionally and is pulled in by `gui/__init__.py:7` -> `gui/pyqt_main_window.py:27` -> `gui/controllers/extraction_controller.py:10`. RAN (E1): `main.py` prints "Missing module: No module named 'pythoncom'" and advises "pip install PyQt5 openpyxl" (main.py:49-54), which would not fix it; `requirements.txt:14` has pywin32 commented out. When the fallback is reached (E2, run in isolation), DataMashup and connections are dropped and B reports success. The fallback **was** reachable before the split (47a1492), when the organiser did not import the extractor (see 8.2).
9. **All Python warnings are silenced for the whole program** (excel_handler.py:17 `warnings.filterwarnings('ignore')`), including openpyxl's warnings that announce it is dropping workbook features. READ HIGH.
10. **If Excel dies during Step 3, clean-up is skipped and a half-made report is left behind with the normal name.** The `finally` block sets `excel.ScreenUpdating`/`DisplayAlerts` without protection (excel_handler.py:204-205); if Excel is gone, that raises, so `Quit()` (209) and `CoUninitialize()` (215) never run and the error escapes. The copied-but-unfilled template stays at the report path (copy at 124 happens before the open at 129). RAN (fake COM, E3 excel_crash and E3b crash-after-copy).
11. **Template values survive in the Links sheet when a client has no file of that type** (`if value:` at excel_handler.py:162 and 249). The report's queries would then read whatever path the template already holds. RAN (E7: B10/B11 kept the template's values).
12. **Step 4 can work on the wrong batch folder.** The Extract tab treats the Level-1 folder with the highest *name* as the latest (pyqt_main_window.py:432-436, 448). Names are "Annual Statement-DDMMYY HHMM", so 31-10-2025 sorts after 01-02-2026. RAN (E7). Step 3's own "latest" uses modification time instead (file_organizer.py:167). HIGH.
13. **The report search pattern also matches B's refreshed copies.** `ITC_Report_*.xlsx`/`Sales_Report_*.xlsx` with `report_files[0]` (report_processor.py:76-89). RAN (E6): on this Linux disk the refreshed copy was picked and a copy of the copy was made and refreshed. On a local Windows NTFS disk alphabetical order usually puts the original first (MEDIUM); on a NAS/Samba share order is not guaranteed (MEDIUM). "Skip refresh" looks only for `_Refreshed_` (report_processor.py:102), so with a custom suffix it silently refreshes again (RAN E6).
14. **Nothing ever kills or times out Excel.** No `taskkill`, no timeout on any COM call, Stop is checked only between clients (extraction_controller.py:61; processing_controller.py:81). A hung Excel hangs the worker; Excel can be left running. READ HIGH.
15. **Steps 3 and 4 can run at the same time** (no mutual lock between `ProcessingController` and `ExtractionController`; pyqt_main_window.py:345-371, 506-547). Step 3's `GetActiveObject` could then grab Step 4's visible Excel. READ HIGH (fact) / MEDIUM (consequence).
16. **COM clean-up order is not what Standard 12 requires**: `CoUninitialize()` is called inside the same function while Excel objects are still held by local variables (excel_handler.py:215; report_processor.py:363, 584). READ HIGH.
17. **A failed summary report is silent** in the Processing tab (excel_handler.py:344-346 returns False; processing_controller.py:190 shows a message only on True). RAN (E4 block_summary).
18. **Extracted figures are copied into three places outside the reports**: the Extraction log text file in the client batch folder, `gst_organizer.log` in the folder the app was started from, and `run_history.json` in the user's home folder (Output 5). READ HIGH.

---

## 2. Output 1 - Every place B touches an Excel workbook (touchpoint table)

Thread names: **PW** = QThread worker `ProcessingWorker` (Step 3; started at processing_controller.py:265-279). **EW** = QThread worker `ExtractionWorker` (Step 4; extraction_controller.py:148-165). **SW** = QThread worker `ScanWorker` (Step 1; scan_controller.py:65-80). **Main** = the window's own thread.
"Can it hold PQ?" for template-derived files: **yes (intended)** = the whole design refreshes these files (README.md:14, 24; report_processor.py docstring line 2), but the actual template files are not in the repository, so I could not confirm that every template has queries.

| # | file:line | function | library | action | which workbook | can it hold PQ? | thread | mark |
|---|---|---|---|---|---|---|---|---|
| 1 | core/excel_handler.py:20-28 | module level | COM (import) | decide `USE_WIN32COM` (True only if `platform.system()=='Windows'` and `win32com.client` + `pythoncom` import) | - | - | Main (at import) | READ, RAN (False on Linux) |
| 2 | core/excel_handler.py:17 | module level | Python `warnings` | silence **all** warnings process-wide | all | - | Main | READ |
| 3 | core/excel_handler.py:31-33 | module level | openpyxl (import) | unconditional import | - | - | Main | READ |
| 4 | core/excel_handler.py:52-53 | `ExcelHandler.__init__` | file system | create `%TEMP%\gst_excel_temp` folder; never used afterwards (grep: `temp_dir` only at 52-53) | none | - | PW (handler made at processing_controller.py:54) | READ |
| 5 | core/excel_handler.py:57-63 | `__del__` | COM | `Quit()` on `self.excel_app` - dead: `excel_app` is always `None` | - | - | whichever thread frees the object | READ |
| 6 | core/excel_handler.py:107 | `_create_report_win32com` | COM | `pythoncom.CoInitialize()` | - | - | PW | READ, RAN (fake COM) |
| 7 | core/excel_handler.py:110-112 | same | COM | `GetActiveObject("Excel.Application")` - attach to the user's running Excel | user's Excel (all their open books) | yes/unknown | PW | READ, RAN (fake COM) |
| 8 | core/excel_handler.py:113-115 | same | COM | `Dispatch("Excel.Application")` if no Excel running | new instance (not guaranteed separate) | - | PW | READ, RAN (fake COM) |
| 9 | core/excel_handler.py:118-120 | same | COM | `Visible=False`, `DisplayAlerts=False`, `ScreenUpdating=False` - applied to the user's instance too when #7 succeeded | instance | - | PW | READ, RAN (fake COM) |
| 10 | core/excel_handler.py:124-125 | same | file copy (`shutil.copy2`) | copy template -> `Version-.../ITC_Report_<client>_<ST>_<DDMMYY_HHMM>.xlsx` (or `Sales_Report_...`), overwriting any file of that name; then sleep 0.5 s | template (source, read only) -> newly created report | yes (intended) | PW | READ, RAN |
| 11 | core/excel_handler.py:129 -> 264-306 | `_safe_open_workbook` | COM | look through `excel.Workbooks` for the same full path and reuse it (274-277); else `Workbooks.Open(path)` with **no** options (286); on error containing "-2147352567" or "path", or path > 218 chars, retry with the 8.3 short path (292-301); else re-raise (306) | newly created report | yes (intended) | PW | READ, RAN (fake COM) |
| 12 | core/excel_handler.py:136-153 | `_create_report_win32com` | COM | find sheet named `Links` (case-insensitive), else first sheet containing "link", else **add a new sheet** and name it `Links` (152-153; message only in the log file) | newly created report | yes | PW | READ |
| 13 | core/excel_handler.py:159-167 | same | COM | write folder paths / file names (as text) into the Links cells from `EXCEL_TEMPLATE_MAPPING` (utils/constants.py:96-120); **empty values are skipped** (162) so the template's own value stays | newly created report | yes | PW | READ, RAN (fake COM) |
| 14 | core/excel_handler.py:173, 176-177 | same | COM | `wb.Save()` then `wb.Close(SaveChanges=True)` (saves twice); `wb=None` | newly created report | yes | PW | READ, RAN (fake COM) |
| 15 | core/excel_handler.py:183-189 | same | file system | "verify": only checks the file exists (size is logged, not checked) | newly created report | - | PW | READ |
| 16 | core/excel_handler.py:197-201 | same (`finally`) | COM | `wb.Close(SaveChanges=False)` if a workbook is still held (errors swallowed) | newly created report | yes | PW | READ |
| 17 | core/excel_handler.py:203-205 | same (`finally`) | COM | `ScreenUpdating=True`, `DisplayAlerts=True` - **not** inside a try | instance (user's if #7) | - | PW | READ, RAN (fake COM: raises when Excel is gone) |
| 18 | core/excel_handler.py:206-211 | same (`finally`) | COM | `Quit()` "if not already open" - but the test `self.excel_app is None` is always true, so it **always** quits, including the user's instance | instance (user's if #7) | - | PW | READ, RAN (fake COM) |
| 19 | core/excel_handler.py:213-217 | same (`finally`) | COM | `pythoncom.CoUninitialize()` (errors swallowed); skipped if #17 raised | - | - | PW | READ, RAN (fake COM) |
| 20 | core/excel_handler.py:228 | `_create_report_openpyxl` (fallback) | file copy | copy template -> report path | template -> new report | yes (intended) | PW | READ, RAN |
| 21 | core/excel_handler.py:231 | same | openpyxl | `load_workbook(report, data_only=False, keep_links=True)` (macros not kept: `keep_vba` default False) | new report (copy of template) | yes | PW | READ, RAN |
| 22 | core/excel_handler.py:239-252 | same | openpyxl | find sheet exactly named Links (case-insensitive; no "link" fallback, no sheet creation); write non-empty values | new report | yes | PW | READ, RAN |
| 23 | core/excel_handler.py:255-256 | same | openpyxl | **`wb.save(report)` over the copy of the template** - drops Power Query; `wb.close()` | new report | **yes - and it is lost** | PW | READ, RAN (E2) |
| 24 | core/excel_handler.py:258 | same | file system | return True if the file exists | new report | - | PW | READ, RAN |
| 25 | core/excel_handler.py:318-340 | `create_summary_report` | openpyxl | `Workbook()` (brand-new), 5 sheets, `save` to `Level1/GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (overwrites same name), `close` | newly created summary | no (created by B) | PW | READ, RAN (E4) |
| 26 | power_query_extractor/core/report_processor.py:11-15 | module level | COM (import) | **unconditional** import of `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` | - | - | Main (at app start) | READ, RAN (E1: app cannot start without them) |
| 27 | report_processor.py:76-77, 89 | `_process_report` | glob | `glob("ITC_Report_*.xlsx")` / `("Sales_Report_*.xlsx")` in the chosen Version folder; take the first match | report created in Step 3 - **or a refreshed copy** (pattern also matches them) | yes | EW | READ, RAN (E6) |
| 28 | report_processor.py:92-94 | same | - | refreshed name = `<stem><suffix>.xlsx`, suffix from the Extract tab (default `_Refreshed_{timestamp}`, timestamp `DDMMYY_HHMM`) | - | - | EW | READ |
| 29 | report_processor.py:97-115 | same | glob, file dates | if "Skip refresh": newest (by modification time) `<stem>_Refreshed_*.xlsx`; the custom suffix is ignored here | earlier refreshed copy | yes | EW | READ, RAN (E6) |
| 30 | report_processor.py:118-120 | same | file copy | `shutil.copy2(original, refreshed)` - overwrites an existing refreshed copy of the same name; empty suffix -> `SameFileError` (caught at 154) | original report (read) -> new refreshed copy | yes | EW | READ, RAN (E5 empty_suffix) |
| 31 | report_processor.py:171 | `_refresh_power_query_simple` | COM | `CoInitialize()` | - | - | EW | READ, RAN (fake COM) |
| 32 | report_processor.py:174 | same | COM | `Dispatch("Excel.Application")` | instance | - | EW | READ, RAN (fake COM) |
| 33 | report_processor.py:178-190 | same | COM | up to 5 tries: `Visible=True` ("keep it visible to avoid issues"), `DisplayAlerts=False`, `WindowState=-4137` (maximised); after 5 failures it carries on anyway (190) | instance | - | EW | READ, RAN (fake COM) |
| 34 | report_processor.py:195-205 -> 638-747 | `_safe_open_workbook` | COM + file | up to 3 outer tries (only for "Call was rejected by callee"; other errors re-raised at 205). Inside: file exists (643-645); open the file for reading 1 byte (648-654); reuse a workbook with the same full path (657-660); if path > 218 chars add `\\?\` or `\\?\UNC\` prefix (668-680); try 1 `Workbooks.Open(path)` (685); try 2 **close every workbook with the same file name without saving** (694-697), wait 3 s, open again (699); try 3 short path (709-716), or without prefix (719-722), or `Open(Filename, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False)` (725-731); else re-raise (747) | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 35 | report_processor.py:208-211 | `_refresh_power_query_simple` | - | sleep 5 s "to fully load" | - | - | EW | READ |
| 36 | report_processor.py:216-230 | same | COM (`WScript.Shell`) | create Shell object; `AppActivate(excel.Caption)` up to 3 times; result value ignored | whichever window Windows matches | - | EW | READ, RAN (fake COM) |
| 37 | report_processor.py:236-249 | same | keystrokes | `SendKeys("%a")` (Alt+A), `SendKeys("r")`, `SendKeys("a")` with 1 s gaps = ribbon Data > Refresh All by keyboard | whichever window has focus | - | EW | READ, RAN (fake COM) |
| 38 | report_processor.py:253-298 | same | - / Windows windows | wait `wait_time` s (Extract tab, 5-120, default 10: extraction_view.py:126-128); 5 checks for error windows 1 s apart (266-275); extended wait `max(10, wait_time)` s (282-286); final check (289-298) | desktop windows | - | EW | READ, RAN (fake COM) |
| 39 | report_processor.py:368-476 | `_find_error_window`, `_get_window_text`, `_close_error_window` | win32gui | enumerate **all** top-level windows; read their text; click OK/Close/Yes/Continue or post WM_CLOSE | any program's dialog | - | EW | READ, RAN (fake COM) |
| 40 | report_processor.py:301 -> 589-636 | `_validate_refresh` | COM | find sheet named `info` (case-insensitive); read `Info!BB2`; fail if None, blank, "0"/"0.0" or an error text | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 41 | report_processor.py:304-307 | `_refresh_power_query_simple` | COM | **`wb.Save()` always** (also after an error dialog or failed BB2 check); sleep 2 s | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 42 | report_processor.py:316-319 | same | COM | `wb.Close(SaveChanges=True)`; `wb` is **not** set to None afterwards | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 43 | report_processor.py:347-353 | same (`finally`) | COM | `wb.Close(SaveChanges=False)` again on the already-closed workbook (error swallowed) | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 44 | report_processor.py:355-360 | same (`finally`) | COM | `excel.Quit()` (errors swallowed) | instance | - | EW | READ, RAN (fake COM) |
| 45 | report_processor.py:362-366 | same (`finally`) | COM | `CoUninitialize()` | - | - | EW | READ, RAN (fake COM) |
| 46 | report_processor.py:139-141 | `_process_report` | - | sleep 3 s "for file to be released" | - | - | EW | READ |
| 47 | report_processor.py:484-489 | `_extract_data` | COM | `CoInitialize()`; `Dispatch("Excel.Application")`; `Visible=False`; `DisplayAlerts=False` | instance | - | EW | READ, RAN (fake COM) |
| 48 | report_processor.py:492-494 -> 638-747 | same | COM | open the refreshed copy through the same `_safe_open_workbook` (#34). The screen says "Opening file in read-only mode" / "File opened in read-only mode", but the open is **not** read-only | refreshed copy (or earlier refreshed copy if "Skip refresh") | yes | EW | READ, RAN (fake COM) |
| 49 | report_processor.py:507-552 | same | COM | for each configured sheet (exact, case-sensitive name), read each configured cell's `.Value` into a dictionary under both the field name and the cell address | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 50 | report_processor.py:569-587 | same (`finally`) | COM | `wb.Close(SaveChanges=False)`; `excel.Quit()`; `CoUninitialize()` (all errors swallowed) | refreshed copy | yes | EW | READ, RAN (fake COM) |
| 51 | power_query_extractor/core/data_consolidator.py:27-167 | `DataConsolidator.create_report` | openpyxl | `Workbook()` (brand-new), one sheet per `EXTRACTION_CONFIG` entry + Summary, `save` to `Level1/PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` (overwrites same name); errors re-raised (172-174) | newly created consolidated report | no (created by B) | EW | READ, RAN (E5) |
| 52 | core/file_organizer.py:273 -> utils/helpers.py:142-185 | `_organize_single_file` -> `safe_copy_file` | file copy | copy each client input file into the organised folders; compare sizes; delete the **copy** if sizes differ (177) | client input files (source, read only) -> organised copies | unknown (never opened by B) | PW | READ, RAN (E4) |
| 53 | core/file_organizer.py:259-270 -> utils/helpers.py:187-204 | `_organize_single_file` -> `create_backup` | file copy | "fresh" mode, destination already exists: copy it to `<stem>_backup_<DDMMYY_HHMM><ext>` in the same folder, then overwrite; "resume": skip; "rerun": overwrite with no backup | organised copies | unknown | PW | READ |
| 54 | core/file_parser.py:176, 195-196 | `_process_file` | file read | `stat()` and read the first 8 bytes to check the Excel signature | client input files (read only) | unknown | SW | READ |
| 55 | utils/helpers.py:249-266 | `find_excel_files` | glob | `*.xlsx`, `*.xls`, `*.xlsm` in the source folder (top level only), skipping `~$` lock files | client input files | - | SW | READ |
| 56 | utils/helpers.py:210-247 | `validate_excel_file` | file read | header check - **never called** (imported at excel_handler.py:38, unused) | - | - | - | READ |
| 57 | utils/helpers.py:473-483 | `get_short_path` | win32api | 8.3 short path (Windows only), used by #11 | - | - | PW | READ |
| 58 | gui/pyqt_main_window.py:432-479 | `_populate_extraction_from_disk` | glob, file dates | pick "latest" Level-1 folder by name (432-448); per Version folder, look for `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` (461-462) and `*_Refreshed*` with .xlsx/.xls/.xlsm (467-471); show "Refreshed (date)" from the newest refreshed file's modification time (474-479) | reports / refreshed copies (not opened) | - | Main | READ, RAN (E7 sort) |
| 59 | gui/controllers/processing_controller.py:188, 220 | `_process`, `_create_report` | - | names: summary `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx`; report always ends `.xlsx` even if the template is `.xlsm` (the file dialog allows both: setup_view.py:612, 619) | - | - | PW | READ |
| 60 | gui/controllers/extraction_controller.py:109-121 | `_save_log_file` | text file | `Level1/Extraction_Log_<YYYYMMDD_HHMM>.txt` with every on-screen log line (includes every extracted figure) | - | - | EW | READ |
| 61 | power_query_extractor/extractor_main.py:9, 36-60 | module / `load_cached_folder` | - | imports `.gui.extractor_window`, deleted in a62d3cc -> module cannot be imported | - | - | - | READ, RAN (E1) |
| 62 | report_processor.py:749-751 | `cleanup` | - | does nothing; never called | - | - | - | READ |

Not touchpoints but relevant: `pyqt_main_window.py:405-410` opens the Level-1 folder in Explorer (`os.startfile`), not Excel. No `SaveAs`, `SaveCopyAs`, `DispatchEx`, `EnsureDispatch`, `GetObject`, `RefreshAll`, `Refresh(`, `CalculateUntilAsyncQueriesDone`, `BackgroundQuery`, `Connections`, `Queries`, `DataMashup`, `.ReadOnly` (as a property read), `taskkill` or `EXCEL.EXE` anywhere in B's .py files (READ via grep; RAN grep for DataMashup/customXml: no hits).

---

## 3. Output 2 - Standard 12 "Excel through COM" check, per COM session

Standard read at `/home/user/Annual_Statement_2.0/.build-docs/toolkit/rebuild_pipeline_FINAL.md:844-878` (table rows 827-828 also apply).

COM sessions in B:
- **S1** = `core/excel_handler.py:_create_report_win32com` (95-217). Step 3, once per report (ITC, then Sales) per client. Writes the Links cells into a fresh copy of the template and saves. Thread PW.
- **S2** = `power_query_extractor/core/report_processor.py:_refresh_power_query_simple` (165-366). Step 4, once per selected report. Refreshes (by keystrokes) a fresh copy and saves. Thread EW.
- **S3** = `report_processor.py:_extract_data` (478-587). Step 4, right after S2 for the same report. Reads cells, closes without saving. Thread EW.
- **S4** = the `WScript.Shell` COM object inside S2 (report_processor.py:216). Not Excel; only R11/R12 apply.
- **S5** = `ExcelHandler.__del__` (excel_handler.py:57-63). Dead (never has an Excel object).

| Rule (Standard 12) | S1 Step 3 create report | S2 Step 4 refresh | S3 Step 4 extract |
|---|---|---|---|
| R0 COM only for PQ workbooks; before an openpyxl save into a workbook the app did not create, check for DataMashup | S1 is COM for the template copy: MET for the COM path. The sibling openpyxl fallback (excel_handler.py:219-262) saves into a template copy with **no DataMashup check**: NOT MET (fallback unreachable in HEAD, see Output 3) | MET (COM) | N/A (no save) |
| R1 Separate hidden instance via `DispatchEx`, never `Dispatch`/`GetActiveObject`, never take over the user's Excel | **NOT MET**: `GetActiveObject` first (111) = deliberately takes the user's Excel; `Dispatch` otherwise (114) | **NOT MET**: `Dispatch` (174) | **NOT MET**: `Dispatch` (487) |
| R2 `Visible = False` | MET in form (118) - but applied to the user's instance when R1 attached to it, and never set back | **NOT MET**: `Visible = True` on purpose, plus maximised (181, 183) | MET (488) |
| R3 `DisplayAlerts = False` | Set (119) but **switched back to True at 205 just before `Quit()` at 209** - NOT MET at the moment the rule protects | MET (182), but inside a retry loop that continues even if all 5 attempts failed (187-190) - PARTLY MET | MET (489) |
| R4 `Quit()` in a `finally`, even after errors | PARTLY MET: in `finally` (209), but (a) it also quits the user's instance (test at 207 always true), (b) unprotected lines 204-205 raise if Excel has died, so `Quit()` is skipped (RAN fake COM) | MET (355-360) | MET (577-582) |
| R5 Open with `IgnoreReadOnlyRecommended=True` | **NOT MET**: plain `Workbooks.Open(path)` (286, 301) | **NOT MET** on the normal path (685, 699, 714, 722); only the very last fallback passes it (725-731) | **NOT MET** (same function as S2) |
| R6 Check `workbook.ReadOnly`; if true close and ask the user | **NOT MET** (never read anywhere) | **NOT MET** | **NOT MET** (low impact: read only, closed without save) |
| R7 `BackgroundQuery = False` on each connection before `RefreshAll` | N/A in intent (no refresh). Open question: if the template's queries refresh automatically on opening, S1 saves while that refresh may be running (Q5) | **NOT MET**: connections are never touched; refresh is triggered by keystrokes (236-249), not `RefreshAll` | N/A (same open-question as S1; no save, so no file damage) |
| R8 `CalculateUntilAsyncQueriesDone` after `RefreshAll`, before saving | N/A | **NOT MET**: fixed waits only - from the last keystroke to `Save()` = `wait_time + 5 + max(10, wait_time)` seconds (default 25 s; minimum 20 s; maximum 245 s) (253-298) | N/A |
| R9 No connection still `Refreshing` before saving | N/A | **NOT MET** (never checked) | N/A |
| R10 Never use `RefreshDate` as proof of a refresh | MET in letter (not used) | MET in letter, but the substitute proof - Info!BB2 non-empty (589-636) - has the same weakness: a value can be left from an earlier refresh or the template | N/A |
| R11 `pythoncom.CoInitialize()` in worker threads; do the Excel work in a separate function; `CoUninitialize()` after that function returns | CoInitialize MET (107, worker thread). CoUninitialize **NOT MET as specified**: it is in the same function's `finally` (215), while `excel` (and loop variables `ws`, `sheet`) still hold Excel objects; and it is skipped when 204-205 raise | Same pattern: CoInitialize 171 MET; CoUninitialize 363 in the same function while `excel`, `wb`, `shell` are still held - NOT MET as specified | Same: 484 MET; 584 NOT MET as specified |
| R12 No Excel object still held afterwards | **NOT MET** (see R11); objects are released only when the function returns, after CoUninitialize | **NOT MET** (`wb` deliberately not cleared after close, 317; `excel`, `shell` held) | **NOT MET** |
| R13 Excel on every PC; logged-in user; never a scheduled job | B needs pywin32 to start at all (RAN E1) and Excel for Steps 3-4 (without Excel, `Dispatch` fails at 114 -> report fails with only "Failed to create ITC report" on screen, reason in the log file). Interactive only (no scheduler) - consistent | Needs more than a logged-in user: an **unlocked desktop, Excel in front, and nobody using keyboard or mouse** for 20-245 s per report (SendKeys). README says macOS/Linux are supported (README.md:52) - contradicted by the code | Same Excel requirement |
| R14 `Queries.Add` / `ListObjects.Add` - ask first | N/A (never used) | N/A | N/A |

S4 (`WScript.Shell`): created after CoInitialize and still referenced at CoUninitialize (363) - R12 NOT MET. S5: N/A (dead).

### 3.1 Notes behind the table
- **Does `Dispatch` attach to the user's Excel?** Standard 12 forbids `Dispatch` regardless. Whether it actually hands back the user's already-open Excel on the office PCs is not certain from reading (it depends on how Excel registered itself; commonly it starts a new instance, but not guaranteed). LOW-MEDIUM. If it does, S2 would make the user's Excel visible/maximised and then `Quit()` it with `DisplayAlerts = False`, i.e. without any "save changes?" question (S2 and S3 both quit). How to check: on a test PC open Excel with an unsaved Book1, run Step 4 on a test client, and watch whether Book1's window is reused or closed (or compare `excel.Hwnd` of the user's window with the one B gets).
- **`GetActiveObject` in S1 definitely attaches** when any Excel is running (excel_handler.py:111-112 logs "Using existing Excel instance"). RAN (fake COM, E3 user_excel_open) shows: `USER_EXCEL.Visible = False` ... `USER_EXCEL.DisplayAlerts = True`, `USER_EXCEL.Quit()` while the user's workbook is still open in it. Real-Excel outcome (MEDIUM/LOW): the user's windows disappear during Step 3; at `Quit()` Excel will ask to save the user's unsaved books - possibly in an invisible window, so B's worker could wait forever (exactly the case Standard 12 line 854-855 warns about). The Sales report that follows calls `GetActiveObject` again on a half-closed instance - outcome unknown.
- The code comment at excel_handler.py:206 says "Don't quit if Excel was already open", but nothing ever records that it was already open (`self.excel_app` stays `None`). The intent and the behaviour disagree. HIGH.

---

## 4. Output 3 - openpyxl saves, DataMashup, and what happens without pywin32

### 4.1 Every openpyxl save in B
| Save | Target | May it hold PQ? | DataMashup check? | Reachable in HEAD? |
|---|---|---|---|---|
| excel_handler.py:255 `_create_report_openpyxl` | copy of the user's ITC/Sales template at the final report path | **yes (intended)** | **none** (no `DataMashup`/`customXml` anywhere in B's code - RAN grep) | No, through the app (4.2). Yes, if `core.excel_handler` is used on its own on a non-Windows machine (E2) |
| excel_handler.py:339 `create_summary_report` | brand-new `Workbook()` | no | not needed | yes |
| data_consolidator.py:167 `create_report` | brand-new `Workbook()` | no | not needed | yes |

**What the fallback does (RAN, E2):** on a synthetic template that contains `customXml/item1.xml` (with `DataMashup`), its relationship files and `xl/connections.xml` (a "Query - Q1" Mashup connection), B's `create_report_from_template` returned **True**; the output kept the sheets and the written Links values but had **no `customXml/*` part, no DataMashup text and no `xl/connections.xml`**. Log lines produced: "Not on Windows - Power Query may not be preserved!" and "Using openpyxl - Power Query content may be lost!" - **log file / console only** (excel_handler.py:90, 224). The Processing tab would show "✅ Success: <report>.xlsx" (processing_controller.py:225-230). Caveat: my workbook was built by hand, not saved by Excel; this matches the author's own tested note (Standard 12 line 827). HIGH.

Other fallback differences from the COM path (READ): it does not fall back to a sheet containing "link" and does not create a Links sheet - if no exact Links sheet exists it just saves the copy unchanged (244-255) and still reports success; it keeps external links (`keep_links=True`, 231), which contradicts README.md:44 "No VBA macros or external links preserved"; macros in an .xlsm template would be dropped (`keep_vba` not set). The process-wide `warnings.filterwarnings('ignore')` (excel_handler.py:17) would also hide openpyxl's own "extension is not supported and will be removed" warnings.

### 4.2 What happens when pywin32 is not installed
- `requirements.txt:13-14`: "# Windows-only: Excel COM automation for Power Query" / "# pywin32>=306" - commented out, so `pip install -r requirements.txt` does **not** install it. README.md:107 instead shows `pywin32>=305; sys_platform == 'win32'` as if it were in the file. (READ)
- Import chain (READ): `main.py:38` imports `gui.pyqt_main_window` -> `gui/__init__.py:7` -> `gui/pyqt_main_window.py:25-27` -> `gui/controllers/__init__.py:4` / `pyqt_main_window.py:27` -> `gui/controllers/extraction_controller.py:10` -> `power_query_extractor/core/report_processor.py:11` `import pythoncom` (no try/except). `core/excel_handler.py:21-28` does guard its own import, so the two files follow different patterns (inconsistency).
- RAN (E1, Linux, no pywin32):
  - `python main.py` -> exit code 1; console: "Missing module: No module named 'pythoncom'" / "Please install required packages:" / "pip install PyQt5 openpyxl" (main.py:49-54); one line written to `gst_organizer.log` in the **current working folder** (main.py:20).
  - `python launch_extractor.py` -> "Failed to launch: No module named 'pythoncom'", exit 1 (launch_extractor.py:22-24); no log file (it sets up no logging).
  - Even `gui.controllers.processing_controller` cannot be imported without pywin32 (package `__init__` files pull in the extractor).
- **What the user sees** (inference): started from a console - the three lines above, then the window closes; the suggested command would not fix it. Started without a console (pythonw/shortcut, as Standard 13 requires) - nothing at all appears (print goes nowhere), only the log line in whatever folder was current. MEDIUM-HIGH.
- **Is there a silent fallback that drops queries?** In B's HEAD app: **no**, because the app cannot start without pywin32, so `USE_WIN32COM` can only be False where the extractor import has already failed. HIGH. (Only exotic case: `win32com.client` importing but `pythoncom` not, or the reverse - both files need both, so still no.) But the fallback code is still present and becomes live the moment the extractor import is made optional (as README.md:52-55 implies is intended) - then queries would be dropped silently with an on-screen "Success". Merge risk, HIGH.
- **History** (READ): at the split point 47a1492 the organiser (Tkinter) never imported the extractor (`git grep` at 47a1492: only `launch_extractor.py:14` and `power_query_extractor/gui/extractor_window.py:12` reference it), so on a Windows PC without pywin32 the organiser started and **took the openpyxl fallback** (dropping PQ) with only a log warning. The unified window in **c258047 (07-02-2026, B only)** made the extractor import mandatory, which closed that path. The requirements line "# pywin32>=306" was added (commented) in the same commit.

---

## 5. Output 4 - Two realistic runs traced, every Excel call in order

### 5.1 Run (a): Step 3 creates the ITC and Sales reports for <Client A>
Set-up assumed: Windows, pywin32 and Excel installed, Excel **not** already open, source folder holds `GSTR3B-<ClientA>-<State>-Apr.xlsx`, `GSTR3B-...-May.xlsx`, `GSTR-2B-Reco-...xlsx`, `Sales-...xlsx`, `AnnualReport-...xlsx` (no IMS Reco, no Sales Reco); mode "Fresh Run". RAN end-to-end as E4 with a placeholder client (fake COM for Excel, real file copies and real openpyxl summary).

1. User clicks Start Processing (validation_view.py:217 or processing_view.py:95) -> `pyqt_main_window._on_start_processing` (345-371): needs a selected client, both template paths existing (353-358) and a target folder (359-361). No check of the template's type, Links sheet or queries.
2. `ProcessingController.start_processing` (processing_controller.py:256-279) refuses if its own thread is running (259-261) - **not** if Step 4 is running. New QThread; `ProcessingWorker.run` -> `_process` (42-202).
3. `FileOrganizer(...)` (46-52) creates the target folder if missing (file_organizer.py:61-65). `ExcelHandler()` (54) creates `%TEMP%\gst_excel_temp` (excel_handler.py:52-53).
4. Per client (80-181): `create_client_structure` makes `Annual Statement-<DDMMYY HHMM>\<Client-ST>\Version-<DDMMYY HHMM>\` and the three category folders (file_organizer.py:67-144). `organize_files` copies each input file (`shutil.copy2`, helpers.py:169) and checks sizes (172-178); originals untouched.
5. **ITC report** - `_create_report` (processing_controller.py:204-239): template path empty/missing -> error line and stop (210-218); output `Version-...\ITC_Report_<client>_<ST>_<DDMMYY_HHMM>.xlsx` (220); `prepare_template_data` (excel_handler.py:540-627) builds: B2 = GSTR-3B folder, B4 = Version folder, B5 = annual file name without extension, B7 = "Other ITC related files" folder, B8 = 2B file name, B10/B11 = "" (no IMS file).
6. `create_report_from_template` (65-93): template exists (78-79); `USE_WIN32COM` True -> **S1** (95-217). Excel calls in order (RAN fake COM, E4):
   1. `pythoncom.CoInitialize()` (107)
   2. `GetActiveObject("Excel.Application")` -> fails (no Excel) (111)
   3. `Dispatch("Excel.Application")` -> instance APP1 (114)
   4. `APP1.Visible = False`, `APP1.DisplayAlerts = False`, `APP1.ScreenUpdating = False` (118-120)
   5. `shutil.copy2(template, report)`; sleep 0.5 s (124-125)
   6. iterate `APP1.Workbooks` (none); `APP1.Workbooks.Open(report)` (274-286)
   7. iterate `wb.Sheets` for "Links" (137-140)
   8. write `Links!B2`, `B4`, `B5`, `B7`, `B8` (B10, B11 skipped because empty - the template's own B10/B11 values stay) (159-167)
   9. `wb.Save()` (173); `wb.Close(SaveChanges=True)` (176); sleep 0.5 s (180)
   10. file-exists check -> return True (183-186)
   11. `finally`: `APP1.ScreenUpdating = True`, `APP1.DisplayAlerts = True` (204-205); `APP1.Quit()` (209); `pythoncom.CoUninitialize()` (215)
7. Screen (RAN): "📊 Creating ITC report..." / "Output: ITC_Report_..." / "✅ Success: ITC_Report_....xlsx".
8. **Sales report**: same 11 steps with a new instance APP2; cells B2 = Sales folder, B3 = Sales file name, B5 = Version folder, B6 = annual name; B8/B9 skipped (no Sales Reco). (RAN)
9. `create_organization_report` writes `_Organization_Report_<client>_<state>.txt` (file_organizer.py:300-362).
10. After all clients: `create_summary_report` -> `Workbook()` -> save `Annual Statement-...\GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (excel_handler.py:308-346) - openpyxl, no Excel. Screen "✅ Summary report created" only on success; on failure nothing (RAN E4 block_summary).
11. `finished` -> `_on_processing_completed` (pyqt_main_window.py:373-410): run history saved, Extract tab re-filled from disk, message box, Explorer opens the Level-1 folder.

Totals per client: 2 Excel starts and 2 quits, 2 template copies, 4 saves through COM (Save + Close(SaveChanges=True) twice), nothing refreshed in Step 3.

**Variants of run (a):**
- **Excel already open (user working in it):** step 6.2 returns the user's Excel; 6.4 hides it and freezes its screen; the report is opened and saved inside the user's Excel; 6.11 re-enables alerts and quits the user's Excel (RAN fake COM, E3). Real outcome: see 3.1 (MEDIUM/LOW).
- **Excel dies after the copy (6.5) but before or during the open (6.6):** error caught (191-193, "Error creating report with win32com" in the log file), then `finally` line 204 raises (Excel gone), so `Quit()` and `CoUninitialize()` are skipped and the error escapes to `_create_report` (processing_controller.py:235-239): screen "❌ ITC Report Error: <COM error text>". The unfilled template copy stays at the report's normal name (RAN fake COM, E3b: "output file left on disk: True"). Step 4 will later list it as an ITC report and refresh it.
- **Excel not installed:** 6.3 raises; caught (191-193); screen "❌ Failed to create ITC report" with no reason (reason in the log file). READ HIGH.
- **Template has no Links sheet:** a new empty sheet named Links is added and filled; report "Success" (152-153; log-file error only). READ HIGH.
- **Template is .xlsm:** copied to a `.xlsx` name (processing_controller.py:220); Excel normally refuses a macro workbook with an .xlsx extension (MEDIUM) -> "Failed to create" for every client. **Template is .xltx** (README.md:120 suggests it): same extension mismatch (MEDIUM).
- **Report path > 218 characters:** retried with the 8.3 short path (292-301); on a network share 8.3 names usually do not exist, so the open fails (MEDIUM).
- **Same minute re-run:** same folder names and same report name -> `shutil.copy2` silently overwrites the earlier report; if that report is open in Excel the copy fails (Windows lock) -> "Failed to create" (READ HIGH / lock behaviour MEDIUM).

### 5.2 Run (b): Step 4 "Extract & Refresh" over several clients
Set-up assumed: three versions ticked for ITC and Sales: <Client A>, <Client B>, <Client C>. Wait time 10 s, suffix `_Refreshed_{timestamp}`, Skip refresh off. <Client B>'s refresh hits a Power Query error; Excel crashes during <Client C>'s Sales save; the user has <Client A>'s ITC report open in their own Excel.

0. Which folder: the Extract tab was filled from the Level-1 folder with the highest **name** (pyqt_main_window.py:432-448) - across months this can be an older batch (RAN E7). The tree shows "Refreshed (date)" if any `*_Refreshed*` file exists (467-479) - a file date, not a refresh proof.
1. Start Extraction -> `_on_start_extraction` (506-547) builds one entry per ticked version: name = "<Level-2 folder> / <Version folder>", path, `process_itc`, `process_sales` (extraction_view.py:368-390). `ExtractionController.start_extraction` (138-165) refuses only if Step 4 itself is running. `ExtractionWorker._process` (48-107) sets `EXTRACTION_OPTIONS['skip_refresh']` in the shared configuration dictionary (54-55).
2. Per client, per report (`process_client` -> `_process_report`, report_processor.py:39-163). Excel calls in order for **one report** (RAN fake COM, E5 normal):
   1. glob `ITC_Report_*.xlsx`, take the first (76-89)
   2. `shutil.copy2(report, report_Refreshed_<DDMMYY_HHMM>.xlsx)` (120)
   3. **S2**: `CoInitialize()` (171); `Dispatch("Excel.Application")` -> APP1 (174); `APP1.Visible = True`, `DisplayAlerts = False`, `WindowState = maximised` (181-183)
   4. open the copy: exists, 1-byte read, same-path check, `APP1.Workbooks.Open(copy)` (195-205, 638-689)
   5. sleep 5 s (208-211)
   6. `Dispatch("WScript.Shell")`; `AppActivate("Excel")` (216-224); sleep 1 s
   7. `SendKeys("%a")`, 1 s, `SendKeys("r")`, 1 s, `SendKeys("a")` (236-249)
   8. wait 10 s (253-258); error-window scans x5 over 5 s (261-278); wait 10 s (281-286); final scan (289-298)
   9. read `Info!BB2` (301 -> 589-632)
   10. `wb.Save()` (305); sleep 2 s; `wb.Close(SaveChanges=True)` (317); sleep 1 s
   11. `finally`: `wb.Close(SaveChanges=False)` on the closed book -> error swallowed (348-353); `APP1.Quit()` (357); `CoUninitialize()` (363)
   12. sleep 3 s (141)
   13. **S3**: `CoInitialize()`; `Dispatch` -> APP2; `Visible = False`; `DisplayAlerts = False` (484-489); open the copy (493; not read-only despite the message); read 17 ITC cells (or 5 Sales cells) on 7 (or 1) sheets (507-552); `Close(SaveChanges=False)`; `Quit()`; `CoUninitialize()` (569-587)
   Elapsed per report with defaults: roughly 5 + 3 + 25 + 3 + 3 = about 40 s plus Excel start-up twice. Two Excel starts per report, four per client.
3. **<Client A>, ITC report already open in the user's Excel:** step 2.2 reads the open file (Windows normally allows reading a workbook Excel has open - MEDIUM), so the copy holds the last **saved** version. B's own Excel opens the copy (different name - no clash). At 2.6 `AppActivate("Excel")`: Windows looks for a window whose title is, begins with, or ends with "Excel"; both the user's window ("<book> - Excel") and B's qualify, so either may be brought forward (MEDIUM). If it is the user's, Alt+A, R, A refreshes **the user's open workbook** (unsaved change to their file) and B's copy stays unrefreshed; B then checks BB2 on its copy: if empty -> "⚠️ Refresh completed with errors: Refresh validation failed: Cell BB2 is empty" and FAIL; if BB2 already had a value -> reported as success with unrefreshed figures (MEDIUM). If the user had the **refreshed copy** open and "Skip refresh" is on, S3 opens it in a second Excel; with alerts off Excel normally opens it read-only (MEDIUM), reads the saved values, closes without saving - no damage. If the copy name of this minute already exists and is open, step 2.2 fails (lock) -> report error "FAIL".
4. **<Client B>, refresh fails (query error):** Excel shows an error box. If its title/class match (4 of section 1), B logs "⚠️ Error detected: ..." in the log file, clicks its OK/Close/Yes/Continue button (RAN fake COM, E5 error_dialog), keeps waiting, checks BB2, **saves the copy anyway** (305), then shows "⚠️ Refresh completed with errors: <dialog text>". Extraction still reads the cells; the consolidated report shows ✗ and "Missing" for every column of that report type (data_consolidator.py:116-117) (RAN). If the failure is silent (no dialog, e.g. a query that returns no rows) and BB2 is filled, B reports success with whatever the cells hold (MEDIUM). If a Power Query prompt (credentials, privacy levels) stays open, later COM calls are refused by Excel ("Call was rejected by callee") or wait; `Save()` fails -> "Error during refresh" (log file only) -> FAIL; `Quit()` may also be refused, leaving a visible Excel open on the user's screen (MEDIUM). If B's keystrokes landed in a cell instead of the ribbon, Excel stays in cell-edit mode with the same effect (MEDIUM).
5. **<Client C>, Excel crashes during the Sales save:** `wb.Save()` raises (305) -> inner `except` (332-337) logs "❌ Error during refresh" (log file only) and returns FAIL; `finally` swallows the failures of `Close`/`Quit`, then `CoUninitialize` (RAN fake COM, E5 crash_on_save). On screen the user sees "Saving workbook..." then "Cleaning up Excel instance..." - no error text - and later "Result: ... Sales: FAIL" (extraction_controller.py:77-84). Extraction then starts a **new** Excel and reads the copy as last saved (unrefreshed, or damaged if the crash happened mid-write: Excel may then refuse to open it, giving `{}` and "Missing"). A crashed EXCEL.EXE or its "stopped working" box may remain; B never kills Excel. Excel **hanging** instead of crashing would block the worker for ever: no COM timeout, Stop is honoured only between clients (61-63).
6. **Open refused 3 times ("Call was rejected by callee"):** `_safe_open_workbook` makes up to 4 opens per try, 3 tries = 12 opens; then `wb` stays `None` and the code **carries on**: 5 s load wait, AppActivate and the three keystrokes are still sent - into whatever window has focus - then BB2 check fails, `Save()` on nothing fails, "Error during refresh: 'NoneType' object has no attribute 'Save'" (log), FAIL (RAN fake COM, E5 open_rejected NFAIL=12). Extraction then also fails to open -> `{}` -> "Missing".
7. After all clients: `DataConsolidator.create_report` -> `Level1\PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` (openpyxl, new file) (data_consolidator.py:20-174) - includes finding 6 (wrong IMS/Extra columns) and 7 ("Missing"). Then `Level1\Extraction_Log_<YYYYMMDD_HHMM>.txt` (extraction_controller.py:109-121). `finished` -> run history updated under the key "<Level-2 folder> / <Version folder>" (pyqt_main_window.py:555-558), which is not the key used by Step 3's records (`<client>-<state code>`, processing_controller.py:150-160 via file_parser.py:262), so `add_extraction_result` (run_history.py:73-95) creates a new record instead of updating the Step 3 one (write/read key mismatch). READ HIGH.
8. Closing the window during extraction: `closeEvent` (pyqt_main_window.py:601-608) saves settings and closes without stopping the worker; Qt normally aborts a program whose QThread is still running (MEDIUM), leaving B's visible Excel and the half-refreshed copy behind.

---

## 6. Output 5 - Other file-safety points

### 6.1 Are originals ever modified, moved or deleted?
- **Client input files (source folder): never modified, moved, renamed or deleted.** Only `stat()`, an 8-byte read (file_parser.py:176, 195-196) and use as the source of `shutil.copy2` (helpers.py:169). Grep: no `os.remove`, `rename`, `shutil.move`, `os.replace` anywhere; the only `unlink` (helpers.py:177) deletes a **destination copy** whose size differs; the only `rmtree` (file_organizer.py:375) is in `cleanup_on_error`, which is never called. READ HIGH.
- **Templates: never modified.** Only `exists()` and copy source (excel_handler.py:78, 124, 228). READ HIGH.
- **Organised copies:** "rerun" mode overwrites a same-named copy without backup; "fresh" mode backs it up first (file_organizer.py:259-270). Only possible when the same Version folder is reused (same minute). READ HIGH.
- **Step 3 reports:** created and then modified in place by S1 (B's own files); overwritten without backup on a same-minute re-run (excel_handler.py:124). READ HIGH.
- **Step 4:** the Step 3 report itself is never opened for writing (a copy is refreshed). Exceptions: an empty suffix makes the copy name equal the original - `shutil.copy2` then refuses (`SameFileError`) so the original is still safe (RAN E5); a refreshed copy of the same name (same minute, or a suffix without `{timestamp}`) is overwritten without warning (READ HIGH).
- **The user's own open workbooks** can be affected through Excel, not through files: S1 quits the user's Excel (finding 3); S2's keystrokes and dialog-clicking can act in the user's Excel or other programs (findings 2, 4); `_safe_open_workbook` "second attempt" closes any workbook with the same **file name** in its Excel instance without saving (report_processor.py:692-699) - harmless in a separate instance, harmful if `Dispatch` ever shares the user's instance (LOW-MEDIUM).

### 6.2 Temporary files
- `%TEMP%\gst_excel_temp` is created on every Step 3 run and never used or deleted (excel_handler.py:52-53). README.md:7, 42, 260 ("All Excel operations use temporary files") is **not true**: no file is written to a temporary name and then renamed; the real report path is used from the start. READ HIGH.
- Excel's own `~$...` lock files appear next to opened workbooks and stay after a crash; B's scanners skip them (helpers.py:259-260; pyqt_main_window.py:470) and the report pattern cannot match them. READ HIGH.
- `EXCEL_SAFETY` (`temp_prefix`, `backup_prefix`, `verification_size_min` 10 KB, `max_retries`, `save_options`) in utils/constants.py:217-228 is stored at excel_handler.py:54 and **never used**. READ HIGH.

### 6.3 Backups
- Only `create_backup` for organised input copies in "fresh" mode (helpers.py:187-204), placed in the **same folder** as `<stem>_backup_<DDMMYY_HHMM><ext>`. Reports, refreshed copies, summary and consolidated reports are never backed up. README.md:47, 306-307 ("Original files are never modified", "Backups created for existing files") are true only in this narrow sense.
- Possible side effect (LOW-MEDIUM, depends on the templates): B writes **folder** paths into the Links sheet (e.g. B2 = the GSTR-3B folder). If a template query reads every file in that folder, a `_backup_` copy sitting in it would be read too. Q6.

### 6.4 Network drop or full disk mid-copy
- `safe_copy_file` (helpers.py:142-185): if `shutil.copy2` raises part-way, the `except` (183-185) returns False **without deleting the partial file**; the size check (172-178) runs only when the copy finished. The partial file keeps the real name in the organised folder. The Links cells are still written for that file type (they depend on the scan, not on copy success: excel_handler.py:569-577 uses `client_info['files']`), so the report's queries would point at a partial file. READ HIGH (consequence MEDIUM).
- Template copy in S1 (excel_handler.py:124) and in the fallback (228): a partial report file is left at the final name on failure (no clean-up). Step 4 will then find it by pattern. READ HIGH.
- Step 4 copy (report_processor.py:120): partial refreshed copy left; the error is caught at 154-163 (FAIL). READ HIGH.
- Excel's own save to a network share (S1 173/176, S2 305/317) uses Excel's internal safe-save; how it behaves on a dropped connection is Excel's, not B's (not checked). LOW.
- No retry, no lock file, nothing stops two people running Step 3/4 on the same target folder at the same time. READ HIGH.

### 6.5 Where figures and paths are written outside the reports (confidentiality)
- `gst_organizer.log` in the folder the app was started from (main.py:16-23; level INFO): receives every `logger.info` line, including `report_processor._log` lines (report_processor.py:33) such as "✓ <field> (<cell>) = <value>" (541) and the "Extraction summary" values (557-562), and the Links paths/file names written in Step 3 (excel_handler.py:165). Under Standard 13 this would be the shared app folder (the other audits cover Standard 13).
- `Level1\Extraction_Log_<YYYYMMDD_HHMM>.txt` (extraction_controller.py:109-121): every on-screen line, including all extracted values.
- `~/.gst_organizer/run_history.json` (models/cache_manager.py:17, 36; models/run_history.py:128-136): the full extraction result per client (status, every extracted value, original and refreshed file paths) via `add_extraction_result` (run_history.py:73-95).
- Note: the logs carry client names inside file and folder names. READ HIGH.

### 6.6 Other Excel-related points
- **Silent wrong figure columns** in the consolidated report (finding 6). Mechanism: `CELL_MAPPINGS` field names are "<output sheet>_<column>" with spaces turned into underscores (cell_mappings.py:276); the consolidator first looks for any field name that **contains** the column title (data_consolidator.py:130-133). Titles with a space ("Diff IGST") never match an underscore name and fall through to the correct exact key (136-138); one-word titles ("IGST", "CGST", "SGST") match the first ITC field containing them, which is "GSTR-3B_Difference_Diff_IGST" etc. Affected: "IMS Reco", "Extra In Purchase", "Extra In GSTR2B" (all three columns each). If the GSTR-3B Difference value is empty (None), the loop takes that None and the exact key then rescues the right value - so the error shows only when the GSTR-3B difference cell has a value. RAN (E5). HIGH. [GST RULE] context.
- `extracted_data` stores each value twice (field name and cell address, report_processor.py:536-539), so "Total values extracted: 34" counts 17 cells twice. READ, RAN.
- Number display: numbers other than 0 get format `#,##0.00` (data_consolidator.py:149-150) - display only; 0 is left unformatted. [GST RULE] (presentation of figures; described only).
- Dead code: `_write_section` (data_consolidator.py:237-339) and `_create_details_sheet` (341-399) - the only places that would show error/validation text in the consolidated report - are never called; `EXTRACTION_OPTIONS['auto_refresh_if_missing']` and `['show_timestamp']` (cell_mappings.py:250-251) are never read; `ReportProcessor.excel_instance` (24-25) and `cleanup()` (749-751) unused; the `except` at report_processor.py:125-137 can never run because `_refresh_power_query_simple` catches every error itself (339-344). READ HIGH.
- `prepare_template_data` computes `state_code` (excel_handler.py:550) and never uses it. When there are several files of one type (e.g. two 2B files), only the first in scan order is written to the Links sheet (561-563, 572-574, 583-585, 595-597, 606-608, 617-619); scan order is sorted by path (helpers.py:262). READ HIGH.
- The template-browse dialog offers `.xlsx` and `.xlsm` (setup_view.py:612, 619); `TEMPLATE_EXTENSIONS` in constants (`.xlsx`, `.xltx`, `.xltm`, utils/constants.py:180) is never used; README says `.xlsx` or `.xltx` (README.md:120-121). Three different answers. READ HIGH.

---

## 7. GST rules met in the Excel code (described only, not judged)
- [GST RULE] TIME-VARYING - **Which report cells feed the consolidated report** (power_query_extractor/config/cell_mappings.py:33-242): ITC report - "Diff GSTR-3B" J15/K15/L15 -> "GSTR-3B Difference" Diff IGST/CGST/SGST; "Reco" C37/D37/E37 -> "Reco Diff" IGST/CGST/SGST Diff; "T6 of R9" K29 -> "Table 6J Match"; "T8 of R9" G6 -> "Table 8D Limits"; "IMS_Reco" AA1/AD1/AG1 -> "IMS Reco" IGST/CGST/SGST; "ExtraInPurchase" O1/P1/Q1 and "ExtraInGSTR2B" S1/T1/U1 -> IGST/CGST/SGST. Sales report - "Diff - GSTR-1" N1/O1/P1/Q1 -> "GSTR-1 Difference" Diff Taxable Value/IGST/CGST/SGST; I1 -> "Turnover GSTR-1" Taxable Value. Sheet/table names (GSTR-9 Tables 6J, 8D) are time-varying. Added in 47a1492 (30-10-2025, shared). The calculation feature (subtract/add/multiply/divide, data_consolidator.py:176-206) is unused by the current configuration (all `calculation: None`).
- [GST RULE] - **Which return's files feed which template cell** (utils/constants.py:96-120, excel_handler.py:540-627): ITC Links B2 GSTR-3B folder, B4/B5 Annual Report folder/name, B7/B8 GSTR-2B Reco folder/name, B10/B11 IMS Reco folder/name; Sales Links B2/B3 Sales folder/name, B5/B6 Annual Report, B8/B9 Sales Reco. (Template contract; tagged because it decides which return data the report's figures come from.)
- [GST RULE] - **Refresh validity cell** Info!BB2 = "client name" (report_processor.py:605-632) - technical check, but its location belongs to the author's template.
- [GST RULE] TIME-VARYING - **State codes** used in report names (processing_controller.py:113 -> utils/helpers.py:485-511 with utils/constants.py:234-275): two-letter abbreviations (not the numeric GSTIN state codes); unknown state -> first letters of the first two words, or first three characters. The list itself (e.g. merged/renamed states and union territories) can change over time.
- [GST RULE] - GSTR-3B may have many files per client; other types flagged as duplicates when more than one (file_parser.py:290-299). Outside Excel scope, noted because only the first file of a type reaches the Links sheet (6.6).

---

## 8. README and comments versus code (Excel-related discrepancies)
| Claim | Where | What the code does | Evidence |
|---|---|---|---|
| "All file operations use temporary files and verification steps" | README.md:7, 42-43, 260-261, 310 | No temporary files are used; verification = file exists (excel_handler.py:183-189) or size equal (helpers.py:172-178) | READ |
| "Clean saves: No VBA macros or external links preserved" | README.md:44, 262 | COM path keeps everything; openpyxl fallback keeps links (`keep_links=True`, excel_handler.py:231) | READ |
| pywin32 is in requirements.txt | README.md:105-108 | Commented out (requirements.txt:14) | READ |
| macOS / Linux supported | README.md:52 | App cannot start without pywin32 (RAN E1) | RAN |
| `python launch_extractor.py` opens the extractor with a "Refresh & Extract" button | README.md:164-176 | Opens the main window on tab 4; button is "Start Extraction" (launch_extractor.py:11-20; extraction_view.py:254) | READ |
| Report names `..._DDMMYYYY_HHMMSS.xlsx` | README.md:188-189, 195, 202-203 | `DDMMYY_HHMM` (helpers.py:23-25; processing_controller.py:114) | READ |
| "Windows path limit of 260 characters" | README.md:317 | Code treats > 218 as long (excel_handler.py:293; report_processor.py:668) | READ |
| "Close all Excel files before processing" | README.md:301, 314 | Not enforced; Step 3 actively uses an open Excel (excel_handler.py:111) | READ |
| "Don't quit if Excel was already open" | excel_handler.py:206 (comment) | Always quits (207: `self.excel_app` is always None) | READ, RAN (fake COM) |
| "Opening file in read-only mode" / "File opened in read-only mode" | report_processor.py:492, 494 | Normal read-write open | READ, RAN (fake COM) |
| "Create report from template preserving Power Query" | excel_handler.py:70 | True only for the COM path | READ, RAN (E2) |
| README describes v3.0 and was last changed 29-10-2025 (d674386), before B's own commits | README.md:1, 361-364 | B is "v4.0" (main.py:2, 42; pyqt_main_window.py:71) | READ |

---

## 9. History of the Excel code (B/main)
- **edac6a3 (31-08-2025, shared initial commit):** every COM pattern already present - `GetActiveObject` then `Dispatch` with the always-true quit test (excel_handler.py then lines 122-125, 218); `Dispatch` + `Visible = True` + SendKeys refresh (report_processor.py then 154-228); BB2 validation; `IgnoreReadOnlyRecommended` only in the last fallback (then 713). (`git grep`, READ.)
- **bd176bf (30-10-2025, shared):** fixed 10 s waits replaced by the configurable `wait_time` and `max(10, wait_time)`; custom suffix; ITC/Sales selection (report_processor.py). A comment "WAIT 7 SECONDS AS REQUESTED" sat over a 10-second loop before this commit.
- **47a1492 (30-10-2025, shared, split point):** added the ITC mappings for Reco Diff, Table 6J, Table 8D, IMS Reco, Extra In Purchase, Extra In GSTR2B (cell_mappings.py) - the one-word column titles that trigger finding 6 arrived here.
- **c258047 (07-02-2026, B only):** clean-up of the Excel code: bare `except:` -> `except Exception:`, debug lines removed, unreachable `return {'success': True}` after the inner try removed, stray `print` removed; no change to how Excel is driven. Unified window made the extractor (and pywin32) mandatory at start-up; requirements.txt gained PyQt5 and the commented pywin32 line. Commit message says it was co-authored with an AI assistant.
- **a62d3cc (07-02-2026, B only):** deleted the Tkinter extractor window (`power_query_extractor/gui/extractor_window.py`), leaving `extractor_main.py` broken (RAN E1).
- **7e09e5c (08-02-2026, B only):** `find_excel_files` no longer validates each file (moved to `file_parser._process_file`) and skips `~$` lock files.
- **b0e7d0e (09-02-2026, B only):** consolidated report and new log file now go to the Level-1 folder (`level1_folder`) instead of the target folder.

---

## 10. Possible client names
None found in the Excel-related code, comments, configuration or the B-only commit messages I read. README.md:128-133 uses "ABC Ltd" / "Maharashtra" as obviously generic examples (not recorded as a client). I did not open `gst_organizer.log` (it is not in b0e7d0e; only in edac6a3 and 6237597). Commit author identity is the author's own (goyal-sidhant).

---

## 11. Questions for the author
1. Do both templates (ITC and Sales) contain Power Query queries, and are any of them set to "Refresh data when opening the file" or "Enable background refresh"? (Decides whether Step 3 saves while a refresh-on-open is running, and how long Step 4's fixed wait must be.)
2. In real use, did you ever have Excel open while running Step 3 (Processing)? Did your Excel windows vanish, or did Excel close or ask to save? (Code: excel_handler.py:111, 118, 205-209.)
3. Was the keystroke refresh (Alt+A, R, A with Excel visible) chosen because `RefreshAll` through COM did not work for these templates? If yes, what went wrong with it? (The code comment "Keep it visible to avoid issues", report_processor.py:181, is the only clue.)
4. Is Info!BB2 always empty in the saved templates and filled only by a query? (If the template was saved with a value there, the refresh check always passes.)
5. Which Excel version and language do the office PCs use? (The Alt+A, R, A KeyTips and the window title "Excel" that `AppActivate` looks for depend on both.)
6. Does any template query read **every file in a folder** (for example all GSTR-3B files in the "GSTR-3B Exports" folder)? If so, `_backup_` copies or partial copies left in those folders would be read too.
7. When a client has no IMS Reco or Sales Reco file, what should the Links cell hold? Today B leaves whatever the template already has in that cell (excel_handler.py:162).
8. Are the templates ever `.xlsm` or `.xltx`? (Output is always named `.xlsx`.)
9. In the consolidated Extract report, did the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets ever show the same numbers as "GSTR-3B Difference"? (Finding 6.)
10. Should a failed refresh still show the values that were read (marked as unrefreshed), or "Missing" as now?
11. Are the client folders on a local disk, a network share (NAS/Windows server), or OneDrive? (Affects file order for the `ITC_Report_*` pattern, long-path handling and partial copies.)
12. Is it acceptable that extracted figures are also written to `gst_organizer.log`, the Extraction log text file and `run_history.json` in the user's home folder?

---

## 12. INCOMPLETE
- Real Excel behaviour was not tested (no Windows/Excel here). Every COM call order above is from reading plus fake-COM runs; real reactions (does `Dispatch` reuse the user's Excel; what `Quit()` does with the user's unsaved books and alerts on; what `AppActivate("Excel")` picks; whether SendKeys reaches a locked or busy desktop; read-only behaviour when a file is open elsewhere; Excel's handling of `\\?\` paths) need a Windows test PC. Each is marked MEDIUM/LOW with a way to check.
- The actual ITC/Sales templates are not in the repository, so "can hold PQ = yes (intended)" and the template-dependent questions (Q1, Q4, Q6, Q7) are unverified.
- Not read line by line (grep only, no Excel/file calls found): `gui/views/setup_view.py` lines 1-199 and 241-599, `gui/views/validation_view.py`, `gui/views/processing_view.py` (except signal lines), `gui/theme/theme_manager.py`, `gui/widgets/collapsible_group.py`, `gui/widgets/colored_section.py`.
- Git history: only the Excel-relevant parts of B's commits were read in full (c258047 Excel files, bd176bf report_processor, b0e7d0e extraction_controller, 7e09e5c processing_controller/helpers, c258047 file_organizer); other files in those commits and the deleted Tkinter handlers (gui/handlers/processing_handler.py, power_query_extractor/gui/extractor_window.py) were only searched with grep.
