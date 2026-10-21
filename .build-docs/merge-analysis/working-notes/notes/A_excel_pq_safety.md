# A-X-EXCEL — Project A: Excel and Power Query safety (working notes)

Project A = goyal-sidhant/AnnualStatement_oldv1, snapshot HEAD `bf2f1dd` (27-07-2026), Tkinter GUI.
All paths below are repository-relative. Line numbers are those of HEAD `bf2f1dd`.
Written for later agents who will turn these notes into the formal audit (Project_Audit_FINAL.md method).

## 0. How to read these notes

- **READ** = I read the code/diff myself. **RAN** = I executed something (scripts are in
  `/root/merge-scratch/run/A-X-EXCEL/exp/`, outputs quoted in section 10).
- Inferences carry **HIGH / MEDIUM / LOW** confidence.
- **[GST RULE]** marks anything that decides a tax figure, which return/table/sheet/cell a figure
  belongs to, rates, rounding, periods, GSTIN/PAN/state-code logic, ITC, due dates. For these I
  only describe what the code does. **TIME-VARYING** marks values that change over time.
- This Linux machine has no Excel and no pywin32, so every Excel-through-COM behaviour is traced by
  READING only. Where the real Windows/Excel behaviour matters and I am not sure, I say so and give
  a way to check.

### Plain-English glossary (terms used below)

- **Excel workbook / template**: `.xlsx` is a normal workbook; `.xltx` is a template (Excel opens a
  new copy of it); `.xlsm`/`.xltm` are the macro-enabled versions. All four are ZIP files of XML "parts".
- **Power Query (PQ)**: Excel's built-in data-import tool. Its queries (the "M code") are stored inside
  the workbook ZIP in a part `customXml/item*.xml` containing an element called **DataMashup**, plus
  connection records in `xl/connections.xml`.
- **openpyxl**: a Python library that reads and writes `.xlsx` files without Excel. When it SAVES a
  workbook it rebuilds the ZIP from what it understands, so parts it does not understand (Power Query's
  DataMashup, connections) are dropped. Reading is harmless.
- **COM**: Windows' way for one program to drive another. "Excel through COM" means Python starts or
  attaches to a real Excel and gives it commands. **pywin32** is the Python package that provides COM
  (`win32com.client`, `pythoncom`) and Windows window functions (`win32gui`, `win32api`, `win32con`).
- **Dispatch / DispatchEx / GetActiveObject**: three ways to get an Excel. `GetActiveObject` grabs an
  Excel that is ALREADY running (normally the user's own). `Dispatch` creates one but, for Excel,
  commonly connects to an already-running Excel if there is one. `DispatchEx` always starts a new,
  separate Excel process. The author's Standard 12 requires `DispatchEx`.
- **Visible / DisplayAlerts / ScreenUpdating**: Excel settings: whether its window is shown, whether it
  shows prompts ("Save changes?"), whether it repaints the screen.
- **Quit()**: tells that Excel to close completely (all its workbooks).
- **RefreshAll / BackgroundQuery / CalculateUntilAsyncQueriesDone / Refreshing / RefreshDate**: the COM
  calls/properties Standard 12 names for refreshing Power Query safely: switch background refresh off
  on each connection, call RefreshAll, wait with CalculateUntilAsyncQueriesDone, check no connection is
  still "Refreshing", and never trust RefreshDate.
- **SendKeys**: pressing keyboard keys by program, into whatever window has keyboard focus.
- **Ribbon key tips**: Excel's Alt-key shortcuts (Alt, A = Data tab; R, A = Refresh All).
- **Worker thread**: a second line of execution inside the app so the window stays responsive.
  **Daemon thread**: a worker thread that is killed instantly, without clean-up, when the app exits.
- **CoInitialize / CoUninitialize**: calls a worker thread must make before / after using COM.
- **Staging file**: a temporary local copy that is worked on and then copied to the final place.
- **UNC path**: a network path like `\\server\share\folder`. **8.3 short path**: old short DOS-style
  name of a long path. **`\\?\` extended path**: Windows prefix to allow very long paths.
- **find_spec**: a Python check "does a module file exist?" that does not actually import (load) it.
- **glob**: listing files that match a wildcard pattern like `ITC_Report_*.xlsx`.
- **mtime**: a file's "last modified" time.
- **Content type**: the label inside a ZIP-based Office file that says "this is a workbook" or "this is
  a template".
- **Orphan / zombie EXCEL.EXE**: an Excel process left running in the background after the program
  that started it has stopped using it.

---

## 1. Most important findings (for the author's top priority)

1. **Step 3 report creation grabs the user's own open Excel, hides it, and then closes it.**
   `core/excel_handler.py:123-128` tries `win32com.client.GetActiveObject("Excel.Application")` first
   (the user's running Excel), and only if that fails calls `Dispatch` (not `DispatchEx`). It then sets
   `Visible = False` (131) on that Excel and never sets it back (finally block 238-246 restores only
   ScreenUpdating and DisplayAlerts). The comment at 241 says "Don't quit if Excel was already open", but
   the test is `if self.excel_app is None` (242) and `self.excel_app` is set only once, to `None`, at 56
   and never assigned anywhere else (READ, grep over all .py files). So `excel.Quit()` (244) is ALWAYS
   called — on the user's own Excel too. Just before quitting, line 240 switches `DisplayAlerts` back to
   `True`, so if the user has unsaved work, Quit runs with prompts on in an Excel that is now invisible.
   READ (code path) HIGH; real-world effect (user's Excel windows vanish, then Excel closes or hangs on an
   invisible "Save changes?" prompt) MEDIUM-HIGH, cannot be tested here.
2. **Step 4 refresh does not use RefreshAll at all: it presses keys.** `power_query_extractor/core/
   report_processor.py:258-293` activates the Excel window with `WScript.Shell.AppActivate` and sends
   `SendKeys("%a")`, `"r"`, `"a"` (Alt+A, R, A = Data > Refresh All by keyboard). Then it waits FIXED times
   (default 10 s, plus 5 s of dialog checks, plus `max(10, wait_time)` s; 300-330) and saves (349)
   regardless. No BackgroundQuery, no CalculateUntilAsyncQueriesDone, no "Refreshing" check exist
   anywhere in A (READ, grep). The keystrokes go to whatever window has focus if activation fails
   (273-274 "continuing anyway"). Excel is made **Visible = True** and maximised (225-227). The design
   spec in the repo keeps this on purpose: `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:125-126`
   lists "Rewriting the Power Query refresh itself (the SendKeys approach and its fixed time.sleep waits
   stay as they are)" as out of scope. HIGH.
3. **Step 4 also uses `Dispatch` (not `DispatchEx`) twice per report** (`report_processor.py:218` refresh,
   `:536` extraction) and calls `Quit()` in `finally` (406, 628) with `DisplayAlerts = False`. If
   `Dispatch` attaches to the user's running Excel, that Excel is maximised and driven by keystrokes,
   and then quit; Microsoft documents that Quit with DisplayAlerts False closes without saving.
   READ HIGH for the code; MEDIUM-HIGH that Dispatch attaches to a running Excel (general pywin32
   behaviour, and the premise of Standard 12); cannot test here.
4. **A failed refresh is still saved, keeps the "Refreshed" name, and is later trusted.** Save at
   `report_processor.py:349` happens before/independent of the validation result (345, 354-357). If the
   Excel open fails, the exception is caught inside `_refresh_power_query_simple` (388-393), the copy made
   at 164 stays in the folder under the `_Refreshed_<timestamp>` name, the client list then shows it as
   "last refreshed" (`extractor_window.py:45-82` uses the name to recognise refreshed files and the file
   mtime as the time), and "Skip Refresh" (`report_processor.py:137-155`) will pick the newest such file
   and report it as `success: True` ("Skipped - Using existing refreshed file"). HIGH by reading.
5. **The only proof of a successful refresh is "Info!BB2 is not empty/0/error"** (`report_processor.py:638-685`).
   It cannot tell a fresh value from one already in the file. RefreshDate is never used (MET). HIGH.
6. **openpyxl fallback silently strips Power Query, and the be6418c guard has a gap.** RAN (E1): on a
   synthetic template with a DataMashup part and connections, A's `create_report_from_template` (openpyxl
   path, `core/excel_handler.py:262-322`) returned `True` and the output had lost
   `customXml/item1.xml` (DataMashup), `itemProps1.xml`, its rels and `xl/connections.xml`. There is no
   DataMashup check anywhere in A (READ, grep). The startup guard added in be6418c
   (`utils/requirements_check.py:44-60`) only checks with `find_spec('pythoncom')` that a file exists,
   while `core/excel_handler.py:21-28` needs `import win32com.client` AND `import pythoncom` to succeed on
   `platform.system() == 'Windows'`. RAN (E2, simulated): with a pythoncom that exists but fails to import,
   the guard passes (no block) and excel_handler falls back to openpyxl. The fallback is announced only by
   `logger.warning` lines (98, 267) in the log file/console; the GUI shows "✅ Success"
   (`gui/handlers/processing_handler.py:319`, `:386`). Real-world likelihood of a broken-but-present
   pywin32: LOW-MEDIUM.
7. **Consolidated workbook: 9 columns receive the wrong cell's value.** RAN (E3): in
   `PQ_Extraction_Report_*.xlsx`, the columns `IGST`, `CGST`, `SGST` of the output sheets "IMS Reco",
   "Extra In Purchase" and "Extra In GSTR2B" are filled with the values read from
   `Diff GSTR-3B!J15/K15/L15` instead of `IMS_Reco!AA1/AD1/AG1`, `ExtraInPurchase!O1/P1/Q1`,
   `ExtraInGSTR2B!S1/T1/U1`. Cause: `power_query_extractor/core/data_consolidator.py:141-145` matches the
   column name as a SUBSTRING of the stored field name (`'IGST' in 'GSTR-3B_Difference_Diff_IGST'` is true
   for the first field), and only falls back to the exact key (147-150) when nothing matched. Column names
   with spaces never substring-match (field names use underscores), which is why the other 13 columns are
   right. Technical mismatch between writer (`report_processor.py:581-588`) and reader. Which cell should
   feed which column is the author's [GST RULE] decision; this note only records that the code does not
   do what `cell_mappings.py` configures. Mappings added in shared commit 47a1492 (30-10-2025); lookup
   code from edac6a3. HIGH (RAN).
8. **Step 3 clean-up can be skipped after an Excel error.** In `core/excel_handler.py:238-240`
   `excel.ScreenUpdating = True` and `excel.DisplayAlerts = True` are outside any try. If Excel has
   crashed or is busy, they raise inside `finally`; the exception escapes, so `Quit()` (244),
   `CoUninitialize()` (250) and the staging-file delete (255-260) are skipped and the error propagates to
   `processing_handler.py:324-326` / `400-406`. HIGH by reading.
9. **No guard stops Step 3 and Step 4 running at the same time, and closing the window mid-run kills the
   Excel work without clean-up.** Step 3 uses `app.is_processing` (`processing_handler.py:75,124,528`);
   Step 4 uses its own `ExtractorPanel.is_processing` (`extractor_window.py:119,872,962`); neither checks
   the other. Both workers are daemon threads (`processing_handler.py:129-133`, `extractor_window.py:878-883`).
   `main_window.py:344-355` asks before closing only if Step 3 is running; for Step 4 it destroys the
   window at once, ending the process and killing the daemon thread without its `finally` blocks → a
   visible, maximised Excel with a half-refreshed workbook can be left behind. HIGH (code), MEDIUM (exact
   Excel aftermath).
10. **Originals are safe.** Client input files are only copied (`utils/helpers.py:199`), templates are only
    copied (`core/excel_handler.py:154`, `:279`), Step 3 reports are only copied before refresh
    (`report_processor.py:164`). No source file is ever modified, moved or deleted (READ, grep of
    remove/unlink/rename/move/rmtree). HIGH. But partially-written copies are left behind when a copy fails
    (section 7).
11. **Docs and UI text that disagree with the code**: Setup help says "organise only" is "Fast - no Excel
    automation involved" (`gui/tabs/setup_tab.py:96-97`) — Step 3 drives Excel through COM on Windows;
    Step 4's "unavailable" notice says "Steps 1-3 are unaffected ... creating the reports works without
    pywin32" (`gui/tabs/extract_tab.py:101-105`) — be6418c made pywin32 required on Windows precisely
    because reports lose Power Query without it; Step 4 logs "Opening file in read-only mode" /
    "File opened in read-only mode" (`report_processor.py:541,543`) but never opens read-only; README
    "Clean saves: No VBA macros or external links preserved" (README.md:44) while the openpyxl path uses
    `keep_links=True` (`core/excel_handler.py:282`). HIGH.

---

## 2. Output 1 — Every place A touches an Excel workbook (touchpoint table)

Columns: # | file:line | function | library | action | which workbook | can it hold Power Query? | thread | READ/RAN.
"Worker (S3)" = the Step 3 processing worker started at `gui/handlers/processing_handler.py:129-134`.
"Worker (S4)" = the Step 4 extractor worker started at `power_query_extractor/gui/extractor_window.py:878-883`.
"Template" = the ITC or Sales template chosen in Setup. "Report" = `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`
created in Step 3. "Refreshed copy" = `<report stem>_Refreshed_<ddmmyy_HHMM>.xlsx` made in Step 4.

### 2.1 core/excel_handler.py (Step 3 report creation, summary report)

| # | file:line | function | library | action | which workbook | PQ? | thread | R |
|---|---|---|---|---|---|---|---|---|
| 1 | core/excel_handler.py:19-28 | module import | COM (pywin32) | detect: `USE_WIN32COM = True` only if `platform.system()=='Windows'` and both `import win32com.client` and `import pythoncom` succeed; else only `logging.warning` | – | – | main (import time) | READ |
| 2 | core/excel_handler.py:31-33 | module import | openpyxl | import `load_workbook`, `Workbook`, styles | – | – | main | READ |
| 3 | core/excel_handler.py:53-54 | `__init__` | file | creates `%TEMP%\gst_excel_temp` (local staging folder; `tempfile.gettempdir()`) | – | – | main (`gui/main_window.py:116`) | READ, RAN (E1 shows it is created under TMPDIR) |
| 4 | core/excel_handler.py:58-64 | `__del__` | COM | quit `self.excel_app` — DEAD: `excel_app` is never set to anything but `None` (56) | – | – | whichever thread garbage-collects | READ |
| 5 | core/excel_handler.py:82-87 | `create_report_from_template` | file | template exists? (else `FileNotFoundError`); create output folder | template; report folder | yes (template) | worker (S3) | READ |
| 6 | core/excel_handler.py:90-102 | `create_report_from_template` | – | chooses COM path if `USE_WIN32COM`, else openpyxl path (silent except log warning 98) | – | – | worker (S3) | READ |
| 7 | core/excel_handler.py:120 | `_create_report_win32com` | COM | `pythoncom.CoInitialize()` | – | – | worker (S3) | READ |
| 8 | core/excel_handler.py:123-125 | `_create_report_win32com` | COM | ATTACH: `GetActiveObject("Excel.Application")` (user's running Excel) | user's Excel and all its workbooks | user's own files: unknown | worker (S3) | READ |
| 9 | core/excel_handler.py:126-128 | `_create_report_win32com` | COM | START/ATTACH: `Dispatch("Excel.Application")` (bare `except:` on 126 catches every error from 124) | – | – | worker (S3) | READ |
| 10 | core/excel_handler.py:131-133 | `_create_report_win32com` | COM | `Visible=False`, `DisplayAlerts=False`, `ScreenUpdating=False` on whichever Excel was obtained | – | – | worker (S3) | READ |
| 11 | core/excel_handler.py:147-152 | `_create_report_win32com` | file | delete stale staging file `%TEMP%\gst_excel_temp\<report name>` (errors ignored) | old staging copy | yes | worker (S3) | READ |
| 12 | core/excel_handler.py:154-155 | `_create_report_win32com` | file copy (`shutil.copy2`) | copy template → local staging file named like the final report (`.xlsx`), then `sleep(0.5)` | template → staging copy | yes | worker (S3) | READ |
| 13 | core/excel_handler.py:159 → 324-369 | `_safe_open_workbook` | COM + file | `exists()` check (329); `find_open_workbook` (337 → utils/excel_com.py:19-29: loops `excel.Workbooks`, compares `Path(wb.FullName).absolute()`) and RETURNS an already-open one; else `excel.Workbooks.Open(file_str)` (349) with no other arguments; on error, if code -2147352567 / "path" / length>218, `Workbooks.Open(get_short_path(...))` (361-364); else re-raise (369) | staging copy | yes | worker (S3) | READ |
| 14 | core/excel_handler.py:162-183 | `_create_report_win32com` | COM | find sheet named `Links` (case-insensitive, 167-170); else first sheet whose name contains "link" (174-178); else ADD a new sheet and name it `Links` (182-183) | staging copy | yes | worker (S3) | READ |
| 15 | core/excel_handler.py:189-197 | `_create_report_win32com` | COM | write values: `ws.Range(cell_ref).Value = str(value)` for each mapped cell whose value is non-empty (`if value:` 192); per-cell errors logged and skipped | staging copy, `Links` sheet | yes | worker (S3) | READ |
| 16 | core/excel_handler.py:203 | `_create_report_win32com` | COM | SAVE: `wb.Save()` | staging copy (report from template) | yes | worker (S3) | READ |
| 17 | core/excel_handler.py:206-207 | `_create_report_win32com` | COM | CLOSE: `wb.Close(SaveChanges=True)` (second save), `wb=None`; `sleep(0.5)` (210) | staging copy | yes | worker (S3) | READ |
| 18 | core/excel_handler.py:215 | `_create_report_win32com` | file copy (`shutil.copy2`) | COPY staging → final report path on the target (network) folder; overwrites silently if it exists | report | yes | worker (S3) | READ |
| 19 | core/excel_handler.py:218-224 | `_create_report_win32com` | file | success = `output_path.exists()` only (size logged, not checked) | report | yes | worker (S3) | READ |
| 20 | core/excel_handler.py:232-236 | finally | COM | CLOSE `wb.Close(SaveChanges=False)` if still open (error path) | staging copy | yes | worker (S3) | READ |
| 21 | core/excel_handler.py:238-240 | finally | COM | `ScreenUpdating=True`, `DisplayAlerts=True` (NOT in try: can raise) | – | – | worker (S3) | READ |
| 22 | core/excel_handler.py:242-246 | finally | COM | QUIT: `excel.Quit()` always (condition `self.excel_app is None` is always true) | that Excel (possibly the user's) | – | worker (S3) | READ |
| 23 | core/excel_handler.py:249-252 | finally | COM | `pythoncom.CoUninitialize()` (while `excel`, `ws`, `sheet` locals still reference COM objects) | – | – | worker (S3) | READ |
| 24 | core/excel_handler.py:255-260 | finally | file | delete staging file (errors ignored) | staging copy | yes | worker (S3) | READ |
| 25 | core/excel_handler.py:273-279 | `_create_report_openpyxl` | file copy | delete stale staging file; copy template → staging | template → staging | yes | worker (S3) | READ, RAN (E1) |
| 26 | core/excel_handler.py:282 | `_create_report_openpyxl` | openpyxl | LOAD `load_workbook(local_temp, data_only=False, keep_links=True)` | staging copy of template | yes | worker (S3) | READ, RAN |
| 27 | core/excel_handler.py:289-303 | `_create_report_openpyxl` | openpyxl | find sheet named `Links` (case-insensitive; NO "link" fallback and NO add); write `ws[cell] = str(value)` for non-empty values; bare `except: pass` per cell | staging copy | yes | worker (S3) | READ, RAN |
| 28 | core/excel_handler.py:306-307 | `_create_report_openpyxl` | openpyxl | SAVE `wb.save(local_temp)` — drops Power Query parts (RAN E1); `wb.close()` | staging copy | yes → lost | worker (S3) | READ, RAN |
| 29 | core/excel_handler.py:308-310 | `_create_report_openpyxl` | file copy | COPY staging → final report; success = `exists()` | report | yes (but stripped) | worker (S3) | READ, RAN |
| 30 | core/excel_handler.py:316-322 | `_create_report_openpyxl` finally | file | delete staging file | staging | – | worker (S3) | READ, RAN (E1: staging folder empty afterwards) |
| 31 | core/excel_handler.py:376-405 | `create_summary_report` | openpyxl | CREATE `Workbook()` (381), 5 sheets, SAVE directly to `<Level1>\GST_Processing_Summary_<ddmmyy_HHMM>.xlsx` (402; called from `processing_handler.py:462-469`), `wb.close()` | newly created summary | no (new file) | worker (S3) | READ |
| 32 | core/excel_handler.py:603-713 | `prepare_template_data` | – | builds the Links-cell values (folder paths and file-name stems); no workbook touched | – | – | worker (S3) | READ |

### 2.2 power_query_extractor/core/report_processor.py (Step 4 refresh and extraction)

| # | file:line | function | library | action | which workbook | PQ? | thread | R |
|---|---|---|---|---|---|---|---|---|
| 33 | report_processor.py:20-36 | module import | COM | guarded import of pythoncom, win32com.client, win32api, win32con, win32gui → `WIN32_AVAILABLE` (only `ImportError` caught) | – | – | main (import at startup via `gui/tabs/extract_tab.py:24-30`) | READ, RAN (tests) |
| 34 | report_processor.py:66-77 | `process_client` | – | refuses with "pywin32 is not installed" for both reports if `WIN32_AVAILABLE` is False — no openpyxl fallback in Step 4 | – | – | worker (S4) | READ, RAN (test_optional_pywin32 passes) |
| 35 | report_processor.py:112-113, 125 | `_process_report` | file (glob) | find `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` in the latest version folder and take `report_files[0]` (directory order; the pattern also matches earlier refreshed copies — RAN E4) | report (or an older refreshed copy) | yes | worker (S4) | READ, RAN |
| 36 | report_processor.py:128-130 | `_process_report` | – | refreshed name = `<stem><suffix>.xlsx`, suffix from the user setting with `{timestamp}` = `%d%m%y_%H%M` (minute resolution) | – | – | worker (S4) | READ |
| 37 | report_processor.py:137-159 | `_process_report` | file (glob, stat) | skip-refresh: newest (by mtime) file `<stem>*.xlsx` recognised by `is_refreshed_copy_of` (`refresh_naming.py:58-66`) | existing refreshed copy | yes | worker (S4) | READ |
| 38 | report_processor.py:164 | `_process_report` | file copy (`shutil.copy2`) | COPY report → refreshed copy, network to network, overwrites silently a same-minute name | report → refreshed copy | yes | worker (S4) | READ |
| 39 | report_processor.py:215 | `_refresh_power_query_simple` | COM | `pythoncom.CoInitialize()` | – | – | worker (S4) | READ |
| 40 | report_processor.py:218 | `_refresh_power_query_simple` | COM | START/ATTACH `Dispatch("Excel.Application")` | – | – | worker (S4) | READ |
| 41 | report_processor.py:221-234 | `_refresh_power_query_simple` | COM | `Visible=True`, `DisplayAlerts=False`, `WindowState=-4137` (maximised), 5 tries 2 s apart, then "continuing anyway" | – | – | worker (S4) | READ |
| 42 | report_processor.py:239-249 → 687-799 | `_safe_open_workbook` | COM + file | `exists()` (692); open file and read 1 byte (697-699); `find_open_workbook` (709) returns an already-open copy; if path > 218 chars, prepend `\\?\` or `\\?\UNC\` (720-732); OPEN `Workbooks.Open(file_str)` (737); on failure CLOSE without saving every open workbook with the same NAME (746-749), wait 3 s, OPEN again (751); then `GetShortPathName` + OPEN (761-766); then OPEN without prefix (771-774); then OPEN `Filename=..., UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False` (777-783); then raise (799). Caller retries 3× only for "Call was rejected by callee" (245-247) | refreshed copy | yes | worker (S4) | READ |
| 43 | report_processor.py:252-255 | `_refresh_power_query_simple` | – | wait 5 × 1 s "for workbook to fully load" | – | – | worker (S4) | READ |
| 44 | report_processor.py:260 | `_refresh_power_query_simple` | COM (not Excel) | `Dispatch("WScript.Shell")` | – | – | worker (S4) | READ |
| 45 | report_processor.py:263-276 | `_refresh_power_query_simple` | COM | `shell.AppActivate(excel.Caption)` up to 3×, then continue even if not activated; `sleep(1)` | Excel window | – | worker (S4) | READ |
| 46 | report_processor.py:281-293 | `_refresh_power_query_simple` | keyboard (SendKeys) | REFRESH by keys: `SendKeys("%a")`, 1 s, `"r"`, 1 s, `"a"` → Excel Data > Refresh All (whatever window has focus) | the active workbook (intended: refreshed copy) | yes | worker (S4) | READ |
| 47 | report_processor.py:297-302 | `_refresh_power_query_simple` | – | fixed wait `self.wait_time` s (GUI default 10, spinbox 5-60, `extractor_window.py:270-279`, passed at 864) | – | – | worker (S4) | READ |
| 48 | report_processor.py:305-342, 417-525 | `_find_error_window`, `_get_window_text`, `_close_error_window` | win32gui (Windows API, not COM) | scan ALL top-level windows on the desktop (`EnumWindows` 450) for a visible, enabled dialog (classes `#32770`, `NUIDialog`, `bosa_sdm_msword`) whose title contains any of: microsoft excel, error, problem, cannot, failed, unable, could not, exception, invalid, missing (425-436); read its text; click the first child button whose text contains ok/close/yes/continue (`BM_CLICK` 506-508), else post `WM_CLOSE` (518-520). 5 checks 1 s apart, then `max(10, wait_time)` s more, then one final check | any program's dialog | – | worker (S4) | READ |
| 49 | report_processor.py:345 → 638-685 | `_validate_refresh` | COM | READ VALUES: sheet named `info` (case-insensitive), cell `BB2`; fails if None, blank, "0"/"0.0", or one of 8 error strings | refreshed copy | yes | worker (S4) | READ |
| 50 | report_processor.py:349-351 | `_refresh_power_query_simple` | COM | SAVE `wb.Save()` ALWAYS (even if an error dialog was found or validation failed); `sleep(2)` | refreshed copy, saved by Excel directly on the network | yes | worker (S4) | READ |
| 51 | report_processor.py:361-363 | `_refresh_power_query_simple` | COM | CLOSE `wb.Close(SaveChanges=True)`; `sleep(1)`; `wb` is NOT set to None | refreshed copy | yes | worker (S4) | READ |
| 52 | report_processor.py:397-402 | finally | COM | CLOSE again `wb.Close(SaveChanges=False)` (on a workbook already closed → error swallowed) | refreshed copy | yes | worker (S4) | READ |
| 53 | report_processor.py:404-409 | finally | COM | QUIT `excel.Quit()` (in try/except) | that Excel (possibly the user's) | – | worker (S4) | READ |
| 54 | report_processor.py:411-415 | finally | COM | `CoUninitialize()` (locals `excel`, `wb`, `shell` still alive) | – | – | worker (S4) | READ |
| 55 | report_processor.py:183-185 | `_process_report` | – | `sleep(3)` "for file to be released", then extract regardless of refresh status (188) | – | – | worker (S4) | READ |
| 56 | report_processor.py:533 | `_extract_data` | COM | `CoInitialize()` | – | – | worker (S4) | READ |
| 57 | report_processor.py:536-538 | `_extract_data` | COM | START/ATTACH `Dispatch("Excel.Application")`; `Visible=False`; `DisplayAlerts=False` | – | – | worker (S4) | READ |
| 58 | report_processor.py:541-543 → 687-799 | `_extract_data` | COM | OPEN via the same `_safe_open_workbook` (read-WRITE; logs claim "read-only mode") | refreshed copy (or skip-refresh file) | yes | worker (S4) | READ |
| 59 | report_processor.py:556-601 | `_extract_data` | COM | READ VALUES: for each sheet in `CELL_MAPPINGS[report_type]` (exact, case-sensitive name match 565), `ws.Range(cell).Value`; stored under both field name and bare cell ref (585-588); missing sheet or cell → None | refreshed copy | yes | worker (S4) | READ |
| 60 | report_processor.py:618-636 | finally | COM | CLOSE `wb.Close(SaveChanges=False)`; QUIT; `CoUninitialize()` | refreshed copy | yes | worker (S4) | READ |
| 61 | report_processor.py:801-804 | `cleanup` | – | does nothing (called from `extractor_window.py:966`) | – | – | worker (S4) | READ |

### 2.3 Step 4 consolidated workbook and GUI

| # | file:line | function | library | action | which workbook | PQ? | thread | R |
|---|---|---|---|---|---|---|---|---|
| 62 | power_query_extractor/core/data_consolidator.py:27-31 | `create_report` | openpyxl | CREATE `Workbook()`, remove default sheet | new consolidated workbook | no | worker (S4) (`extractor_window.py:945`) | READ, RAN (E3) |
| 63 | data_consolidator.py:45-169 | `create_report` | openpyxl | one sheet per `EXTRACTION_CONFIG` entry; writes values; `'Missing'` for every column when the report's status is not success (128-129) | new | no | worker (S4) | READ, RAN |
| 64 | data_consolidator.py:175-179 | `create_report` | openpyxl | SAVE directly to `<Level1>\PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` (no staging; same-minute name overwritten) | new | no | worker (S4) | READ, RAN |
| 65 | power_query_extractor/gui/extractor_window.py:20-42 | `_classify_version_reports` | file (scandir) | has_itc / has_sales = any `itc_report_*.xlsx` / `sales_report_*.xlsx` (includes refreshed copies) | report names | – | main | READ |
| 66 | extractor_window.py:45-82 | `_latest_refresh_times` | file (scandir, stat) | "last refreshed" = newest mtime of a file whose NAME looks refreshed | refreshed copies | – | main | READ |
| 67 | extractor_window.py:611-668 | `scan_folder` | file (glob, stat) | Level 1 = newest `Annual Statement-*` by mtime (631); per client newest `Version-*` by mtime (648) | folders | – | main | READ |
| 68 | extractor_window.py:968-975 | `_offer_to_open_report` | shell | `os.startfile(report_path)` → opens the consolidated workbook in the user's own Excel | consolidated | no | main (via `_ui`) | READ |

### 2.4 Other files that decide what Excel touches

| # | file:line | what | R |
|---|---|---|---|
| 69 | gui/handlers/file_handler.py:27-47 | template pickers accept `*.xltx *.xlsx` only | READ |
| 70 | gui/utils/setup_validation.py:21, 56-64 | Setup checklist accepts `.xlsx .xltx .xltm .xlsm` as a template (wider than the picker; a cached `.xlsm/.xltm` path passes) | READ |
| 71 | gui/handlers/processing_handler.py:292-298, 355-356 | report names `ITC_Report_<sanitised client>_<state code>_<ddmmyy_HHMM>.xlsx`, `Sales_Report_...`; always `.xlsx` whatever the template type | READ |
| 72 | utils/requirements_check.py:37-61, 64-70 | startup check of `openpyxl` and `pythoncom` (find_spec, no import); pywin32 essential only when `sys.platform == 'win32'` | READ, RAN (E2, tests) |
| 73 | main.py:66-106 | startup check runs before the GUI import; offers pip install; blocks if essential missing | READ |
| 74 | utils/helpers.py:275-287 | reads the first 8 bytes of each client input file (signature `PK\x03\x04` or OLE) — read-only | READ |
| 75 | utils/helpers.py:172-215, 217-234 | `safe_copy_file` (copy2 + size compare, deletes dest on size mismatch), `create_backup` (copy2) for organised input files | READ |
| 76 | utils/constants.py:99-125 | `EXCEL_TEMPLATE_MAPPING`: Links cells per report (see 9.1) | READ |
| 77 | utils/constants.py:234-246 | `EXCEL_SAFETY` (max_retries, temp_prefix, backup_prefix, verification_size_min, save_options keep_vba/keep_links False) — stored in `self.safety_config` (excel_handler.py:55) but NEVER read anywhere (grep) — dead config, and its `keep_links: False` contradicts `keep_links=True` at excel_handler.py:282 | READ |
| 78 | power_query_extractor/config/cell_mappings.py:33-278 | which sheets/cells Step 4 reads and where they go (see 9.2); `EXTRACTION_OPTIONS` (248-253): `skip_refresh` used; `missing_data_text` used; `auto_refresh_if_missing` and `show_timestamp` never read (grep) | READ |

Can-it-hold-Power-Query reasoning: templates and everything copied from them (staging copies, reports,
refreshed copies) — YES: README.md:14 and :24 and the 12d505f message ("Power Query is preserved: the
report's data sources are ABSOLUTE paths stored in the Links sheet") say the reports carry Power Query.
Summary and consolidated workbooks — NO, created fresh by openpyxl. Client input files — UNKNOWN, but A
never opens them with any Excel library; it only copies them byte-for-byte and reads 8 header bytes.

---

## 3. Output 2 — Each COM session checked against Standard 12 "Excel through COM"

Standard read at `/home/user/Annual_Statement_2.0/.build-docs/toolkit/rebuild_pipeline_FINAL.md:844-877`.

COM sessions in A (places Excel is started or attached):
- **S1** = Step 3 report creation, `core/excel_handler.py:104-260` (`_create_report_win32com`). Runs once
  per report (2 per client), on the Step 3 worker thread.
- **S2** = Step 4 refresh, `power_query_extractor/core/report_processor.py:209-415`
  (`_refresh_power_query_simple`). Once per report, on the Step 4 worker thread.
- **S3** = Step 4 value extraction, `report_processor.py:527-636` (`_extract_data`). Once per report,
  after S2, on the Step 4 worker thread.
- Not an Excel session: `WScript.Shell` COM object inside S2 (260) and the `win32gui` window scan (417-525).
- Dead: `ExcelHandler.__del__` (excel_handler.py:58-64) never has an Excel to quit.

### 3.1 S1 — Step 3, core/excel_handler.py:_create_report_win32com

| Standard 12 rule | Verdict | Evidence |
|---|---|---|
| Use COM only for a PQ workbook the app writes into or refreshes | MET | S1 writes the Links cells of a template copy that carries Power Query (189-197). |
| Before saving with openpyxl into a workbook the app did not create, check for DataMashup | NOT MET (for the fallback, see section 4) | No DataMashup check anywhere in A (grep). |
| Separate hidden Excel via `DispatchEx`, never take over the user's Excel | NOT MET | `GetActiveObject` first (124) = deliberately takes over the user's running Excel; fallback `Dispatch` (127), never `DispatchEx`. History: present since the initial commit edac6a3 (git log -S). |
| `Visible = False` | MET as written (131), with a side effect | When S1 got the user's Excel via GetActiveObject, the user's Excel is hidden and never made visible again (finally 238-246 restores ScreenUpdating and DisplayAlerts only). HIGH by reading; effect on screen MEDIUM-HIGH. |
| `DisplayAlerts = False` | NOT MET at the moment it matters | Set False at 132, but set back to **True** at 240 immediately before `Quit()` at 244. The standard's reason for the rule ("Without DisplayAlerts = False, Quit() can wait forever on an invisible save prompt") applies exactly here: Excel is invisible (131) and, when it is the user's own, may hold unsaved workbooks. HIGH (code); hang MEDIUM. |
| `Quit()` in a `finally`, even after errors | NOT MET (reliably) | Quit is in `finally` (244) inside try/except, but the two property sets before it (239-240) are not protected: if Excel crashed / RPC is gone / Excel is busy, they raise, the exception escapes `finally`, and Quit, CoUninitialize (250) and the staging delete (255-260) never run. Also the "don't quit if Excel was already open" intent (comment 241) is broken: `self.excel_app` is always None (56; never assigned), so Quit closes the user's own Excel when S1 attached to it. HIGH. |
| Open with `IgnoreReadOnlyRecommended=True` | NOT MET | `Workbooks.Open(file_str)` (349) and `Workbooks.Open(short_path)` (364) pass no other arguments. |
| Check `workbook.ReadOnly` afterwards; if true, close and ask the user | NOT MET | Never checked. Note: `shutil.copy2` (154) copies the template's file attributes; if the template file is marked read-only on the share, the staging copy is read-only too, Excel would open it read-only and `wb.Save()` (203) would fail → report "Failed". MEDIUM (copy2 copies permission bits on Windows — the read-only flag — by documented behaviour; Excel reaction untested). |
| `BackgroundQuery = False` on each connection before RefreshAll | NOT APPLICABLE | S1 does not refresh. Caveat: if a template query has "Refresh data when opening the file", opening it (349) could start a refresh that S1 then saves over mid-way (203). LOW; check the template's connection properties. |
| `CalculateUntilAsyncQueriesDone` after RefreshAll and before save | NOT APPLICABLE | No refresh (same caveat). |
| No connection still `Refreshing` before save | NOT APPLICABLE | No refresh (same caveat). |
| `RefreshDate` never used as proof | MET | Not used anywhere in A (grep). |
| Worker thread: `CoInitialize` first | MET | 120, on the Step 3 worker thread (`processing_handler.py:129-134` → 309/376). |
| Excel work in a separate function; `CoUninitialize` after it returns | NOT MET | `CoUninitialize` (250) is called inside the same function that holds the Excel objects. |
| No Excel object still held at CoUninitialize | NOT MET | Locals `excel`, `ws` and the loop variable `sheet` still reference COM objects when 250 runs (`wb` was set to None at 207 only on the success path). HIGH (code); possible effect: lingering EXCEL.EXE or COM errors at garbage collection — MEDIUM-LOW. |
| Excel installed on every PC; logged-in user; never a scheduled job | MET for "never scheduled" (interactive GUI only; no scheduler code). Excel presence is NOT checked: the startup check (`utils/requirements_check.py`) checks only pywin32. Without Excel, `Dispatch` fails → each report returns False and the log advises "Check if template has 'Links' sheet" (`processing_handler.py:322-323`, `396-398`) — a misleading cause. HIGH (code). |
| (Extra) Writing a Links sheet that does not exist | – | If no sheet is called "Links" (or contains "link"), S1 ADDS a new `Links` sheet (182-183), writes the paths there, saves and reports SUCCESS — the template's queries would not read that new sheet. The openpyxl path instead writes nothing and also reports success (295-310). HIGH. |

### 3.2 S2 — Step 4 refresh, report_processor.py:_refresh_power_query_simple

| Standard 12 rule | Verdict | Evidence |
|---|---|---|
| COM only for a PQ workbook the app refreshes | MET | It refreshes the refreshed copy. |
| `DispatchEx`, never take over the user's Excel | NOT MET | `Dispatch` (218). |
| `Visible = False` | NOT MET (deliberate) | `Visible = True` (225, comment "Keep it visible to avoid issues") and maximised (227): the keystroke method needs a visible, focused window. |
| `DisplayAlerts = False` | MET, with a gap | 226, inside a 5-try loop; if all 5 tries fail the code "continues anyway" (234) and DisplayAlerts may still be True. |
| `Quit()` in `finally` | MET in structure | 404-409, inside try/except. Hazard: it quits whatever Excel `Dispatch` returned, which may be the user's (MEDIUM-HIGH), and with DisplayAlerts False Excel closes without saving the user's unsaved work (Microsoft's documented Quit behaviour). If Excel is stuck on a modal dialog, Quit fails ("call rejected") silently and Excel is left running (MEDIUM). |
| `IgnoreReadOnlyRecommended=True` | NOT MET (normal path) | Only on the 6th and last open attempt (777-783); attempts 1-5 (737, 751, 766, 774) pass no options. |
| Check `workbook.ReadOnly` | NOT MET | Never checked. If the copy opens read-only (file locked by antivirus/backup/another user, or read-only attribute copied by copy2 at 164), the keystroke refresh still runs and `wb.Save()` (349) fails → caught (376-381) → "success: False". MEDIUM. |
| `BackgroundQuery = False` on each connection | NOT MET | No connection object is ever touched (grep: no `Connections`, `BackgroundQuery`, `Queries`). |
| `RefreshAll()` then `CalculateUntilAsyncQueriesDone()` before saving | NOT MET | No `RefreshAll()` call; the refresh is keystrokes (281-293). No `CalculateUntilAsyncQueriesDone` anywhere. The time allowed from the last keystroke to the save is fixed: `wait_time` + 5 s of error checks + `max(10, wait_time)` s = 25 s at the default 10, 125 s at the GUI maximum 60 (297-330). |
| No connection still `Refreshing` before save | NOT MET | Never checked. Save at 349 happens whether or not Excel has finished. With background refresh on (Excel's default for Power Query connections — general knowledge, MEDIUM), a save during refresh can store the old data, which is exactly the failure Standard 12 records as tested (old 2 rows saved instead of 60,000). |
| `RefreshDate` never used as proof | MET | Not used. The proof used instead is `Info!BB2` non-empty/non-zero/non-error (638-685), which cannot tell new data from old (see 5.2). |
| `CoInitialize` in the worker thread | MET | 215, on the Step 4 worker thread (`extractor_window.py:878-883` → `process_client` 914). |
| `CoUninitialize` after the Excel function returns | NOT MET | 412, inside the same function. |
| No Excel object still held | NOT MET | `excel`, `wb`, `shell` still referenced at 412. |
| Excel on every PC; logged-in user; never scheduled | MET for "never scheduled". Stronger need than the standard states: the keystroke method needs an unlocked desktop where the user is NOT typing or clicking for the whole wait, since keys go to the focused window (READ 263-293; HIGH). Windows may refuse to bring Excel to the front while the user is busy in another program (foreground-lock rule — general Windows behaviour, MEDIUM); the code then still sends the keys (273-274). |
| (Extra) Desktop-wide dialog clicking | – | `_find_error_window` scans every program's top-level windows (450), not just this Excel's, and `_close_error_window` clicks a button whose text contains "ok", "close", "yes" or "continue" (506) — e.g. "Yes" in a "replace existing file?" dialog of another program titled with one of the keywords. HIGH (code), consequence MEDIUM-LOW. |
| (Extra) Closing other workbooks | – | Open attempt 2 closes, without saving, every workbook in that Excel with the same file name (746-749). If Dispatch attached to the user's Excel and the user has a same-named workbook open, it is closed without saving. LOW likelihood, high harm. |

### 3.3 S3 — Step 4 extraction, report_processor.py:_extract_data

| Standard 12 rule | Verdict | Evidence |
|---|---|---|
| Use COM only for a workbook the app writes into or refreshes | NOT MET | S3 only READS cell values (584) and closes without saving (622). Standard 12 (line 827) says reading a Power Query workbook with openpyxl is fine. (Fact only; not a recommendation.) |
| `DispatchEx` | NOT MET | `Dispatch` (536). |
| `Visible = False` | MET (537) | Same hazard as S1 if Dispatch attached to the user's Excel: their Excel is hidden, then quit (628). |
| `DisplayAlerts = False` | MET (538) | – |
| `Quit()` in `finally` | MET in structure (626-631) | Same shared-instance hazard. |
| `IgnoreReadOnlyRecommended=True` | NOT MET (normal path) | Same `_safe_open_workbook` as S2. Logs say "Opening file in read-only mode" (541) and "File opened in read-only mode" (543), but no `ReadOnly=True` is passed; the last-resort call even passes `ReadOnly=False` (780). |
| Check `workbook.ReadOnly` | NOT MET | Not checked (low impact: nothing is saved). |
| BackgroundQuery / CalculateUntilAsyncQueriesDone / Refreshing | NOT APPLICABLE | No refresh. Caveat as in S1 (refresh-on-open). |
| RefreshDate | MET (not used) | – |
| CoInitialize in worker | MET (533) | – |
| CoUninitialize after return / no object held | NOT MET (633 in the same function; `excel`, `wb`, `ws`, `sheet` alive) | – |
| Excel on every PC / never scheduled | MET for never scheduled | – |

### 3.4 Pattern note (consistency across files)
- S1 attaches with GetActiveObject; S2/S3 use Dispatch; none uses DispatchEx (READ).
- S1 hides Excel; S2 shows and maximises it; S3 hides it (READ).
- Both `_safe_open_workbook` methods are intentionally different (comment `utils/excel_com.py:3-14`,
  commit ba2a753), and neither passes `IgnoreReadOnlyRecommended` in its normal path.
- All three sessions were written in the shared initial commit edac6a3 (31-08-2025) (git log -S for
  `GetActiveObject`, `SendKeys`, `Dispatch("Excel.Application")`, `excel.Visible = True`,
  `IgnoreReadOnlyRecommended`, `_validate_refresh`, `BB2`). After the split, A changed only: staging on
  local disk (12d505f), the shared already-open check (ba2a753), log levels and the stray print (064c634),
  the guarded imports (c463256), the configurable-suffix lookup (2aaa2b1). So B very probably carries the
  same COM sessions — whoever audits B should confirm. (MEDIUM, from history only.)

---

## 4. Output 3 — Every openpyxl save, the DataMashup question, and the pywin32 logic

### 4.1 The three openpyxl saves in A

| Save | Target | Could it hold PQ? | DataMashup check? | Notes |
|---|---|---|---|---|
| `core/excel_handler.py:306` `wb.save(str(local_temp))` | staging copy of the ITC/Sales TEMPLATE (then copied to the report path, 308) | YES | NONE | RAN E1: Power Query parts dropped, function returned True. Reached only when `USE_WIN32COM` is False (90-102). |
| `core/excel_handler.py:402` `wb.save(str(output_path))` | `GST_Processing_Summary_<ts>.xlsx`, created by `Workbook()` (381) | NO | not needed | Written straight to the Level 1 folder (no staging). |
| `power_query_extractor/core/data_consolidator.py:179` `wb.save(report_path)` | `PQ_Extraction_Report_<ts>.xlsx`, created by `Workbook()` (27) | NO | not needed | Written straight to the Level 1 folder. |

- A never loads-and-saves a refreshed report or a client input file with openpyxl (READ: the only
  `load_workbook` is excel_handler.py:282).
- openpyxl 3.1.5 (the version in venv-A) writes the ZIP straight onto the target path:
  `openpyxl/writer/excel.py:291` `ZipFile(filename, 'w', ...)` (READ of the library source). A network
  drop mid-save leaves a broken summary/consolidated file (no temp-then-rename).
- RAN E1 extra: the openpyxl output keeps the TEMPLATE content type (`spreadsheetml.template.main`) even
  though it is named `.xlsx`, because openpyxl remembers that it loaded a template
  (`openpyxl/reader/excel.py:158` `wb.template = wb_part.ContentType in (XLTX, XLTM)`). Whether Excel opens
  such a file without complaint: NOT SURE — check by opening one on Windows.
- The same question applies to the COM path: S1 copies the template bytes into a file named `.xlsx`
  (154) and Excel opens and saves that (349, 203). Whether Excel saves it back as a template-type package
  or a workbook, and whether it warns: NOT SURE (LOW confidence either way). How to check: run Step 3 once
  with an `.xltx` template on Windows, then unzip the report and look for `template.main` in
  `[Content_Types].xml`. The tool has evidently been used in production with templates (picker filter
  `*.xltx *.xlsx`, Setup help `gui/tabs/setup_tab.py:77-80`), which suggests it works, but no test or
  commit message confirms the `.xltx` case.
- Macro templates: the Setup checklist accepts `.xltm`/`.xlsm` (setup_validation.py:21) but the output is
  always named `.xlsx` (processing_handler.py:298, 356). openpyxl path: `keep_vba` is not passed, so macros
  would be dropped. COM path: Excel generally refuses a macro-enabled file named `.xlsx` (MEDIUM-LOW; untested).

### 4.2 What happens when COM is not available

| Situation | What A does | Evidence |
|---|---|---|
| Windows, pywin32 missing (find_spec finds no `pythoncom`) | App refuses to start (after offering to pip-install): pywin32 is "essential" on Windows | `utils/requirements_check.py:44-60` (`essential=ON_WINDOWS`), `main.py:75-102`; READ; tests pass (RAN) |
| Windows, pywin32 files present but `import pythoncom` or `import win32com.client` fails (broken install, DLL not loadable) | Startup check passes (find_spec only looks for the file); `core/excel_handler.py:21-28` catches ImportError → `USE_WIN32COM=False` → **silent openpyxl fallback** → reports without Power Query, GUI says Success | RAN E2 (simulated); READ. Likelihood LOW-MEDIUM. Also: the check looks only for `pythoncom`, not `win32com` (37-61). |
| Windows, pywin32 fine, Excel not installed | COM path; `Dispatch` fails → each report "Failed" (no fallback) — fails safe, misleading hint about the Links sheet | READ excel_handler.py:226-228; processing_handler.py:320-323 |
| Not Windows (Linux/macOS) | pywin32 "optional": startup offers install; then Step 3 uses openpyxl fallback with only log warnings; Step 4 refuses | requirements_check.py:57; excel_handler.py:96-102; report_processor.py:66-77; RAN E1 |
| Step 4 without pywin32 | refuses per client ("pywin32 is not installed ... cannot be refreshed"), no openpyxl fallback | report_processor.py:66-77; RAN tests |

Small inconsistency: excel_handler tests `platform.system() == 'Windows'` (21), requirements_check tests
`sys.platform == 'win32'` (17). On normal Windows Python both are true (HIGH); they differ only on unusual
builds (Cygwin/MSYS). LOW impact.

### 4.3 The pywin32-required logic, commit by commit (all four diffs read in full for the source files)

- **c463256** (26-07-2026 23:10) "Fix: app failed to start when pywin32 was missing". Moved
  report_processor's five win32 imports into `try/except ImportError` with `WIN32_AVAILABLE`
  (report_processor.py:20-36), made `process_client` refuse cleanly (66-77), guarded the
  `ExtractorPanel` import in `gui/tabs/extract_tab.py:24-30` and added the "unavailable" card (84-105).
  Added `tests/test_optional_pywin32.py` (subprocess with win32 modules blocked). The message states
  "Steps 1-3 do not need pywin32 at all (core/excel_handler.py already degrades to openpyxl)" and the card
  text says "creating the reports works without pywin32" — this view is reversed 10 minutes later by
  be6418c, but the card text was never updated (still at extract_tab.py:101-105 in HEAD). Commit message
  explains the root cause: `py` and `python` resolved to different interpreters (one a free-threaded
  build without pywin32).
- **f363946** (23:15) "Check dependencies at startup". New `utils/requirements_check.py`
  (find_spec-based, "so nothing is imported merely to test for it"), wired into `main.py` before the GUI
  import; pywin32 still `essential=False`. Message: "verified that with pywin32 installed ... a live
  Excel.Application Dispatch still works" — evidence that the author's PC test used a real Excel.
- **be6418c** (23:20) "Require pywin32 on Windows so reports cannot lose Power Query".
  `essential=ON_WINDOWS` plus a `consequence` text "Reports would be created WITHOUT their Power Query
  connections, and Step 4 could not refresh them." and "The app stops here rather than produce damaged
  output." (requirements_check.py:44-60, 114-123). `_pause()` helper in main.py (45-54). The guard is
  ONLY at startup in `main.py`; `create_report_from_template` itself still falls back silently
  (excel_handler.py:96-102). Gap shown in 4.2 (RAN E2).
- **a5ad487** (23:27) "Offer to install missing dependencies, then restart". `utils/dependency_installer.py`
  runs `sys.executable -m pip install -r requirements.txt` (26-38, 41-60, 133-134) after a Yes/No dialog,
  600 s timeout, then relaunches with `subprocess.Popen([sys.executable] + sys.argv)` (63-74). Note for
  the Standard 13 auditor: this installs into whatever Python runs the app — on a shared `_runtime` that
  would be a user's PC writing into the shared folder (fact; Standard 13 lines 933-940 reserve that for
  the author). On non-Windows PCs the offer appears at every start because pywin32 is always "missing"
  (main.py:78-95). `requirements.txt` has unpinned versions (`>=`) and includes pytest.
- Earlier related: **064c634** (23-07-2026) added `pywin32>=305; sys_platform == 'win32'` to
  requirements.txt and removed a stray `print` that sat above the docstring of
  `_refresh_power_query_simple`.

---

## 5. Output 4 — Two realistic runs traced by reading (every Excel call in order)

### 5.1 Run (a): Step 3 creates the ITC and Sales reports for <Client A> (Windows, pywin32 present)

Set-up: Setup tab has an ITC template and a Sales template (both `.xltx`), source and target folders;
mode "fresh"; <Client A> has a GSTR-2B Reco, an IMS Reco, monthly GSTR-3B files, a Sales file, a Sales
Reco and an Annual Report. User presses Start.

1. `processing_handler.py:69-134` — guard, input checks (`file_handler.py:132-151`: both templates
   exist), worker thread started (daemon).
2. `processing_handler.py:189-195` — `FileOrganizer(...)`; `file_organizer.py:66-69` creates the target.
3. `processing_handler.py:253` → `file_organizer.py:71-148` creates
   `Annual Statement-<ddmmyy HHMM>\<client-key>\Version-<ddmmyy HHMM>\GSTR3B\ ITC\ Sales\` (names from
   `utils/constants.py:81-96`).
4. `processing_handler.py:271-273` → `file_organizer.py:198-300`: each input file copied with
   `shutil.copy2` + size check (`helpers.py:172-215`); names sanitised; `client_info` updated with the
   new names (241-246). Originals untouched.
5. ITC: `processing_handler.py:292-298` builds `ITC_Report_<client>_<state code>_<ddmmyy_HHMM>.xlsx` in
   the version folder; `prepare_template_data(...,'ITC')` (`excel_handler.py:603-656`) returns 7 values:
   GSTR-3B folder, version folder + Annual Report stem, ITC folder + first GSTR-2B Reco stem, ITC folder +
   first IMS Reco stem.
6. `excel_handler.py:82-87` template exists, folder ensured; `USE_WIN32COM` True → `_create_report_win32com`.
7. Excel calls, in order:
   1. `pythoncom.CoInitialize()` (120)
   2. `win32com.client.GetActiveObject("Excel.Application")` (124).
      - If the user has Excel open → S1 now holds the USER's Excel.
      - Else → `win32com.client.Dispatch("Excel.Application")` (127) starts one.
   3. `excel.Visible = False`; `excel.DisplayAlerts = False`; `excel.ScreenUpdating = False` (131-133)
      → if it is the user's Excel, all of the user's Excel windows disappear now.
   4. delete old staging file (147-152); `shutil.copy2(template, %TEMP%\gst_excel_temp\ITC_Report_....xlsx)`
      (154); sleep 0.5 s.
   5. `_safe_open_workbook`: `exists()`; loop `excel.Workbooks` reading each `FullName` (excel_com.py:26-28);
      `excel.Workbooks.Open(<staging path>)` (349). (On a path error: `GetShortPathName` + Open, 361-364.)
   6. Loop `wb.Sheets` reading `.Name` until `Links` found (167-170).
   7. For B2, B4, B5, B7, B8, B10, B11: `ws.Range(cell).Value = "<text>"` (193) — skipped when the value
      is empty (e.g. no IMS file → B10/B11 left as in the template).
   8. `wb.Save()` (203); `wb.Close(SaveChanges=True)` (206); sleep 0.5 s.
   9. `shutil.copy2(staging, <network>\...\ITC_Report_....xlsx)` (215); `exists()` → return True.
   10. finally: `excel.ScreenUpdating = True`; `excel.DisplayAlerts = True` (239-240); `excel.Quit()` (244)
       → closes the Excel S1 started, OR the user's whole Excel (with prompts ON while invisible: if the
       user has unsaved changes, Excel waits on an invisible "Save changes?" dialog and the worker thread
       waits with it — MEDIUM; if not, the user's Excel simply closes — HIGH by code).
   11. `pythoncom.CoUninitialize()` (250); delete staging file (258).
8. GUI log: "✅ Success: ITC_Report_....xlsx" (processing_handler.py:319).
9. Sales: same sequence with the Sales template; 6 values B2, B3, B5, B6, B8, B9
   (`excel_handler.py:657-708`). A second, separate Excel start/attach/quit. If the user's Excel was
   attached and is still stuck at step 10, `GetActiveObject` may return that stuck instance again
   (MEDIUM-LOW).
10. `file_organizer.py:317-379` writes `_Organization_Report_<client>_<state>.txt` (plain text, raw
    unsanitised client name in the file name — 330).
11. After all clients: `create_summary_report` (`excel_handler.py:371-409`): openpyxl `Workbook()`, 5
    sheets, `wb.save()` straight to `<Level1>\GST_Processing_Summary_<ddmmyy_HHMM>.xlsx`.
12. `messagebox.showinfo` and `os.startfile(level1_folder)` (processing_handler.py:498-516); in 'full'
    workflow, Step 4 is selected and scanned but not started (136-175).

Totals per client: 2 Excel start/attach + 2 Quit cycles, 2 local saves, 2 network copies.

Failure branches in this run (READ):
- Excel busy (user editing a cell in their attached Excel) → `Visible=False` raises "call rejected"
  → except (226-228) → finally: `excel.ScreenUpdating = True` raises again → escapes → no Quit, no
  CoUninitialize, staging file kept → `processing_handler.py:324-326` logs "ITC Report Error". HIGH (code),
  MEDIUM (that Excel rejects calls while in edit mode — general Excel behaviour).
- Template without a `Links` sheet → COM path ADDS one and reports Success (182-183). HIGH.
- Network drop during step 7.9 → copy2 raises → except → returns False; a partial `ITC_Report_....xlsx`
  can remain at the destination (no clean-up of `output_path`) → Step 4 will later find it by its name.
  HIGH (code), MEDIUM (partial file left: `shutil.copy2` on Windows copies in chunks into the destination).
- User closes the main window during processing → asked (main_window.py:346-353); if Yes, the window is
  destroyed 1 s later; the daemon worker dies wherever it is; an S1 Excel left running hidden with the
  staging workbook open. HIGH (code), MEDIUM (how long the orphan Excel stays).

### 5.2 Run (b): Step 4 refresh and extract over several clients

Set-up: Step 4 tab scanned (`extractor_window.py:611-668`): newest `Annual Statement-*` folder by mtime,
each client's newest `Version-*` by mtime; ITC/Sales boxes enabled when any `itc_report_*.xlsx` /
`sales_report_*.xlsx` exists (20-42). Default settings: wait 10 s, suffix `_Refreshed_{timestamp}`, skip
refresh off. Clients <Client A>, <Client B>, <Client C>, both reports ticked. User presses Start Processing.

Per client, per report (ITC then Sales), in order:
1. `report_processor.py:66` WIN32_AVAILABLE True.
2. `glob("ITC_Report_*.xlsx")` in the version folder (112-113); take `report_files[0]` (125).
   - If an earlier refreshed copy exists in that folder, it also matches (RAN E4). Which one is first
     depends on the server's directory order: on an NTFS share the original usually sorts first because
     "." sorts before "_" (MEDIUM); on a Linux-based NAS the order can be arbitrary — on this Linux
     machine the refreshed copy came first (RAN E4). If a refreshed copy is picked, the new copy's name
     doubles up (`..._Refreshed_020126_0900_Refreshed_030126_1005.xlsx`, RAN E4) and its BB2 is already
     filled.
3. `refreshed_file = <stem>_Refreshed_<ddmmyy_HHMM>.xlsx` (128-130).
4. `shutil.copy2(original, refreshed_file)` on the network (164). (Original never touched again.)
5. S2 `_refresh_power_query_simple(refreshed_file)`:
   1. `CoInitialize()` (215); `Dispatch("Excel.Application")` (218).
   2. `Visible=True`, `DisplayAlerts=False`, `WindowState=xlMaximized` (225-227).
   3. `_safe_open_workbook`: `exists()`, 1-byte read, loop `excel.Workbooks` (`FullName`), possibly
      `\\?\` prefix for >218 chars, `Workbooks.Open(path)` (737) [fallbacks 746-783].
   4. sleep 5 × 1 s.
   5. `Dispatch("WScript.Shell")`; `AppActivate(excel.Caption)`; sleep 1 s.
   6. `SendKeys("%a")`, sleep 1, `SendKeys("r")`, sleep 1, `SendKeys("a")` → Refresh All starts (if Excel
      really had focus).
   7. sleep 10 × 1 s; 5 × (`EnumWindows` scan + sleep 1 s); sleep 10 × 1 s; one more scan.
   8. `_validate_refresh`: loop `wb.Sheets` names; `info_sheet.Range('BB2').Value`.
   9. `wb.Save()` (349) — always; sleep 2; `wb.Close(SaveChanges=True)` (361); sleep 1.
   10. return `{'success': not error_detected, ...}`; finally: `wb.Close(SaveChanges=False)` again
       (fails silently), `excel.Quit()`, `CoUninitialize()`.
6. sleep 3 s (185).
7. S3 `_extract_data(refreshed_file)`: `CoInitialize()`; `Dispatch`; `Visible=False`;
   `DisplayAlerts=False`; open (same helper, read-write); for each mapped sheet (ITC: `Diff GSTR-3B`,
   `Reco`, `T6 of R9`, `T8 of R9`, `IMS_Reco`, `ExtraInPurchase`, `ExtraInGSTR2B`; Sales: `Diff - GSTR-1`)
   loop `wb.Sheets` for an exact name, then `ws.Range(cell).Value` per cell; `wb.Close(SaveChanges=False)`;
   `Quit()`; `CoUninitialize()`.
8. Then the Sales report: steps 1-7 again.

So each client = 4 Excel starts/attaches, 4 Quits, 2 keystroke refreshes, 2 network copies, 2 Excel saves
on the network. Minimum fixed waiting per report with default settings: 5+1+2+10+5+10+2+1 = 36 s in S2,
plus 3 s before S3 (READ, counted from the sleeps); at wait 60 s: 136 s + 3 s.

After all clients: `DataConsolidator.create_report(results, level1_folder)` (`extractor_window.py:945`):
openpyxl, one sheet per output group, saved straight to `<Level1>\PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx`;
"Open the consolidated report?" → `os.startfile` (973) opens it in the user's own Excel.

Scenario: **a workbook already open in Excel**
- The ORIGINAL report is open in the user's Excel: `shutil.copy2` (164) reads it; Excel normally lets
  other programs read an open `.xlsx`, so the copy gets the last SAVED version (MEDIUM). Then
  `Dispatch` (218) may attach to that same user Excel (MEDIUM-HIGH): it is made visible and maximised,
  the copy opens inside it and becomes the active workbook, the keystrokes refresh it, it is saved and
  closed, and `Quit()` (406) closes the user's Excel with DisplayAlerts False → the user's unsaved
  changes are discarded without a prompt (MEDIUM-HIGH, untested here). S3 then repeats the
  attach/hide/quit with `Visible=False` (537).
- The REFRESHED name already exists and is open (same client re-run within the same minute): `copy2`
  onto an open file raises PermissionError on Windows → not inside the inner try (167) → caught by the
  outer except (198-207) → `success: False`, no data. HIGH (code).
- The refreshed copy is already open in the attached Excel (e.g. left over from a killed run):
  `find_open_workbook` returns it (709-712) and S2 refreshes and saves that open copy, then closes it.
- If the first `Workbooks.Open` fails, attempt 2 closes same-named workbooks in that Excel WITHOUT saving
  (746-749).
- Nothing in A checks `ReadOnly`; README.md:301, 314 only tell users to close Excel files first.

Scenario: **a refresh that fails**
- Failure WITH an error dialog (e.g. a source file the query needs is missing): one of the 5+1 scans finds
  a dialog titled with "Microsoft Excel"/"error"/... (417-453), reads its text (455-490), clicks
  OK/Close/Yes/Continue or posts WM_CLOSE (492-525). `error_detected=True`. BB2 still checked. **The
  workbook is still saved** (349) under the `_Refreshed_` name. Status `success: False` with the dialog
  text. S3 still extracts (188); `extractor_window.py:924-925` (and 935-936 for Sales) log "(Data extracted despite error)";
  the consolidator writes "Missing" for all of that report's columns (data_consolidator.py:128-129).
- Failure WITHOUT a dialog, or a refresh slower than the fixed wait: nothing detects it; if BB2 holds
  any non-empty, non-zero, non-error value — including one already present in the template or carried
  over from an earlier refresh (if a refreshed copy was picked in step 2) — status is `success: True`
  and the possibly old or half-refreshed figures go into the consolidated workbook as good. HIGH (code).
  A save during an active background refresh may cancel it (Excel asks "This will cancel a pending data
  refresh"; with DisplayAlerts False the default answer is taken) — MEDIUM-LOW, untested.
- Open failure (all open attempts fail): exception caught inside S2 (388-393) → `success: False`. The
  "Open method of Workbooks class failed" branch in `_process_report` (169-181) is therefore effectively
  unreachable — S2 never lets an exception out (its finally steps are each wrapped). The unrefreshed copy
  stays in the folder with the `_Refreshed_` name. S3 then opens it again in another Excel and extracts
  stale values (shown as "Missing" in the consolidated workbook because status is False). The client
  list's "Last Refresh Status" then shows that time as the last refresh (`extractor_window.py:45-82`,
  `refresh_status_display` 759-782), and with "Skip Refresh" ticked the next run reuses it as
  `success: True` (137-155). HIGH (code).
- Also unreachable: `report_processor.py:383-386` (a return after an inner try that always returns).
  The in-repo audit already noted this (PROJECT_AUDIT_ANNUALSTATEMENT.md:340).

Scenario: **Excel crashing**
- During the waits nothing calls Excel, so the crash is seen only by the dialog scan: a Windows
  "Microsoft Excel has stopped working" dialog (class `#32770`, title containing "Microsoft Excel") would
  match and get its "Close the program" button clicked (MEDIUM). Then `_validate_refresh` fails with a
  COM error (caught, returns False), `wb.Save()` raises (caught at 376-381 → `success: False`), finally:
  Close/Quit fail silently, `CoUninitialize()`. The copy on the network is whatever copy2 wrote (or a
  partly saved file if the crash hit during Save — then S3's open may trigger Excel's repair, which with
  DisplayAlerts False may fail or open a repaired file; MEDIUM-LOW). An owner lock file `~$<name>.xlsx`
  may remain in the version folder (normal Excel crash behaviour, MEDIUM); A's file lists ignore it
  (names start with `~$`).
- Crash in S1 (Step 3): see 5.1 "Excel busy" — the finally block raises, so no clean-up; the exception
  reaches the GUI log.
- If the user closes the main window while Step 4 runs: no prompt (main_window.py:346 checks Step 3's flag
  only), the process ends, the daemon worker dies mid-sequence, the visible Excel stays open with the
  refreshed copy and the keystroke sequence may be half-sent. HIGH (code).
- If Step 3 and Step 4 run at once (nothing prevents it): S1's `GetActiveObject` may grab the Excel that
  S2 started (an automation-started Excel that is visible usually registers as the active Excel — MEDIUM-LOW),
  hide it (so the keystrokes miss), and `Quit()` it mid-refresh. Same across processes when
  `launch_extractor.py` runs next to `main.py`. HIGH that no guard exists; MEDIUM-LOW that the collision
  happens as described.

---

## 6. Output 4 addendum — the Step 4 READ path (where extracted figures go)

- `_extract_data` stores each value under the configured field name AND the bare cell reference
  (`report_processor.py:585-588`). Bare cell references could collide between two sheets of the same
  report type; with the current mappings they do not (READ: ITC cells J15,K15,L15,C37,D37,E37,K29,G6,
  AA1,AD1,AG1,O1,P1,Q1,S1,T1,U1 are all different; Sales cells N1,O1,P1,Q1,I1 all different). No
  calculation is configured (every `'calculation': None`), so the cell-ref keys are currently unused.
- **Consolidator lookup mismatch (RAN E3)** — see finding 7. Writer: `report_processor.py:581-588`. Reader:
  `data_consolidator.py:137-150`. 13 of 22 columns get the configured cell; 9 get the GSTR-3B difference
  cells. Precise list in section 10, E3.
- `Range.Value` of a cell showing an Excel error (#N/A, #REF!...) comes back through pywin32 as an
  integer error code, not text (MEDIUM — general pywin32 behaviour; untested here). The consolidator would
  then write that number with format `#,##0.00` (161-162). How to check: put `=NA()` in a mapped cell and
  run Step 4 once.
- A mapped cell holding a date would come back as a timezone-aware datetime; openpyxl refuses
  timezone-aware datetimes, which would make the whole consolidated save fail (LOW — mapped cells are
  amounts; untested).
- Missing sheet/cell → None → "Missing" text (154).
- `_validate_refresh`'s BB2 check and the client-name expectation are technical, not GST rules.

---

## 7. Output 5 — Other file-safety points

### 7.1 Are originals ever modified, moved or deleted?
- Client input files: only read (8-byte signature, `helpers.py:275-287`) and copied (`helpers.py:199`).
  Never modified, moved or deleted. HIGH (grep for `os.remove|unlink|rename|move|replace|rmtree` —
  every hit is on a copy, a staging file or a folder the app made).
- Templates: only copied (`excel_handler.py:154`, `:279`); Excel opens the copy. HIGH. Matches Setup help
  "the template itself is never modified" (setup_tab.py:77-82).
- Step 3 reports: Step 4 works on a copy (`report_processor.py:164`). HIGH. Exception: an attached Excel
  that already has the original open is not affected; only same-NAME workbooks are closed (746-749).
- Destination copies of input files in the organised tree:
  - 'fresh' mode: if the destination exists (same-minute re-run), it is first backed up as
    `<stem>_backup_<ddmmyy_HHMM><ext>` (`helpers.py:217-234`) and then overwritten
    (`file_organizer.py:277-282`).
  - 'rerun' mode: an existing destination is OVERWRITTEN without backup (the `if/elif` at 278-287 has no
    branch for 'rerun'). HIGH.
  - 'resume': skipped (283-287).
- `safe_copy_file` deletes the destination only when sizes differ after the copy (`helpers.py:202-208`).
- Dead code: `FileOrganizer.cleanup_on_error` (`file_organizer.py:381-396`, uses `shutil.rmtree`) is
  never called (grep). `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path` in helpers
  are exported (`utils/__init__.py`) but never called (grep).

### 7.2 Temporary files
- Staging folder `%TEMP%\gst_excel_temp\` (`excel_handler.py:53`), per Windows user, local disk. File name =
  final report name (contains the client name and state code). Deleted in `finally` (best effort,
  255-260, 316-322). Left behind when: the delete fails silently (file still held by Excel, or read-only
  attribute copied from the template — `os.unlink` refuses read-only files on Windows, MEDIUM), or the S1
  `finally` raises before reaching the delete (5.1). RAN E1 confirms clean removal on the openpyxl path.
- `EXCEL_SAFETY['temp_prefix'] = 'temp_gst_'` and `'backup_prefix'` (`constants.py:234-246`) are never
  used; `.gitignore` still lists `temp_gst_*` and `backup_*`.
- `main.py:35-43` creates `logs/` and `temp/` folders inside the APP folder at every start; neither is
  used by the code (grep). The log file is `gst_organizer.log` in the CURRENT folder (`main.py:24`), and
  the standalone extractor writes `pq_extractor.log` in the current folder (`extractor_main.py:12-19`).
  Both logs contain client names and paths (e.g. `excel_handler.py:195` logs every Links value at INFO;
  `report_processor.py:590, 609` log every extracted value via the GUI callback + `logger.info` at 56).
  If the app is started from the shared folder, these land there. (Confidentiality and Standard 13 point;
  I did not open any log file.)

### 7.3 Local disk, then copy to network (12d505f, 23-07-2026)
- Diff read in full. Step 3 only: template → local staging → Excel works locally → one `shutil.copy2` to
  the network (`excel_handler.py:135-155, 212-215`); same for the openpyxl fallback (271-279, 305-308).
- Claim in the comment/message: "Saving locally then copying yields a byte-identical file at the
  destination" and Power Query still works because the Links paths are absolute. Plausible (HIGH that the
  copy is byte-identical; whether any query uses a path RELATIVE to the workbook's own location is
  unknown — LOW; check the templates' M code for `Excel.CurrentWorkbook()`/relative paths). Message also
  says "the win32com path requires real Excel and could not be executed in CI; validate one report on a
  Windows machine before bulk processing".
- Step 4 does NOT stage locally: the refreshed copy is made network-to-network (164) and Excel opens,
  refreshes and saves it directly on the share (349). The summary and consolidated workbooks are written
  straight to the share by openpyxl.

### 7.4 What if the network drops mid-copy / mid-save?
- Step 3 input copies: `shutil.copy2` raises → `safe_copy_file` returns False (213-215); a partial
  destination file can remain (no clean-up in the except) under the sanitised real name → it would be
  used by Power Query if the Links sheet points at it. MEDIUM-HIGH.
- Step 3 report copy (215 / 308): partial `ITC_Report_....xlsx` / `Sales_Report_....xlsx` can remain at
  the destination; Step 4 will pick it up by name. HIGH (code).
- Step 4 refreshed copy (164): partial `_Refreshed_` file can remain; outer except reports the error; the
  name makes it look refreshed in the client list and to "Skip Refresh". HIGH (code).
- Step 4 Excel save over the network (349): a drop during Excel's save can damage the refreshed copy
  (Excel saves via its own temp file and rename, so usually the old file survives — MEDIUM, general Excel
  behaviour). Original unaffected.
- Summary / consolidated openpyxl saves: written directly → a broken file at the final name (MEDIUM-HIGH;
  openpyxl writes straight into the target ZIP, section 4.1).
- Folder creation failing mid-run: `ensure_path_exists` raises → that client is marked failed
  (`processing_handler.py:442-453`) and the loop continues.

### 7.5 Name collisions and overwrites (minute-resolution timestamps)
- Level 1 `Annual Statement-<ddmmyy HHMM>`, version `Version-<ddmmyy HHMM>`, reports
  `..._<ddmmyy_HHMM>.xlsx`, summary `GST_Processing_Summary_<ddmmyy_HHMM>.xlsx` (helpers.py:25-27,
  processing_handler.py:186, 294, 462-464) — two runs in the same minute reuse the same names; reports and
  summary are overwritten without backup (copy2 / openpyxl save overwrite). HIGH (code).
- Refreshed copies `_Refreshed_<ddmmyy_HHMM>` and consolidated `PQ_Extraction_Report_<YYYYmmdd_HHMM>` —
  same. HIGH (code).
- If the overwritten file is open in Excel, the copy/save fails with a permission error (caught, logged).

### 7.6 Read-only attribute carried by copy2
- `shutil.copy2` copies file permission bits; on Windows that includes the read-only flag (documented
  Python behaviour, MEDIUM-HIGH). A read-only template → read-only staging copy and report → Excel opens
  read-only and `Save` fails (S1, S2) → reports or refreshes "Failed"; staging delete fails. How to check:
  look at the template's file properties on the share.

### 7.7 Long paths
- Step 3: path > 218 chars → `GetShortPathName` retry (excel_handler.py:355-366). Staging paths are
  short (%TEMP%), so this mainly protects nothing on the network side (MEDIUM).
- Step 4: path > 218 chars → `\\?\` or `\\?\UNC\` prefix (report_processor.py:720-732). Whether
  `Workbooks.Open` accepts `\\?\` paths is NOT SURE (LOW); the later attempts fall back to short and
  normal paths anyway. Folder names were shortened in 7347383 because "Excel/Power Query fail above ~218
  characters" (constants.py:85-90).

---

## 8. Documentation vs code (high-value discrepancies for the audit)

| Claim | Where | What the code does | Conf. |
|---|---|---|---|
| "Don't quit if Excel was already open" | core/excel_handler.py:241 | Always quits (excel_app never set) | HIGH |
| "Fast - no Excel automation involved" (organise-only workflow) | gui/tabs/setup_tab.py:96-97 | Step 3 creates reports through Excel COM on Windows (excel_handler.py:90-95, 104-260) | HIGH |
| "Steps 1-3 are unaffected: organising files and creating the reports works without pywin32" | gui/tabs/extract_tab.py:101-105 | Since be6418c the app will not start on Windows without pywin32; off Windows reports lose Power Query | HIGH |
| "Opening file in read-only mode" / "File opened in read-only mode" | report_processor.py:541, 543 | Opens read-write | HIGH |
| "Clean saves: No VBA macros or external links preserved" | README.md:44, 262 | openpyxl path keeps external links (`keep_links=True`, excel_handler.py:282); COM path keeps everything | HIGH |
| "All Excel operations use temporary files" / "Verification: Files are verified before finalization" | README.md:7, 42-43, 260-261, 310 | Only Step 3 stages locally; Step 4, summary and consolidated save in place; "verification" = `exists()` only for reports | HIGH |
| "Excel files must be closed during processing" / "Close all Excel files" | README.md:301, 314 | Not checked or enforced anywhere; the code attaches to the user's Excel instead | HIGH |
| `EXCEL_SAFETY.save_options.keep_links: False` | utils/constants.py:240-244 | Never read; code uses `keep_links=True` | HIGH |
| In-repo audit (05-03-2026, f7eb8f1) says requirements.txt lacks pywin32; GetActiveObject "might interact with those files unexpectedly" | PROJECT_AUDIT_ANNUALSTATEMENT.md:20, 643, 661 | pywin32 added 064c634 (stale claim now); the audit did not record the Quit-on-user's-Excel and never-restored Visible | HIGH |
| "Power Query is preserved ... byte-identical" | excel_handler.py:143-146, 12d505f | True for the copy step; depends on the queries using absolute Links paths | MEDIUM |
| "verified that ... a live Excel.Application Dispatch still works" | f363946 message | Confirms a real-Excel smoke test was run on the author's PC (date 26-07-2026) | – |

---

## 9. [GST RULE] items met inside the Excel paths (described only; no judgement)

### 9.1 Links-sheet cells written in Step 3 (template layout; author-owned)
`utils/constants.py:102-125` (`EXCEL_TEMPLATE_MAPPING`), values built in `excel_handler.py:603-713`:
- ITC (sheet `Links`): B2 = GSTR-3B folder; B4 = version folder; B5 = Annual Report file stem;
  B7 = ITC folder; B8 = GSTR-2B Reco file stem; B10 = ITC folder; B11 = IMS Reco file stem.
- Sales (sheet `Links`): B2 = Sales folder; B3 = Sales file stem; B5 = version folder; B6 = Annual Report
  file stem; B8 = Sales folder; B9 = Sales Reco file stem.
- [GST RULE] Only the FIRST file of each type is linked (`[0]` at 628, 639, 650, 666, 682, 698); the
  GSTR-3B entry is a folder (all monthly files) with no file name. Which file feeds the report is the
  author's rule.
- File stems are written without extension (`helpers.py:553-555`); the template's queries must supply the
  extension themselves (LOW technical note; check the M code).

### 9.2 Cells read in Step 4 and where they go — [GST RULE] [TIME-VARYING]
`power_query_extractor/config/cell_mappings.py:33-242`:
- Output "GSTR-3B Difference" (ITC): `Diff GSTR-3B` J15 → "Diff IGST", K15 → "Diff CGST", L15 → "Diff SGST".
- Output "GSTR-1 Difference" (Sales): `Diff - GSTR-1` N1 → "Diff Taxable Value", O1 → "Diff IGST",
  P1 → "Diff CGST", Q1 → "Diff SGST".
- Output "Turnover GSTR-1" (Sales): `Diff - GSTR-1` I1 → "Taxable Value".
- Output "Reco Diff" (ITC): `Reco` C37 → "IGST Diff", D37 → "CGST Diff", E37 → "SGST Diff".
- Output "Table 6J Match" (ITC): `T6 of R9` K29 → "Table 6J Match". [TIME-VARYING: GSTR-9 table number]
- Output "Table 8D Limits" (ITC): `T8 of R9` G6 → "Table 8D Limits". [TIME-VARYING: GSTR-9 table number]
- Output "IMS Reco" (ITC): `IMS_Reco` AA1 → "IGST", AD1 → "CGST", AG1 → "SGST".
- Output "Extra In Purchase" (ITC): `ExtraInPurchase` O1/P1/Q1 → "IGST"/"CGST"/"SGST".
- Output "Extra In GSTR2B" (ITC): `ExtraInGSTR2B` S1/T1/U1 → "IGST"/"CGST"/"SGST".
- What the code actually writes differs for the last three groups (RAN E3, finding 7).
- "Successfully Run" column = ✓ when the refresh status is success, else ✗ and every value "Missing"
  (data_consolidator.py:108-129, 152-154).
- [GST RULE] Display rounding: non-zero numbers get number format `#,##0.00` (data_consolidator.py:160-162);
  the stored value is not rounded.
- Sheet-name matching for reading is exact and case-sensitive (`report_processor.py:565`).

### 9.3 State code in file names — [GST RULE]
Report names use `client_info['state_code']` or `get_state_code(state)` (`processing_handler.py:293`,
`helpers.py:592-618`, table `constants.py:251-291`); an unknown state falls back to initials or the first
3 letters (helpers.py:608-618). Only noted because it shapes the report file names Step 4 globs for.

---

## 10. RAN log (scripts in /root/merge-scratch/run/A-X-EXCEL/exp/, code copy in .../run/A-X-EXCEL/A/)

All runs: `/root/merge-scratch/venv-A/bin/python -B` with `PYTHONDONTWRITEBYTECODE=1`, TMPDIR pointed into
the run folder. A-copy checksums were taken before and compared after: unchanged.

- **E1 `e1_openpyxl_fallback.py`** — synthetic `.xltx` with `Links` + `Info` sheets, `customXml/item1.xml`
  (DataMashup), `customXml/itemProps1.xml`, `customXml/_rels/item1.xml.rels`, `xl/connections.xml` (PQ-style
  OLE DB connection), rels and content types. Called A's `ExcelHandler().create_report_from_template(...)`
  on Linux (`USE_WIN32COM: False`). Output:
  - "Not on Windows - Power Query may not be preserved!" / "Using openpyxl - Power Query content may be lost!"
    (log only)
  - `create_report_from_template returned: True`
  - Parts LOST: `customXml/_rels/item1.xml.rels`, `customXml/item1.xml`, `customXml/itemProps1.xml`,
    `xl/connections.xml`; report has DataMashup: False
  - main part content type of the `.xlsx` output: still `template`
  - Links B2/B4/B5 written as given; staging folder empty afterwards.
  - Limitation: synthetic file, not an Excel-made one; real Excel PQ workbooks also carry query tables and
    other parts — Standard 12 says the loss was tested on real files.
- **E2 `e2_guard_gap.py`** — fake `pythoncom.py` that raises ImportError on import; `platform.system`
  patched to "Windows"; requirements entry forced to Windows-essential. Output: startup check "pywin32
  installed? True", blocking list `[]`; `excel_handler.USE_WIN32COM: False` → openpyxl fallback.
  (Simulation; shows the mechanism, not how often a broken install happens.)
- **E3 `e3_consolidator_lookup.py`** — `data` dict built exactly like `_extract_data` with unique markers,
  run through A's `DataConsolidator.create_report`. 13 OK; 9 MISMATCH:
  `[IMS Reco] IGST: configured IMS_Reco!AA1 -> written Diff GSTR-3B!J15`; CGST AD1 → K15; SGST AG1 → L15;
  `[Extra In Purchase]` O1/P1/Q1 → J15/K15/L15; `[Extra In GSTR2B]` S1/T1/U1 → J15/K15/L15.
- **E4 `e4_glob_pick.py`** — version folder with an original and an earlier `_Refreshed_` copy: the
  pattern `ITC_Report_*.xlsx` matches both; on this Linux filesystem the refreshed copy is listed first;
  sorted (NTFS-like) order puts the original first; picking the refreshed copy yields a doubled name.
- **Tests** — `pytest -p no:cacheprovider tests/test_excel_com.py tests/test_optional_pywin32.py
  tests/test_requirements_check.py tests/test_refresh_naming.py tests/test_dependency_installer.py` in the
  scratch copy: 54 passed. No test in A exercises `_create_report_win32com`, `_create_report_openpyxl`,
  `_refresh_power_query_simple`, `_extract_data` or `DataConsolidator.create_report` (READ: test file list
  and greps).

---

## 11. Possible client names (locations only — not copied)

- `PROJECT_AUDIT_ANNUALSTATEMENT.md` lines 130, 133, 134, 158, 160, 161, 162, 177, 179, 180, 182, 227,
  228, 231, 232 — the in-repo audit's worked examples use what looks like a real company name (with a
  Maharashtra state code). Not copied here.
- Nothing else found in the files/commit messages I read: tests use generic "ABC", "ABC Private
  Limited", "Client01"; README examples use "ABC Ltd". Commit messages of the A-only commits I read contain
  no names.

## 12. Questions for the author

1. Do you or your staff usually keep your own Excel open while running Step 3 or Step 4? Have you ever
   seen your Excel windows vanish, Excel close by itself, or lose unsaved work during a run?
   (Findings 1 and 3.)
2. In the ITC and Sales templates, are the queries set to "Enable background refresh" and/or "Refresh data
   when opening the file"? (Affects whether S2's fixed waits are long enough and whether S1/S3 save or read
   during a refresh.)
3. Are the template files marked read-only on the share? (shutil.copy2 would copy that flag to every
   report and refreshed copy.)
4. Are your templates `.xltx` or `.xlsx`? Does a Step 3 report open in Excel without any "file format or
   extension" warning?
5. What does `Info!BB2` contain in the empty template, and what should it contain after a good refresh?
   (It is the only success check.)
6. Is "Skip Refresh" used in practice? (It can reuse a copy whose refresh failed.)
7. Is the office share a Windows server or a NAS? (Decides which file `report_files[0]` picks when a
   version folder already holds a refreshed copy.)
8. Have the consolidated workbook's "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" columns ever been
   compared with the source report cells? (RAN E3 shows they carry the GSTR-3B difference cells.) Which
   cells should feed them is your rule.
9. Does every PC that runs Step 3 have Excel installed? (Without it, every report fails with a message
   that blames the Links sheet.)
10. Is Excel on any PC set to a non-English language? (The Alt, A, R, A key tips may differ.)
11. Should a refresh that showed an error dialog or failed the BB2 check still be saved under the
    `_Refreshed_` name? (Current behaviour: yes, always saved.)
12. Was Step 3 ever run while Step 4 (or the standalone extractor) was refreshing? (No guard prevents it.)

## 13. INCOMPLETE / limits

- Not verifiable on this Linux machine (traced by reading only): every real Excel/COM behaviour — whether
  `Dispatch` attaches to a running Excel; the screen effects of `Visible`; Quit with DisplayAlerts on/off;
  the SendKeys/key-tip refresh; read-only opening; Excel's handling of a template package named `.xlsx`;
  `\\?\` paths in `Workbooks.Open`; COM error values from `Range.Value`; orphan Excel lifetimes.
- Read in full: core/excel_handler.py, power_query_extractor/core/report_processor.py, utils/excel_com.py,
  power_query_extractor/core/data_consolidator.py, power_query_extractor/config/cell_mappings.py,
  power_query_extractor/core/refresh_naming.py, power_query_extractor/gui/extractor_window.py,
  power_query_extractor/extractor_main.py, launch_extractor.py, gui/tabs/extract_tab.py,
  gui/handlers/processing_handler.py, gui/handlers/file_handler.py, gui/main_window.py, main.py,
  utils/requirements_check.py, utils/dependency_installer.py, utils/helpers.py, core/file_organizer.py,
  gui/utils/setup_validation.py, gui/utils/cache_manager.py, gui/utils/ui_helpers.py, utils/constants.py,
  requirements.txt, pytest.ini, .gitignore, .claude/settings.local.json, tests/test_requirements_check.py,
  tests/test_processing_handler.py, tests/conftest.py, tests/test_excel_com.py and
  tests/test_optional_pywin32.py (via their creating diffs, unchanged since).
- Read only in part (outside Excel scope): gui/handlers/client_handler.py (220-256 export), gui/tabs/setup_tab.py
  (60-110 help texts + grep), core/file_parser.py (grep hits only), README.md (sections on Excel),
  PROJECT_AUDIT_ANNUALSTATEMENT.md (Excel sections), docs/superpowers/specs/2026-07-25-merge-extractor-design.md
  (1-40, 118-149 + grep). Other test files not read.
- Diffs read in full: c463256, be6418c, 12d505f, ba2a753; source-file parts of f363946, a5ad487, 064c634,
  47a1492 (cell_mappings); `git diff 47a1492 A/main` for report_processor.py, excel_handler.py,
  file_organizer.py and the first part of helpers.py. For c1f425e, 510f7b4, bfcb24a, a677670, 2acffed,
  2aaa2b1, 275d9a4, c1583d6 I read the commit messages and file lists only (their final code was read in
  full at HEAD).
- Project B was not examined (not in this task). gst_organizer.log was not opened.
