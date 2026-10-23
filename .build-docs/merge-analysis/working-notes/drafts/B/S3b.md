### B-F18: Find each client's reports for Step 4 (latest run folder, version folder)

> **How the evidence in B-F18 to B-F34 was produced.** Every `file:line` below is in project B at snapshot b0e7d0e and was checked by reading the code (READ). Where a statement says **RAN**, the code was executed on this Linux machine with Python 3.14, PyQt5 5.15 and openpyxl, on a copy of the project in `/root/merge-scratch/run/B-S3b/proj` (`python -B`, Qt in "offscreen" mode, which means Qt draws its windows in memory instead of on a screen), with the home folder redirected into the run folder so that the program's settings folder never touched a real home folder. B cannot even be imported without pywin32 (the Python add-on for Windows that lets Python remote-control Excel; see B-F1 in Section 3), and this machine has neither Windows nor Excel. So small **stand-in modules** (fake versions of `pythoncom`, `win32com.client`, `win32gui`, `win32con`, `win32api`, in `/root/merge-scratch/run/B-S3b/stubs/`) were put in front of the project. They do nothing except **record, in order, what B asks Excel and Windows to do**, and they answer cell reads from real `.xlsx` files made with openpyxl (a Python library that reads and writes `.xlsx` files without Excel). A stand-in run therefore proves B's own Python logic and the order of its requests; it proves nothing about how real Excel or Windows react. Those parts are marked **READ** with a confidence level (HIGH / MEDIUM / LOW) and a way to check. `time.sleep` (Python's "wait N seconds") was replaced in the scripts so that waits were counted, not waited. All client names, folder names and figures are invented placeholders (`ClientA-DL`, `<Client A>`, figures such as 1101). Scripts: `/root/merge-scratch/run/B-S3b/scripts/`; raw outputs: `/root/merge-scratch/run/B-S3b/out/`.

#### What it does (plain English)

Step 4 ("Extract & Refresh") does not use the results that Step 3 just produced in memory. Instead it looks at the disk. It opens the **target folder** saved in the settings, picks **one** run folder named `Annual Statement-<DDMMYY HHMM>` (the code calls it the "Level 1" folder), and inside it lists every **client folder** ("Level 2", for example `<Client A>-DL`) together with each of its **version folders** (for example `Version-050126 0900`) that holds an ITC report or a Sales report. Every version row gets an "ITC" tick box, a "Sales" tick box and a refresh status ("Pending" in orange, or "Refreshed (dd/mm/yy HH:MM)" in green). The user ticks the reports to refresh and extract; the ticked rows become the job list handed to the extraction (B-F19 onwards).

#### Where it lives

| What | File and lines |
|---|---|
| Disk walk that builds the list (`_populate_extraction_from_disk`) | `gui/pyqt_main_window.py:419-504` |
| — which target folder | `gui/pyqt_main_window.py:421` (`self.settings.target_folder`, the last **saved** value) |
| — choice of the run folder: names starting `Annual Statement-`, sorted by name Z→A, first one taken | `gui/pyqt_main_window.py:431-448` |
| — client folders A→Z, version folders Z→A by name | `gui/pyqt_main_window.py:453`, `:457` |
| — "has a report" test: glob (a wildcard file search) `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` | `gui/pyqt_main_window.py:461-464` |
| — refresh status: any `*_Refreshed*` file with `.xlsx/.xls/.xlsm`, not starting `~$`; time = newest such file's modification time | `gui/pyqt_main_window.py:466-479` |
| — folder-read error handling | `gui/pyqt_main_window.py:437-440` (listing the target), `:496-497` (walking clients: warning in the program log only — "program log" here and below means B's own text log file `gst_organizer.log`, B-F32, not the on-screen log box) |
| When it runs | start-up `gui/pyqt_main_window.py:189-190`; after processing `:384`; after extraction `:561`; "Refresh List" button `:171-172` (button built at `gui/views/extraction_view.py:180-193`) |
| Drawing the tree (`populate_tree`) | `gui/views/extraction_view.py:291-356` |
| "No data" messages (`show_no_data_message`) | `gui/views/extraction_view.py:358-362`; texts at `gui/pyqt_main_window.py:423, 427, 439, 443-445, 502-504` |
| Selection buttons ITC Only / Sales Only / Select All / Deselect All | built `gui/views/extraction_view.py:200-224`; helpers `:420-456` |
| Reading the ticks (`get_selected_clients`) | `gui/views/extraction_view.py:368-390` |
| Turning ticks into the job list (`_on_start_extraction`) | `gui/pyqt_main_window.py:506-547` |
| Rules | B-R39 (which run folder, order), B-R136 (tree content and status); the report file patterns are the same as B-R61 |

#### Input → Output

- **Input:** the saved target folder and what is on disk under it (folder names, the report file names, the modification times of `*_Refreshed*` files).
- **Output on screen:** a two-level tree; label "Source: `<run folder name>`" (`gui/views/extraction_view.py:300`).
- **Output to the extraction:** one dictionary (a named list of values) per ticked version, `{'name': '<client folder> / <version folder>', 'client_key': <client folder>, 'latest_version': <version folder path>, 'process_itc': True/False, 'process_sales': True/False}` (`gui/pyqt_main_window.py:522-528`), plus the run folder path and the three Step 4 options (`:540-547`).

#### User interaction

Automatic at start-up, after processing and after every extraction. "Refresh List" rebuilds it by hand. ITC Only / Sales Only / Select All / Deselect All set the ticks; the user can also click single boxes. "Start Extraction" (`gui/views/extraction_view.py:254-270`) reads the ticks.

#### Origin

The disk-driven tree was added in **b0e7d0e** (09-02-2026, B only (after the split)). It replaced the list built from run history that 7e09e5c (08-02-2026, B only (after the split)) had introduced, which itself replaced the list fed from Step 3's results in c258047 (07-02-2026, B only (after the split)) (READ, notes H2 for b0e7d0e; confirmed in this check with `git show c258047:gui/pyqt_main_window.py` — `_populate_extraction_from_results` at line 343 — and `git show 7e09e5c:gui/pyqt_main_window.py` — `_populate_extraction_from_history` at line 374). Before the split, the separate Tkinter extractor window (`power_query_extractor/gui/extractor_window.py`, created in edac6a3, shared - before the split on 30-10-2025; deleted from B in a62d3cc) also scanned the disk, but chose the run folder by newest **modification time** and showed one version per client (`47a1492:power_query_extractor/gui/extractor_window.py:517, 525, 529`, READ via the H2 notes).

#### How the code decides, step by step (READ)

1. `target_str = self.settings.target_folder` (`gui/pyqt_main_window.py:421`). Empty → "No target folder configured." (`:422-424`); not a folder → "Target folder does not exist." (`:425-428`).
2. All sub-folders whose name starts with `Annual Statement-` are sorted **by name, Z to A** (`:432-436`) and the first is taken (`:448`). The names carry day-month-year (`DDMMYY`), so name order is not date order. Listing error → "Error scanning: …" (`:437-440`); none → "No 'Annual Statement-...' folders found in target." (`:442-446`).
3. For every client folder, sorted A→Z (`:453`), and every sub-folder of it, sorted Z→A by name (`:457`) — any sub-folder name counts, not only `Version-…`:
   - `has_itc = any(version_dir.glob("ITC_Report_*.xlsx"))`, `has_sales = …"Sales_Report_*.xlsx"` (`:461-462`); a folder with neither is skipped (`:463-464`). These patterns also match B's own refreshed copies (`ITC_Report_…_Refreshed_….xlsx`), so a version holding only a refreshed copy still counts as "has ITC".
   - Refresh status: every file matching `*_Refreshed*` whose extension is `.xlsx/.xls/.xlsm` and whose name does not start with `~$` (Excel's lock files) (`:467-471`). If any exists: status "Refreshed", time = the newest such file's modification time formatted `%d/%m/%y %H:%M` (`:474-479`). This is per **version**, not per report, and it is a file-name test, not "the refresh succeeded".
   - Client folders with no qualifying version are dropped (`:490`).
   - Any `OSError` (a file-system error such as "access denied") while walking clients is written to the program log only (`:496-497`), and whatever was collected so far is shown.
4. `populate_tree` (`gui/views/extraction_view.py:291-356`) clears the tree, shows "Source: …", writes one bold parent row per client with "N version(s)" (`:307-319`), and one child row per version with the name indented by two spaces (`:323-325`). The ITC cell is made tickable only if `has_itc` (`:328`) but **always** gets the state "Unchecked" (`:329`), which makes Qt draw an empty box even where it cannot be ticked; the same for Sales (`:334-339`). Status text green if exactly "Refreshed", otherwise orange (`:341-350`). All boxes start unticked at every rebuild.
5. On "Start Extraction", `get_selected_clients` (`:368-390`) returns one entry per version with at least one tick; `_on_start_extraction` (`gui/pyqt_main_window.py:506-547`) warns "No versions selected for extraction." (`:508-510`) or "No Level 1 folder available." (`:512-515`), keeps only entries whose version folder still exists (`:519-531`, a missing one only gives a program-log warning), warns "Could not find version folders for the selected items. … Make sure clients have been processed first (Step 3)." if none is left (`:532-538`), and starts the extraction with `output_folder = self.settings.target_folder`, the run folder shown in the tree, and the three options read straight from the Step 4 widgets (`:540-547`).

#### Concrete trace (RAN, `scripts/t18_tree.py`, output `out/t18.txt`)

**Scenario.** The saved target folder holds two run folders: `Annual Statement-311225 1000` (an older batch made on 31-12-2025, one client `ClientD-GJ`) and `Annual Statement-050126 0900` (the batch the user has just processed, made on 05-01-2026). In the newer batch: `ClientA-DL` has three version folders (`Version-050126 0900`, `Version-280226 1400`, `Version-030326 1000`, i.e. 05-01, 28-02 and 03-03-2026), each with an ITC and a Sales report; `ClientB-MH` has one version with only an ITC report plus a `…_Refreshed_060126_1015.xlsx` copy left behind by an earlier refresh that **failed**; `ClientC-TN` has a version folder with no reports; a loose summary file lies in the run folder. Modification times were set to match the dates.

| Step | Code | What happened (RAN) |
|---|---|---|
| Window opens | `gui/pyqt_main_window.py:63 → 190 → 419` | target read from `settings.json` |
| Run-folder choice | `:432-448` | name order Z→A puts `Annual Statement-311225 1000` first → **the 31-12-2025 batch is shown**. Label: "Source: Annual Statement-311225 1000". Only `ClientD-GJ` appears. For comparison the script applied the organiser's rule (newest modification time, `core/file_organizer.py:167`, used by Re-Run/Resume): it would pick `Annual Statement-050126 0900`. |
| User moves the old batch away, presses Refresh List | `:171-172 → 419` | label "Source: Annual Statement-050126 0900" |
| Client order | `:453` | `ClientA-DL`, `ClientB-MH` (`ClientC-TN` dropped: no report; the loose file skipped: not a folder) |
| Version order for `ClientA-DL` | `:457` | `Version-280226 1400`, `Version-050126 0900`, `Version-030326 1000` — by name, so the newest (03-03-2026) is **last** and 28-02 is first |
| Boxes for `ClientB-MH` | `extraction_view.py:328-339` | ITC tickable; Sales **not** tickable but still has a check state, so an empty box is drawn; its "user-checkable" flag is off, so a click does nothing |
| Status for `ClientB-MH` | `:466-479` | **"Refreshed (06/01/26 10:15)" in green**, although that copy came from a failed refresh |
| Status for `ClientA-DL` versions | same | "Pending" in orange (no copies yet) |
| Select All | `extraction_view.py:430-435` | 4 entries: three `ClientA-DL` versions with ITC+Sales, `ClientB-MH / Version-050126 0900` with ITC only. Every version of a client becomes its own job. |
| ITC Only / Sales Only | `:444-456` | ITC Only → 4 entries ITC only; Sales Only → 3 entries (ClientB has no tickable Sales) |
| Start Extraction (handed-over values captured) | `gui/pyqt_main_window.py:518-547` | `name` = `"ClientA-DL / Version-280226 1400"` etc.; `client_key` = `"ClientA-DL"` (never read later, see B-F25); `output_folder` = the saved target; `level1_folder` = `…/Annual Statement-050126 0900`; `wait_time` 10; `suffix_pattern` `_Refreshed_{timestamp}`; `skip_refresh` False |
| User types a different target in Step 1 but does not save, presses Refresh List | `:421` | the tree still shows the **old saved** target's batch |
| After any rebuild | `extraction_view.py:329, 336` | 0 ticked rows: all ticks lost |
| `ClientB-MH` folder unreadable (simulated "access denied") | `:452-497` | program-log warning "Error scanning client folders: [Errno 13] …"; the tree shows `ClientA-DL` only, with **no message on screen** that the list is incomplete; every client after the failing one is also skipped |

#### Traced output versus what the feature claims

The feature claims to show "the latest run" and each client's reports with their refresh status. The trace shows (each item is a defect to carry to Section 7):

1. **Wrong batch:** the "latest" run folder is the one with the highest **name**; with day-first names a 31-12-2025 batch beats a 05-01-2026 batch (RAN). Step 3's own "latest" (Re-Run/Resume) uses modification time, so the two steps can work on different batches (RAN comparison). Only one run folder is ever shown, so clients of earlier Fresh runs disappear from Step 4 (READ, HIGH).
2. **Version order** is by name, not date (RAN), so "the top row" is not the newest version.
3. **"Refreshed" does not mean the refresh worked:** a copy from a failed refresh shows green "Refreshed" (RAN here; also RAN in B-F20 Scenario C and B-F25 Run B, where a version whose refresh check failed turned "Refreshed"). A version holding only a refreshed copy also counts as having a report (READ `:461`, the pattern matches copies; RAN by the MW reader, e6b). A custom file suffix always shows "Pending" (see B-F21).
4. **Untickable cells still draw an empty tick box** (RAN).
5. **Ticks are lost** at every rebuild, including the automatic one after each extraction (RAN).
6. **Saved target, not the box:** "Refresh List" uses the last saved target, so a newly typed target is ignored until something saves the settings (RAN).
7. **Silent partial list** after a folder-read error (RAN).
8. The disk walk runs on the window's own thread (the line of work that draws the window): on a slow network share the window pauses while it walks (READ `gui/pyqt_main_window.py:190, 384, 561` call it from window code; MEDIUM on how long; check by timing start-up against the office share).

#### Older documents versus code

- README.md:171 says "The extractor auto-loads target folder from cache". In B the target comes from `~/.gst_organizer/settings.json` and the tree is filled at start-up from disk (READ `gui/pyqt_main_window.py:50-51, 189-190`) — true in B, with the "latest by name" defect above.
- README.md:180-191 (output structure) lists no refreshed copies, no `PQ_Extraction_Report_*.xlsx` and no `Extraction_Log_*.txt`; B creates all three (B-F21, B-F23, B-F25).

---

### B-F19: Refresh reports through Power Query

#### What it does (plain English)

For each ticked report, Step 4 makes a time-stamped **copy** of the report next to it, opens the copy in a **visible, maximised Excel**, and makes Excel run "Data → Refresh All" by **typing keystrokes** into the screen (Alt+A, then R, then A), as if a person pressed them. Power Query (Excel's built-in data-import feature, whose queries live inside the workbook) then re-reads the client's organised files. B waits fixed times, looks for error dialogs (B-F26), checks one cell (B-F20), saves and closes the copy, and quits Excel. The original report made in Step 3 is never changed.

#### Where it lives

| What | File and lines |
|---|---|
| Per-version job: ITC then Sales, or "Not selected for processing" | `power_query_extractor/core/report_processor.py:39-70` |
| Per-report job: find the report, name the copy, skip-or-refresh, wait 3 s, extract | `report_processor.py:72-163` |
| Which file is "the report": first match of `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` | `report_processor.py:76-77, 89` (rule B-R61) |
| Copy (a plain file copy; overwrites a same-name copy) | `report_processor.py:118-120` |
| The refresh (`_refresh_power_query_simple`) | `report_processor.py:165-366` |
| — COM start (COM = Windows' mechanism for one program to drive another), Excel by `Dispatch` (pywin32's call that connects to Excel — possibly to an Excel that is already running, see below) | `:171, 174` |
| — Visible = True, DisplayAlerts = False, maximised; 5 tries 2 s apart | `:176-190` (rule B-R66) |
| — open the copy, 3 tries 3 s apart only for "Call was rejected by callee" | `:192-205`; opening routine `_safe_open_workbook` `:638-747` (rule B-R99) |
| — 5 s "loading" wait | `:207-211` |
| — bring Excel to the front (`AppActivate`, 3 tries), then `SendKeys` "%a", "r", "a" | `:213-249` (rule B-R67) |
| — fixed waits and error-dialog checks | `:253-298` (rule B-R68; dialogs B-F26) |
| — Info!BB2 check, save, close, result | `:300-330` (rule B-R69, B-R70; check B-F20) |
| — error returns and clean-up (`Close`, `Quit`, `CoUninitialize`) | `:332-366` (rule B-R110) |
| A special "could not open — continue without refresh" branch that can never run | `report_processor.py:122-137` (rule B-R72) |
| Worker that runs the per-version jobs | `gui/controllers/extraction_controller.py:48-84` (thread start `:124-184`) |
| Window side | `gui/pyqt_main_window.py:506-547` |
| Rules | B-R61, B-R66, B-R67, B-R68, B-R69, B-R72, B-R99, B-R110 |

#### Input → Output

`Version-…/ITC_Report_<…>.xlsx` (and/or `Sales_Report_<…>.xlsx`) → a new file `<report name without extension><suffix>.xlsx` in the same folder (default suffix `_Refreshed_<DDMMYY_HHMM>`, see B-F21), refreshed (if the keystrokes worked) and saved by Excel; plus a status `{'success': True/False, 'error': …, 'validation_message': …}` per report.

#### User interaction

The user ticks reports in the Step 4 tree (B-F18), optionally changes the wait time (5-120 s, default 10, `gui/views/extraction_view.py:126-128`), and presses "Start Extraction". During the run the user **must not touch the keyboard or mouse** and the screen must stay unlocked, because the refresh is triggered by typed keys (READ `report_processor.py:213-249`; HIGH).

#### Origin

All the COM steps (`Dispatch`, visible Excel, SendKeys, BB2 check) date from **edac6a3** (31-08-2025, shared - before the split on 30-10-2025) (READ `git show edac6a3:power_query_extractor/core/report_processor.py`, lines 154, 196, 218-228, 589-590). The configurable wait (`wait_time`, `max(10, wait_time)`), the configurable suffix and the ITC/Sales choice came in **bd176bf** (30-10-2025, shared - before the split on 30-10-2025) (READ `git show bd176bf`). **c258047** (07-02-2026, B only (after the split)) only tidied it (bare `except:` narrowed, a stray `print` and an unreachable `return` removed; no change in how Excel is driven — notes PQ, XL §9).

#### Concrete trace (RAN with stand-ins for the Python order of calls; Excel behaviour READ) — `scripts/t19_refresh.py`, `out/t19.txt`

**Scenario.** One version folder `ClientA-DL/Version-050126 0900` holds `ITC_Report_ClientA_DL_050126_0900.xlsx` and `Sales_Report_ClientA_DL_050126_0900.xlsx`; both are ticked; wait time 10 s; default suffix; Skip refresh off. The ITC report's `Info!BB2` holds `<Client A>`.

Every request B made for the **ITC report**, in order (stand-in record; "sleep" = a fixed wait that B would really wait):

| # | Request | Code |
|---|---|---|
| 1 | glob `ITC_Report_*.xlsx`, take element 0 | `report_processor.py:76-77, 89` |
| 2 | copy name `…_0900_Refreshed_211026_0639.xlsx`; `shutil.copy2` (file copy) | `:92-94, 120` |
| 3 | `pythoncom.CoInitialize()` | `:171` |
| 4 | `Dispatch("Excel.Application")` → Excel #1 | `:174` |
| 5 | Excel #1: Visible = True, DisplayAlerts = False, WindowState = −4137 (maximised) | `:181-183` |
| 6 | `Workbooks.Open(<copy>)` with no options | `:197 → 685` |
| 7 | wait 5 s | `:209-211` |
| 8 | `Dispatch("WScript.Shell")`; `AppActivate(excel.Caption)` — B passes Excel's own caption text; the stand-in answered `"Excel"`, so the record shows `AppActivate("Excel")`; what real Excel answers depends on its version (not sure — check on Windows with `print(excel.Caption)`) | `:216, 222` |
| 9 | wait 1 s; `SendKeys("%a")` (Alt+A); wait 1 s; `SendKeys("r")`; wait 1 s; `SendKeys("a")` | `:232-248` |
| 10 | wait 10 s (wait time) + up to 5 × 1 s error checks + 10 s (`max(10, wait time)`) = 25 s | `:256-258, 266-275, 282-286` |
| 11 | final error check; read `Info!BB2` → `<Client A>` | `:289-298, 301 → 607` |
| 12 | `wb.Save()`; wait 2 s; `wb.Close(SaveChanges=True)`; wait 1 s | `:305-319` |
| 13 | clean-up: `wb.Close(SaveChanges=False)` on the already-closed book (the error is swallowed), `Quit()`, `CoUninitialize()` | `:348-366` |
| 14 | wait 3 s "for file to be released" | `:140-141` |
| 15 | extraction in a second Excel (B-F22) | `:144 → 478-587` |

The same 15 steps then ran for the Sales report with a third and fourth Excel. **Waiting time counted: 39 s per report, 78 s for this one version** (5 + 1 + 1 + 1 + 25 + 2 + 1 + 3 = 39; formula 16 + w + max(10, w) + 3 seconds for wait time w), plus two Excel start-ups per report that the stand-in cannot time. At the maximum wait of 120 s it is 16 + 120 + 120 + 3 = 259 s per report. Status returned: `{'success': True, 'error': None, 'validation_message': 'Validated - Client Name: <Client A>'}`. Folder afterwards: the two originals (unchanged) and two `…_Refreshed_211026_0639.xlsx` copies.

**Failure paths (RAN with stand-ins, `scripts/t19b_failures.py`, `out/t19b.txt`):**

| Situation | What B did |
|---|---|
| Excel answers "Call was rejected by callee" (Excel busy) to every open | 3 tries × 3 opening attempts = 9 opens for the refresh (4 per try on Windows when the file has a separate 8.3 short name — the old DOS-style short form of a path, such as `ITC_RE~1.XLS` — `:709-716`; the stand-in's `GetShortPathName` returns the path unchanged, so the run saw 3 per try; RAN in this check: 12 `Workbooks.Open` calls in all = 9 for the refresh + 3 for the extraction), then it **carried on with no workbook**: the 5 s wait, `AppActivate` and the three keystrokes were still sent — into whatever window has the keyboard focus — and the result was `{'success': False, 'error': "'NoneType' object has no attribute 'Save'"}`. Extraction then failed too (no data). |
| The open fails with "Open method of Workbooks class failed" | No keystrokes; result `{'success': False, 'error': 'Open method of Workbooks class failed'}`. The special branch at `:125-137` that should say "Cannot open file for refresh … Proceeding with extraction without refresh" **never ran**, because `_refresh_power_query_simple` catches every error itself (`:339-344`) and returns a status instead of raising. |
| `AppActivate` fails all 3 times | Only a program-log warning (`:229-230`); **the three keystrokes were still sent**; the stand-in then reported success. |

**What real Excel and Windows would do (READ, cannot run here):**

- `Dispatch` (rather than `DispatchEx`, which always starts a separate Excel) may hand back an Excel the user already has open (MEDIUM). Then the user's Excel is made visible and maximised, receives the keystrokes, and is closed by `Quit()` (`:357`) with alerts switched off, so unsaved work in it is not offered for saving. How to check: on a Windows test PC open Excel with an unsaved Book1, run Step 4 on one dummy client, and see whether Book1's Excel is reused or closed.
- `AppActivate(excel.Caption)` (`:222`) looks for a window by Excel's caption text, so it may bring forward the user's own Excel window instead of B's (MEDIUM); the refresh would then run in the user's workbook, and B's copy would be saved unrefreshed.
- Alt+A, R, A are Excel's ribbon "key tips" for Data → Refresh All in an English Excel; another Excel version or display language may map them differently (MEDIUM).
- B never asks Excel whether the queries finished: no `RefreshAll()`, no `BackgroundQuery = False`, no `CalculateUntilAsyncQueriesDone()`, no check of a connection's `Refreshing` (READ, grep finds none of these in B). A slow query may still be running when B saves at `:305` (MEDIUM; depends on the templates, which are not in the repository).
- The copy is **always saved** at `:305`, also after an error dialog or a failed BB2 check (RAN stand-in in B-F26 and B-F20).
- Stop is not honoured inside these waits (see B-F24).

#### Comparison with Standard 12 "Excel through COM" (facts only)

| Standard 12 says (rebuild_pipeline_FINAL.md) | B's refresh does | Where |
|---|---|---|
| Separate hidden Excel with `DispatchEx`, `Visible = False`, `DisplayAlerts = False`; never take over the user's Excel (lines 851-855) | `Dispatch`, `Visible = True`, maximised, `DisplayAlerts = False` | `report_processor.py:174, 181-183` |
| `Quit()` in a `finally` block (line 856) | yes | `:355-360` |
| Open with `IgnoreReadOnlyRecommended=True`; check `workbook.ReadOnly` (lines 857-859) | only on the last-resort open; `ReadOnly` never read | `:725-731` |
| `BackgroundQuery = False`, `RefreshAll()`, `CalculateUntilAsyncQueriesDone()`, no `Refreshing` connection before saving (lines 860-867) | none; keystrokes + fixed waits + BB2 check | `:236-305` |
| `CoInitialize()` first in a worker thread; Excel work in a separate function; `CoUninitialize()` after that function returns (lines 872-875) | `CoInitialize()` first; `CoUninitialize()` inside the same function's clean-up while Excel objects are still held | `:171, 362-366` |
| Excel on every PC, a logged-in user, never a scheduled job (lines 876-877) | also needs an unlocked desktop, Excel in front, nobody using keyboard or mouse | `:213-249` |

#### Traced output versus what the feature claims

"Refresh reports through Power Query" works on the happy path in the stand-in. The defects found by the trace (Section 7): keystrokes are sent with no workbook open and when the Excel window could not be brought forward (RAN); the copy is saved whatever happened (RAN); the "could not open" branch is unreachable (RAN); there is no proof that the refresh finished (READ, MEDIUM in effect); `Dispatch` may take over the user's Excel (MEDIUM). Which report file is used and whether a copy of a copy is refreshed is traced in B-F21.

#### Older documents versus code

- README.md:173-176 says "Click 'Refresh & Extract' to: Refresh Power Query connections …". The button is "Start Extraction" (`gui/views/extraction_view.py:254`); "Refresh Power Query connections" is done by keystrokes, not through Excel's connection objects (READ).
- README.md:52 lists macOS and Linux as supported; the refresh (and the whole window) needs Windows, Excel and pywin32 (RAN: B-F1, B-F30).

---
### B-F20: Refresh validation check (Info!BB2)

> **Evidence for B-F20 to B-F34.** The method described at the start of B-F18 applies unchanged. The trace scripts of 21-10-2026 were **re-run on 23-10-2026** against the same copy (`/root/merge-scratch/run/B-S3b/proj`, byte-identical to the repository at b0e7d0e, RAN `diff -rq`); apart from clock times their output was identical (RAN `diff`). Raw outputs of the re-runs and of the new scripts written for this part (`t23_consolidated.py`, `t24b_progress.py`, `t30_entries.py` with `t32_wrapper.py`, `t34b_jobs.py`, `t34c_close.py`) are in `/root/merge-scratch/run/B-S3b/out2/`. Wherever a result comes from another reader's run (named in the registry notes key: PQ, MW, VW, LT, XL, OF, DC), it says "RAN by <reader>". **Independent check (23-10-2026):** a verifier copied the scripts and stand-ins to `/root/merge-scratch/run/VERIFY-B-S3b/` (a fresh copy of the project, again byte-identical, RAN `diff -rq`), re-ran the 19 trace scripts cited in this part there (`t18_tree`, `t19_refresh`, `t19b_failures`, `t20_validate`, `t21_suffix`, `t21b_sorted`, `t22_extract`, `t23_consolidated`, `t24_window`, `t24b_progress`, `t25_full`, `t26_dialogs`, `t27_settings`, `t28_theme`, `t29_dots`, `t30_entries` with `t32_wrapper`, `t33_help`, `t34b_jobs`, `t34c_close`; outputs in `/root/merge-scratch/run/VERIFY-B-S3b/out2/`, named after the scripts) and got the same results apart from clock times; two extra checks it wrote are in `/root/merge-scratch/run/VERIFY-B-S3b/scripts2/`. Corrections it made are worded in place.

#### What it does (plain English)

After B has pressed Data → Refresh All by keystrokes and waited its fixed times (B-F19), it decides whether that refresh "worked". It looks at exactly one cell: **BB2 on a sheet named "Info"** (any letter case) in the refreshed copy. If no error dialog was noticed (B-F26) **and** BB2 holds something other than nothing, blanks, the text "0"/"0.0" or one of eight error texts, the refresh counts as successful, with the message "Validated - Client Name: <whatever BB2 holds>". Otherwise the refresh counts as failed. This one yes/no answer then decides what the user is told ("Result: ITC: OK" or "FAIL"), whether the consolidated workbook shows ✓ and the figures or ✗ and "Missing" (B-F23), and the "Reports Processed" counts (B-F25).

[GST RULE] (tagged under the "when in doubt" rule: the answer decides whether a report's figures are reported at all). ⚠ The sheet name "Info" and the cell BB2 are positions inside the author's templates, which can change between template versions. The templates are not in the repository, so what fills BB2 is not visible (question for the author).

#### Where it lives

| What | File and lines |
|---|---|
| The check (`_validate_refresh`) | `power_query_extractor/core/report_processor.py:589-636` |
| — look for a sheet whose name, lower-cased, is `info` | `report_processor.py:594-599` |
| — no such sheet → `(False, "Info sheet not found")` | `:601-603` |
| — read `Info!BB2` through COM (`.Value`) | `:605-607` |
| — nothing in the cell (`None`) → `"Cell BB2 is empty"` | `:610-612` |
| — text after removing spaces is empty → `"Cell BB2 is blank"` | `:615-619` |
| — text exactly `"0"` or `"0.0"` → `"Cell BB2 contains 0"` | `:621-623` |
| — upper-cased text is one of `#REF!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#ERROR!` → `"Cell BB2 contains error: <text>"` | `:626-629` |
| — anything else → `(True, "Validated - Client Name: <text>")` | `:631-632` |
| — any Python error inside the check → `(False, "Validation error: <error>")` | `:634-636` |
| Where the check is called: after the waits and the dialog checks, **before** the save | `:300-301` (save at `:305`) |
| How the answer becomes the report's status | `:309-313` (failure → `error_detected = True`, error text `"Refresh validation failed: <message>"`), `:321-330` (status `{'success': not error_detected, 'error': …, 'validation_message': <message>}`) |
| Not called at all when "Skip refresh" re-uses an earlier copy | `:100-111` (status `success: True`, `'Skipped - Using existing refreshed file'`) |
| What uses the answer | `gui/controllers/extraction_controller.py:77-84` (OK/FAIL line); `power_query_extractor/core/data_consolidator.py:97-117` (✓/✗ and "Missing"), `:222-226` (Summary counts); `gui/pyqt_main_window.py:555-557` → `models/run_history.py:83` (stored; but `pq_refreshed` is set True whatever the answer, see B-F25) |
| Rules | B-R70, B-R71 |

#### Input → Output

- **Input:** the refreshed copy while it is still open in the refreshing Excel (the workbook object `wb`), after the waits.
- **Output:** a pair (yes/no, message), folded into the report's status dictionary `{'success': True/False, 'error': <text or None>, 'validation_message': <message>}` (`report_processor.py:326-330`).

#### User interaction

None. It runs automatically after every refresh (not after a skipped one). The user sees its result only as lines in the Step 4 log — all coloured blue as "info", even the ❌ lines (B-F24) — and in the "Result: ITC: OK / FAIL" line, the ✓/✗ column of the consolidated workbook and its Summary counts.

#### Origin

Present with the same messages in the first commit **edac6a3** (31-08-2025, shared - before the split on 30-10-2025) (READ `git show edac6a3:power_query_extractor/core/report_processor.py`, lines 589-615). `_validate_refresh` is unchanged since the split point (READ `git diff 47a1492 B/main -- power_query_extractor/core/report_processor.py`: no hunk inside lines 589-636). B's own commit **c258047** (07-02-2026, B only (after the split)) only removed a stray `print` and an unreachable `return {'success': True}` elsewhere in the refresh function.

#### How the code decides, step by step (READ)

1. `_refresh_power_query_simple` has pressed the keys and waited `wait_time + 5 + max(10, wait_time)` seconds plus checked for dialogs (B-F19, B-F26). It now calls `self._validate_refresh(wb)` (`report_processor.py:301`).
2. Inside, every sheet of the workbook is visited in order; the first whose name lower-cased equals `info` is used (`:596-599`). A sheet called "INFO" or "info" counts; "Info " (with a trailing space) or "Information" does not.
3. `info_sheet.Range('BB2').Value` is read (`:607`). Through real COM a number arrives as a decimal number, text as text, an empty cell as `None`; an Excel error cell is believed to arrive as a large negative whole number such as −2146826246 for `#N/A` (MEDIUM, from pywin32 behaviour; not runnable here — how to check: put `=NA()` in BB2 of a test copy on Windows and look at the Step 4 log line "Cell BB2 populated successfully: …").
4. The value is tested in this order: `None` (`:610`); text form after `strip()` empty (`:615-617`); text form `'0'` or `'0.0'` (`:621`); upper-cased text in the eight-item list (`:626-627`). The first test that fires gives a failure message. If none fires, the check passes (`:631-632`).
5. **BB2 is never compared with the client the version folder belongs to**, and nothing in B ever writes BB2 (READ: the only writes into a report are the Links cells, `core/excel_handler.py:159-167` on the Excel path and `:245-250` on the openpyxl path; grep for `BB2` finds only `report_processor.py:605-631`). The code comment at `:605` ("Check cell BB2 for client name") and the returned message at `:632` ("Validated - Client Name: …") call it "Client Name" — the on-screen log line at `:631` says only "Cell BB2 populated successfully" — so the template is presumably meant to put the client's name there when its queries run (MEDIUM).
6. Back in the refresh function the copy is **saved first** (`:305`), then the answer is applied: a failure sets `error_detected` and the error text `"Refresh validation failed: <message>"` (`:310-313`). The returned status always carries `validation_message`, also when the failure came from an error dialog — so a failed refresh can carry the message "Validated - Client Name: …" (RAN in B-F26).
7. If the workbook never opened (Excel answered "Call was rejected by callee" three times, B-F19), `wb` is `None`; `wb.Sheets` then raises inside the check, which returns `(False, "Validation error: 'NoneType' object has no attribute 'Sheets'")`; the save at `:305` then raises too and the inner `except` (`:332-337`) returns `{'success': False, 'error': "'NoneType' object has no attribute 'Save'"}` without a `validation_message`.

#### Concrete trace (RAN)

**Scenario A — three clients in one run folder, ITC only** (RAN `scripts/t24b_progress.py`, output `out2/t24b.txt`; the real window and worker thread, stand-in COM). `Annual Statement-050126 0900` holds `ClientA-DL`, `ClientB-MH` and `ClientC-TN`, each with one `Version-050126 0900` and one ITC report. In the three reports `Info!BB2` holds, respectively, the text `<Client A>`, **nothing**, and the text `Client C`. The user presses **ITC Only**, then **Start Extraction**.

| Version | What the check saw | Lines sent to the Step 4 log (level) | Result |
|---|---|---|---|
| ClientA-DL | BB2 = `<Client A>` | `✅ Cell BB2 populated successfully: <Client A>` (info) | `Result: ITC: OK` (success) |
| ClientB-MH | BB2 empty | `❌ Cell BB2 is empty (None)` (info); `⚠️ Refresh validation failed: Cell BB2 is empty` (info); `⚠️ Refresh completed with errors: Refresh validation failed: Cell BB2 is empty` (info) | `Result: ITC: FAIL` (error, red) |
| ClientC-TN | BB2 = `Client C` | `✅ Cell BB2 populated successfully: Client C` (info) | `Result: ITC: OK` |

Two facts stand out:

- **The stand-in performs no refresh whatsoever** — it records the keystrokes and does nothing with them — and B still reported ClientA and ClientC as refreshed successfully ("✅ Refresh completed successfully", RAN), because BB2 already held a value in the saved file. The check therefore proves only that BB2 was filled when the file was last saved, not that this refresh ran.
- On screen the first line appeared as `✅ Cell BB2 populated successfully: ` — the text `<Client A>` vanished, because the log widget treats every message as HTML (a web-page markup language in which `<...>` is a tag) (RAN: `displayed` lines in `out2/t24b.txt`). Real client names rarely contain angle brackets, but error texts can (B-F24).

**Scenario B — what the check accepts, value by value** (RAN `scripts/t20_validate.py`, output `out2/t20_validate.txt`; B's own `_validate_refresh` on small stand-in workbook objects):

| `Info` sheet / `BB2` content | Result returned by B |
|---|---|
| sheet `Info`, BB2 = `<Client A>` (this version's own client) | `(True, 'Validated - Client Name: <Client A>')` |
| sheet `Info`, BB2 = `<Client Z>` (**another client**, e.g. left in the template or from an earlier refresh) | `(True, 'Validated - Client Name: <Client Z>')` |
| no sheet named info | `(False, 'Info sheet not found')` |
| sheet named `INFO` | `(True, …)` (letter case ignored) |
| BB2 empty | `(False, 'Cell BB2 is empty')` |
| BB2 = three spaces | `(False, 'Cell BB2 is blank')` |
| BB2 = 0 / 0.0 | `(False, 'Cell BB2 contains 0')` |
| BB2 = text `0.00` | `(True, 'Validated - Client Name: 0.00')` |
| BB2 = text `#N/A` | `(False, 'Cell BB2 contains error: #N/A')` |
| BB2 = text `#SPILL!` (a newer Excel error) | `(True, 'Validated - Client Name: #SPILL!')` |
| BB2 = −2146826246 (the number COM is believed to return for `#N/A`; MEDIUM) | `(True, 'Validated - Client Name: -2146826246')` |
| BB2 = `False` | `(True, 'Validated - Client Name: False')` |
| no workbook (open failed three times) | `(False, "Validation error: 'NoneType' object has no attribute 'Sheets'")` |

**Scenario C — what the answer does downstream** (RAN `scripts/t25_full.py`, output `out2/t25_full.txt`). With ClientB's BB2 empty, the consolidated workbook's row for ClientB shows ✗ and `Missing` in every column of every ITC sheet (e.g. "Table 6J Match": `['ClientB-MH / Version-050126 0900', '✗', 'Missing']`) although the stand-in had read every figure; the Summary says "ITC Reports Processed: 2/3"; the Step 4 tree afterwards shows ClientB as **"Refreshed (21/10/26 06:42)"** in green (the copy exists, B-F18); and run history stores ClientB with `pq_refreshed: True` next to the failed status (B-F25).

**Scenario D — "Skip refresh" bypasses the check** (RAN `scripts/t21b_sorted.py`, output `out2/t21b_sorted.txt`). A version folder holds the ITC report and an earlier copy `…_Refreshed_060126_1015.xlsx` whose BB2 is **empty** (that earlier refresh had failed). With "Skip refresh" ticked, B used the old copy with status `{'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}`; the check never ran; the consolidated workbook showed ✓ and the figures `1101, 1102, 1103` (B-F21).

#### Traced output versus what the feature claims

The feature is meant to tell whether the refresh succeeded. The traces show (each is a defect for Section 7; see the IDs given there):

1. **It passes without any refresh having happened** whenever BB2 already holds a value (RAN, Scenario A) — see B-G6 in Section 7. This is the same weakness the author's Standard 12 records for a connection's `RefreshDate` (rebuild_pipeline_FINAL.md lines 866-867: "never use it as proof of a refresh") — a fact, not a recommendation.
2. **Another client's name in BB2 passes** — BB2 is never compared with the client (RAN, Scenario B) — B-G6.
3. **Values that are not in the short list pass**: the text `0.00`, newer error texts such as `#SPILL!`, `False`, and (MEDIUM) the number that COM returns for an error cell (RAN with an assumed value) — B-G6, B-G11.
4. **The copy is saved before the answer is known and whatever the answer** (`report_processor.py:305` before `:310`; RAN in B-F19) — B-G6.
5. **"Skip refresh" reports success without running the check**, even for a copy whose own BB2 is empty (RAN, Scenario D) — B-G7.
6. **A failed check hides every figure of that report as "Missing"**, although the figures were read; "not ticked", "not found" and "failed" look the same (RAN, Scenario C) — B-G10.
7. **Its messages, including the ❌ ones, all arrive in blue "info"** (RAN: 536 of the 541 lines of Scenario A were "info", one "error", four "success") and text in angle brackets disappears from the screen (RAN) — B-G71, B-G65.
8. **A failure caused by an error dialog keeps the message "Validated - Client Name: …"** in the stored status (RAN in B-F26), which reads like a success in `run_history.json`.

#### Older documents versus code

- README.md never mentions the `Info` sheet or cell BB2 (READ README.md:164-176, 206-253); a template without that sheet makes every refresh count as failed and every figure "Missing". README.md:174 "Refresh Power Query connections" implies a check of the connections; B checks one cell instead.
- The code comment "Check cell BB2 for client name" (`report_processor.py:605`) and the message "Validated - Client Name:" (`:632`) suggest a name check; the code tests only that the cell is filled.

---
### B-F21: Skip refresh and refreshed-file naming

#### What it does (plain English)

Every refresh works on a **copy** of the report (B-F19). This feature decides the copy's name and whether a refresh is needed at all:

- **Naming.** The copy is called `<report name without extension><suffix>.xlsx` and sits next to the report. The suffix is whatever the user typed in the Step 4 box "File suffix:" (default `_Refreshed_{timestamp}`), with the text `{timestamp}` replaced by the current time as `DDMMYY_HHMM` (day, month, two-digit year, hour, minute).
- **Skip refresh.** When the box "Skip refresh (use existing refreshed files)" is ticked, B looks for an earlier copy named `<report name>_Refreshed_*.xlsx`, takes the newest one (by its "last modified" time) and reads the figures from it without refreshing again. If none exists, it refreshes as normal.

#### Where it lives

| What | File and lines |
|---|---|
| The option widgets: wait time 5-120 s, default 10; "File suffix:" free text, default `_Refreshed_{timestamp}`; "Skip refresh" tick, default off | `gui/views/extraction_view.py:124-131`, `:133-138` (box at `:136`), `:140-142` (tick at `:141`) |
| Filled from the saved settings at start-up / collected for saving | `gui/pyqt_main_window.py:185-187` / `:205-207`; fields `models/app_state.py:21-23` |
| Read from the screen when Start Extraction is pressed | `gui/pyqt_main_window.py:544-546` → `gui/controllers/extraction_controller.py:138-152` (kept at `:30-32`) |
| Wait time and suffix handed to the refresh object | `extraction_controller.py:50-51` |
| "Skip refresh" handed over by overwriting a shared setting in the configuration module | `extraction_controller.py:53-55` → `power_query_extractor/config/cell_mappings.py:248-253` (`EXTRACTION_OPTIONS['skip_refresh']`) |
| Which file is "the report": the first match of `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` | `power_query_extractor/core/report_processor.py:76-77, 89` |
| The copy's name | `report_processor.py:91-94` |
| Skip logic: fixed search `<stem>_Refreshed_*.xlsx`, newest by modified time, status "Skipped - Using existing refreshed file"; none found → refresh | `report_processor.py:96-115` (search `:102`, newest `:105`, status `:107-111`, fall-through `:112-115`) |
| The copy itself (plain file copy, overwrites a file of the same name) | `report_processor.py:118-120` |
| The Step 4 tree's "Refreshed"/"Pending" status, which also looks for the fixed text `_Refreshed` | `gui/pyqt_main_window.py:466-479` (B-F18) |
| The default suffix typed by hand in five places | `report_processor.py:28`, `models/app_state.py:22`, `extraction_controller.py:24`, `:139`, `extraction_view.py:136` |
| Unused option "refresh if no copy exists" (the behaviour is hard-wired) | `cell_mappings.py:250` (never read; behaviour at `report_processor.py:112-115`) |
| Rules | B-R61, B-R62, B-R63, B-R64, B-R65 |

#### Input → Output

- **Input:** the suffix text and the tick, read from the screen at Start Extraction; the files already in the version folder.
- **Output:** the name of each new copy (`<stem><suffix>.xlsx`), or — with the tick — the name of the earlier copy that is re-used, plus that report's status (`'Skipped - Using existing refreshed file'` counts as success).

#### User interaction

Step 4 → "Extraction Options": type a suffix, tick or untick "Skip refresh", press **Start Extraction**. The values are used at once but are written to `settings.json` only at the next settings save (a scan, the end of processing, the dark-mode toggle, File > Save Settings or closing the window — B-F27); nothing saves them at Start Extraction (READ `gui/pyqt_main_window.py:540-547` against `:205-207, 210-214`).

#### Origin

- **Skip refresh** and its fixed search for `_Refreshed_` existed in the first commit **edac6a3** (31-08-2025, shared - before the split on 30-10-2025): the Tkinter extractor window had the tick box "Skip Refresh (Use existing refreshed files)" (`edac6a3:power_query_extractor/gui/extractor_window.py:188-198`), which wrote `EXTRACTION_OPTIONS['skip_refresh']` when clicked (`:234-238`), and the copy name was then fixed as `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` (`edac6a3:power_query_extractor/core/report_processor.py`, the lines after "Create refreshed copy"), so the search and the name always agreed (READ).
- The **configurable suffix** (and the wait-time box) came in **bd176bf** (30-10-2025, shared - before the split on 30-10-2025) (READ `git show bd176bf`: `self.suffix_pattern` at line 28, name built from it at line 93) — but the skip search stayed fixed at `_Refreshed_` (line 103). From then on a changed suffix and "Skip refresh" no longer meet.
- In **c258047** (07-02-2026, B only (after the split)) the PyQt5 worker began overwriting `EXTRACTION_OPTIONS['skip_refresh']` at the start of every run (`extraction_controller.py:55`), so the default in the configuration file no longer has any effect. The tree status that also hard-codes `_Refreshed` arrived with the disk-driven tree in **b0e7d0e** (09-02-2026, B only (after the split)).

#### How the code decides, step by step (READ)

1. At Start Extraction the window passes the suffix text exactly as typed (no trimming, no check) and the tick (`gui/pyqt_main_window.py:545-546`).
2. The worker gives the suffix to the refresh object (`extraction_controller.py:51`) and writes the tick into the shared module-wide dictionary `EXTRACTION_OPTIONS` (`:54-55`); it stays there after the run.
3. Per report, "the report" is the **first** file the folder listing returns for `ITC_Report_*.xlsx` (or `Sales_Report_*.xlsx`) (`report_processor.py:76-89`). Earlier copies also start with `ITC_Report_` and end in `.xlsx`, so they match too; which file comes first depends on the order in which the operating system lists the folder.
4. Time stamp `DDMMYY_HHMM` (`:92`); suffix = text with `{timestamp}` replaced (`:93`); copy path = `<stem of "the report"><suffix>.xlsx` (`:94`) — always `.xlsx`, characters that Windows forbids in names are not checked.
5. If the tick is on (`:98-100`): search `<stem>_Refreshed_*.xlsx` (`:102`) — **the fixed text, not the user's suffix**. Found → newest by modified time (`:105`), status success "Skipped - Using existing refreshed file" (`:107-111`), no copy, no refresh, no BB2 check (B-F20). Not found → log "No existing refreshed file found - will refresh" and continue as if unticked (`:112-115`).
6. Otherwise `shutil.copy2(report, copy)` (`:120`) — a copy of the same name (same minute, or a suffix without `{timestamp}`) is overwritten without warning — then the refresh (B-F19).
7. In every case that gets this far, B waits 3 s and extracts the cells from the chosen file (`:140-144`, B-F22). A copy that fails in step 6 (for example the empty suffix of case 5 below) raises an error that jumps straight to the error return at `:154-163`: no refresh, no wait and no extraction for that report (RAN case 5).

#### Concrete trace (RAN `scripts/t21_suffix.py` → `out2/t21_suffix.txt`; `scripts/t21b_sorted.py` → `out2/t21b_sorted.txt`; stand-in COM, waits counted)

Each case starts from a version folder `ClientA-DL/Version-050126 0900` holding `ITC_Report_ClientA_DL_050126_0900.xlsx` (BB2 = `<Client A>`), plus whatever the case adds. "Linux order" is the listing order this machine's disk returned; "name order" simulates the alphabetical listing a local Windows NTFS disk normally returns (the script sorts the listing; NTFS behaviour itself is MEDIUM, not runnable here).

| # | Set-up | What B did | Folder afterwards | Status |
|---|---|---|---|---|
| 1 | default suffix, Skip off | copy `…_0900_Refreshed_211026_0640.xlsx`, refreshed; 39 s of fixed waits counted | report + one copy | success, "Validated - Client Name: <Client A>" |
| 2a | Skip **on**; an earlier copy `…_0900_Refreshed_060126_1015.xlsx` from a **failed** refresh (its BB2 empty); **Linux order** (the old copy is listed first) | took the **old copy** as "the report", searched `…_1015_Refreshed_*` (none), so **refreshed a copy of the copy**: `…_Refreshed_060126_1015_Refreshed_211026_0640.xlsx` | 3 files | failure, "Refresh validation failed: Cell BB2 is empty" |
| 2b | same, **name order** | took the original, found the failed old copy, **re-used it without refreshing** (3 s wait only) | 2 files (nothing new) | **success**, "Skipped - Using existing refreshed file"; the consolidated workbook showed ✓ and the figures 1101 / 1102 / 1103 |
| 3 | suffix `_PQ_{timestamp}`, Skip on, an earlier `…_PQ_060126_1015.xlsx` exists | did **not** find the `_PQ_` copy ("No existing refreshed file found - will refresh"); refreshed again into `…_PQ_211026_0640.xlsx` | 3 files | success |
| 3b | the Step 4 tree's status test on that folder | `*_Refreshed*` matched nothing → the version shows **"Pending"** for ever | — | — |
| 4a | suffix `_Refreshed` (no `{timestamp}`), Skip off, run twice; Linux order | 2nd run took the 1st run's copy as "the report" → `…_Refreshed_Refreshed.xlsx` | 3 files | success |
| 4b | same, name order | 2nd run **overwrote** `…_Refreshed.xlsx` without warning (RAN: modified time changed) | 2 files | success |
| 5 | **empty** suffix | copy name = report name → `shutil.SameFileError`, caught at `report_processor.py:154-163`; no copy, report untouched | 1 file | failure, the error text is the full path twice: "… are the same file" |
| 6 | name order with suffixes that begin with `-`, a space or `(` | the copy sorts **before** the report even in name order (RAN: `…0900-PQ 211026_1000.xlsx`, `…0900 PQ.xlsx`, `…0900(PQ).xlsx` first), so a later run would refresh the copy | — | — |
| 7 | the shared setting after these runs | READ, not shown by this run: the script writes the tick into `EXTRACTION_OPTIONS` itself before each case (as the worker does at `extraction_controller.py:55`), so its final print (`skip_refresh: False`, the value of case 5) only echoes the script's own last write. What the code shows (READ, HIGH): nothing in B ever resets the dictionary, so after a real run it keeps that run's tick until the next run overwrites it | — | — |

#### Traced output versus what the feature claims

1. **A changed suffix breaks "Skip refresh" and the "Refreshed" status** (RAN cases 3, 3b) — B-G9 in Section 7.
2. **"Skip refresh" reports success for any earlier copy, also one whose refresh failed**, and the figures are then shown with ✓ (RAN case 2b) — B-G7.
3. **"The report" can be an earlier copy**, so a later run refreshes a copy of a copy (RAN case 2a/4a on Linux; on NTFS with the default suffix the original normally comes first, MEDIUM; with suffixes starting `-`, space or `(` the copy comes first even in name order, RAN case 6) — B-G8.
4. **An empty suffix fails each report with a full-path error**, with no warning when the box is emptied (RAN case 5); **a suffix without `{timestamp}` silently overwrites the previous copy** (RAN case 4b) — B-G9.
5. **Characters that Windows forbids** (for example `:` or `?`) are not rejected in the box (READ: no check between `extraction_view.py:136` and `report_processor.py:93`); the copy would then fail with an operating-system error, caught at `report_processor.py:154-163` as a FAIL (READ, HIGH for the path; the exact Windows message cannot be produced here).
6. **"Skip refresh" travels through a module-wide setting that the worker overwrites** (READ `extraction_controller.py:54-55`) — B-G76.
7. The options in use are not saved until some later settings save (READ) — a crash before then loses them (B-F27).

#### Older documents versus code

- README.md:247-252 shows `EXTRACTION_OPTIONS = {'skip_refresh': False  # Set to True to skip Power Query refresh, …}` as the way to configure skipping. In B editing that line has no effect, because every run overwrites it with the tick box (`extraction_controller.py:55`) (READ; also DO R85). README.md:249 `auto_refresh_if_missing: True` is never read (`cell_mappings.py:250`, grep).
- README.md:174 and the output tree at README.md:178-204 mention neither the suffix box nor the refreshed copies that pile up in each version folder (no clean-up exists, READ grep: the only `unlink` in B, `utils/helpers.py:177`, removes a damaged Step 3 file copy; nothing removes refreshed copies).

---

### B-F22: Extract cells from refreshed reports

#### What it does (plain English)

Right after each refresh — **also after a failed one** — B starts a second Excel, this time hidden, opens the refreshed copy (or, with "Skip refresh", the re-used copy), reads a fixed list of cells from a fixed list of sheets, closes the file without saving and quits that Excel. The result is a dictionary (a lookup table of name → value) holding each figure twice: under a field name such as `IMS_Reco_IGST` and under its bare cell address such as `AA1`.

[GST RULE] ⚠ — which sheet and cell each figure is read from is the author's configuration in `power_query_extractor/config/cell_mappings.py:33-242` (sheet names, cell positions and the GSTR-9 table labels can change between template versions or years). Described only:

| Report | Input sheet (exact name) | Cells read | Output sheet / column in the consolidated workbook (B-F23) | Rule |
|---|---|---|---|---|
| ITC | `Diff GSTR-3B` | J15, K15, L15 | "GSTR-3B Difference": Diff IGST / Diff CGST / Diff SGST | B-R73 |
| Sales | `Diff - GSTR-1` | N1, O1, P1, Q1 | "GSTR-1 Difference": Diff Taxable Value / Diff IGST / Diff CGST / Diff SGST | B-R74 |
| Sales | `Diff - GSTR-1` | I1 | "Turnover GSTR-1": Taxable Value | B-R75 |
| ITC | `Reco` | C37, D37, E37 | "Reco Diff": IGST Diff / CGST Diff / SGST Diff | B-R76 |
| ITC | `T6 of R9` | K29 | "Table 6J Match": Table 6J Match | B-R77 |
| ITC | `T8 of R9` | G6 | "Table 8D Limits": Table 8D Limits | B-R78 |
| ITC | `IMS_Reco` | AA1, AD1, AG1 | "IMS Reco": IGST / CGST / SGST | B-R79 |
| ITC | `ExtraInPurchase` | O1, P1, Q1 | "Extra In Purchase": IGST / CGST / SGST | B-R80 |
| ITC | `ExtraInGSTR2B` | S1, T1, U1 | "Extra In GSTR2B": IGST / CGST / SGST | B-R81 |

17 ITC cells on 7 sheets and 5 Sales cells on 1 sheet (RAN by the registry task; the table is re-read here, READ). Whether the last three rows reach their own columns is a separate matter (B-F23).

#### Where it lives

| What | File and lines |
|---|---|
| Called after every refresh, whatever its status | `power_query_extractor/core/report_processor.py:139-144` (3 s wait first) |
| `_extract_data` | `report_processor.py:478-587` |
| — COM start; a second Excel by `Dispatch`; hidden; alerts off | `:484-490` |
| — log says "Opening file in read-only mode" / "✓ File opened in read-only mode", but opens through the normal routine | `:492-494` → `_safe_open_workbook` `:638-747` (first try `Workbooks.Open(path)` with no options, `:685`) |
| — the cell list for this report type | `:497-505` (built once at start-up from the table above, `cell_mappings.py:259-278`) |
| — per sheet: exact, letter-case-sensitive name match | `:507-519` |
| — sheet not found → every field of that sheet = `None` (warning to the program log only) | `:521-527` |
| — per cell: `ws.Range(cell).Value`, stored under the field name and under the cell address | `:532-541` |
| — a cell that cannot be read → `None` | `:542-545`; a sheet-level error → `None` for its fields `:547-552` |
| — "Total values extracted: N" and every `field: value` line to the screen log | `:554-562` |
| — any other error → empty dictionary `{}` | `:566-568` |
| — clean-up: close without saving, quit, COM end | `:569-587` |
| Rules | B-R73 to B-R82, B-R101 |

#### Input → Output

- **Input:** the path of the refreshed (or re-used) copy and the report type ("ITC" or "Sales").
- **Output:** `{field name: value, cell address: value, …}`; `None` for every field of a missing sheet; `{}` if the workbook could not be opened. It is stored as both `data` and `extracted_values` of that report's result (`report_processor.py:146-152`; the same dictionary object twice).

#### User interaction

None on screen. The only way to change what is read is to edit `cell_mappings.py` and restart the program (the cell list is built once when the module is first loaded, `cell_mappings.py:259-278`).

#### Origin

The extraction routine and the first three output sheets (rows "GSTR-3B Difference", "GSTR-1 Difference", "Turnover GSTR-1") date from **edac6a3** (31-08-2025, shared - before the split on 30-10-2025); the six further ITC sheets (Reco, T6 of R9, T8 of R9, IMS_Reco, ExtraInPurchase, ExtraInGSTR2B) were added in **47a1492** (30-10-2025, the last shared commit, shared - before the split on 30-10-2025) (READ notes PQ, LT; `git log B/main -- power_query_extractor/config/cell_mappings.py`). **c258047** (07-02-2026, B only (after the split)) only narrowed bare `except:` lines in the clean-up (READ the diff above).

#### How the code decides, step by step (READ)

1. `pythoncom.CoInitialize()`; `Dispatch("Excel.Application")` — the same call as the refresh, so it may hand back an already-running Excel (MEDIUM, see B-F19); `Visible = False`; `DisplayAlerts = False` (`:484-490`). Unlike the refresh, there is no retry loop here — not for setting the Excel properties (the refresh tries 5 times, `:178-190`) and not for "Call was rejected by callee" (the refresh tries 3 times, `:195-205`); only `_safe_open_workbook`'s own three ways of opening apply (RAN in B-F19's failure table: one rejected open sequence made the extraction return `{}`).
2. Open the copy through `_safe_open_workbook` (`:493`): file exists, one byte readable, same-path workbook already open → reuse it, path > 218 characters → long-path prefix, then up to three ways of opening (B-F19 table). The first way passes **no** `ReadOnly` option, although the screen says "read-only mode" (`:492`, `:494`).
3. For each input sheet of this report type, in configuration order (`:507`), look for a sheet whose name **equals** the configured name exactly (`:515-519`). "Diff GSTR-3b" is not "Diff GSTR-3B".
4. Missing sheet: each of its field names gets `None` and a screen line "`<field> = None (sheet not found)`"; the warning "Sheet '…' not found" goes only to the program log (`:521-527`). No cell-address entry is made for those cells.
5. Present sheet: for each cell, `ws.Range(cell).Value` (`:535`). Excel's `.Value` (not `.Value2`) is used, so through COM a currency-formatted cell is believed to arrive as a Python `Decimal` (Python's exact decimal-number type, which openpyxl writes as a number but which B's number-format test does not recognise, B-F23 step 6), a date as a date-time (in recent pywin32 versions with a time zone attached), an error cell as a large negative whole number (all MEDIUM; not runnable here — how to check: on Windows put a currency value, a date and `=NA()` into mapped cells of a test copy and look at `run_history.json` and the consolidated workbook). The value is stored under the field name and under the cell address (`:536-539`).
6. The count line and one line per field are sent to the screen log (`:554-562`) — so every extracted figure also lands in the Step 4 log, in `Extraction_Log_*.txt` and in `gst_organizer.log` (B-F25, B-F32).
7. Clean-up closes the workbook without saving, quits Excel and ends COM, swallowing any errors (`:569-587`).

#### Concrete trace

**Scenario A — a normal ITC and Sales report** (RAN `scripts/t24_window.py` → `out2/t24_window.txt`, through the real window): for `ClientA-DL / Version-050126 0900` the log showed `Processing sheet: 'Diff GSTR-3B'` … `'ExtraInGSTR2B'` (7 sheets) and "✅ Extraction complete. Total values extracted: **34**" for the ITC report (17 cells counted twice, once by field name and once by address) and **10** for the Sales report (5 × 2). Lines such as `✓ GSTR-3B_Difference_Diff_IGST (J15) = 1101` and `✓ IMS_Reco_IGST (AA1) = 5101` went to the screen and into `Extraction_Log_*.txt`. Call order (RAN `scripts/t19_refresh.py`, B-F19 table rows 14-15): after the refresh's Excel had quit and the 3 s wait, a **second** Excel was dispatched, `Visible=False`, `DisplayAlerts=False`, `Workbooks.Open(<copy>)` with **no options**, 17 `Range.Value` reads, `Close(SaveChanges=False)`, `Quit()`, `CoUninitialize()`.

**Scenario B — sheet name typed differently, one sheet missing** (RAN `scripts/t22_extract.py` → `out2/t22_extract.txt`). The refreshed copy of `<Client E>`'s ITC report has its sheet saved as `Diff GSTR-3b` (lower-case "b") and no `T6 of R9` sheet:

| Field | Value returned |
|---|---|
| `GSTR-3B_Difference_Diff_IGST`, `_CGST`, `_SGST` | `None` (sheet "not found" — only the letter case differs) |
| `Table_6J_Match_Table_6J_Match` | `None` (sheet missing) |
| `Reco_Diff_*` | 2101, 2102, 2103 |
| `Table_8D_Limits_Table_8D_Limits` | 4101 |
| `IMS_Reco_*`, `Extra_In_Purchase_*`, `Extra_In_GSTR2B_*` | 5101-5103, 6101-6103, 7101-7103 |

30 entries in all (17 field names + 13 cell addresses; none for the two missing sheets). The open call recorded by the stand-in was `Workbooks.Open('ITC_Report_ClientE_KA_…_Refreshed_….xlsx')` with no options; the screen said "Opening file in read-only mode" and "✓ File opened in read-only mode"; the screen showed `GSTR-3B_Difference_Diff_IGST = None (sheet not found)`; the "Sheet … not found" warnings went to the program log only (READ: `logger.warning` at `report_processor.py:522`, no `self._log`; RAN in a different scenario, B-F25 Run B, where the warning for a missing "T6 of R9" sheet was not in the log file: `are those warnings in Extraction_Log_*.txt? [False]` in `out2/t25_full.txt`).

**Scenario C — unusual value types reaching the consolidated workbook** (RAN `scripts/t22_extract.py` part 2 and `scripts/t23_consolidated.py`; the values were fed as COM is believed to deliver them, MEDIUM): a `Decimal('1250.50')` was written as the number 1250.5 without the `#,##0.00` format; −2146826246 (an error cell as COM delivers it, MEDIUM) was written as the figure **−2,146,826,246.00**; the text `'12,345'` stayed text; the text `'=SUM(A1:A2)'` became a **live formula** in the output; a time-zone-aware date made the **whole consolidated workbook fail** (details in B-F23).

#### Traced output versus what the feature claims

1. **A sheet whose name differs only in letter case, or is missing, gives silent `None`s**, later shown as "Missing" under a ✓ (RAN Scenario B; B-F23) — B-G52 in Section 7.
2. **"Read-only" is announced but the file is opened read-write** (RAN, stand-in open call with no options) — harmless for the file (closed without saving), but the message is wrong and the copy is locked while read — B-G79.
3. **Figures are read even when the refresh failed — and then thrown away** by the consolidated workbook (B-F23) — B-G10.
4. **Values are not checked or converted** (types, error numbers, formulas, dates) (RAN Scenario C with assumed COM values) — B-G11, B-G43.
5. **Warnings about missing sheets reach only the program log**, not the screen or `Extraction_Log_*.txt` (RAN) — B-G71.
6. If the extraction Excel cannot start or open the file, the result is `{}` with the refresh status unchanged — RAN by PQ (scenario `setprop`): ✓ with "Missing" everywhere — B-G10.
7. The second Excel also comes from `Dispatch`, not `DispatchEx` (Standard 12, rebuild_pipeline_FINAL.md:851-855), and two Excel start-ups are needed per report (READ; B-F19) — B-G63, B-G57, B-G77.
8. Cell-address keys from two sheets of one report would collide (latent; none today — READ by LT §4.10, HIGH, and re-checked here: no cell address appears twice among the 17 ITC cells or among the 5 Sales cells of `cell_mappings.py:33-242`; ITC `O1`/`P1`/`Q1` and Sales `O1`/`P1`/`Q1` sit in different reports, so they do not meet) — B-G78.

#### Older documents versus code

- README.md:175 "Extract configured cell values" and README.md:231 "Edit `power_query_extractor/config/cell_mappings.py`" are true; README does not say that a restart is needed after editing, or that sheet names must match exactly including letter case (READ).
- The log texts at `report_processor.py:492, 494` ("read-only mode") disagree with the open call.

---
### B-F23: Consolidated extraction workbook

#### What it does (plain English)

At the end of a Step 4 run that produced at least one result, B writes **one new Excel file** into the run folder shown on Step 4: `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` (year first, to the minute). It has a "Summary" sheet first and then one sheet per configured output table (nine today). Each output sheet has one row per extracted version, labelled "`<client folder> / <version folder>`", a "Successfully Run" column with ✓ or ✗, and the figures read in B-F22. The file is written with openpyxl (it is a new workbook with no Power Query, so openpyxl is allowed by Standard 12, rebuild_pipeline_FINAL.md:827, and :840 names "an Excel report written with `openpyxl`" for summaries).

[GST RULE] ⚠ — three decisions in this feature belong to the author and are described, not judged: (a) **which extracted figure goes under which output column** (the configuration of B-F22); (b) **a report whose refresh check failed shows "Missing" instead of figures**; (c) **non-zero numbers are displayed with the format `#,##0.00`** (two decimals, thousands separator; display only — the stored value is not rounded).

#### Where it lives

| What | File and lines |
|---|---|
| Called once per run, only if there is at least one result | `gui/controllers/extraction_controller.py:86-95` (into `self.level1_folder`, the run folder shown on Step 4; error → log line "  Report error: <text>" only) |
| `create_report` | `power_query_extractor/core/data_consolidator.py:20-174` |
| — new workbook, default sheet removed; header style bold white 12 pt on blue `0078D4` | `:27-42` |
| — one sheet per `EXTRACTION_CONFIG` entry, in configuration order | `:45-49` |
| — title "`<sheet> Report`" in A1, merged A1:C1, headers on row 3 (the `description` variant is unused) | `:51-71` |
| — headers "Client Name", "Successfully Run", then each configured column label | `:73-86` |
| — one row per result: column A = `result['client']` | `:92-94` |
| — ✓ (green `107C10`) if that report's **refresh status** is success, else ✗ (`D83B01`) | `:96-109` |
| — ✗ → every figure of the row = `'Missing'` | `:116-117` |
| — (unused) calculated columns | `:118-124`, `:176-206` |
| — **figure look-up: the first extracted field whose name *contains* the column label; only if that gives nothing, the exact field name** | `:125-138` |
| — nothing (`None`) or empty text → `'Missing'` | `:140-142` |
| — written as is; `#,##0.00` for a whole or decimal number other than 0 | `:144-150` |
| — column widths = longest text + 2, at most 50 | `:157`, `:401-417` |
| — Summary sheet (inserted first) | `:160`, `:208-235` |
| — file name and save; any error is logged and passed on | `:162-174` |
| The text "Missing" | `power_query_extractor/config/cell_mappings.py:252` |
| Rules | B-R83, B-R86, B-R87, B-R88, B-R89, B-R90, B-R91 (the unused `description` variant), B-R92 (the unread `show_timestamp` option, `cell_mappings.py:251`), B-R93 (the two never-called older layouts `_write_section` `:237-339` and `_create_details_sheet` `:341-399`) — the registry gives B-F23 the range B-R86 to B-R93 — plus B-R84 (the unused calculated columns) and the mapping rules B-R73 to B-R81 |

#### Input → Output

- **Input:** the list of per-version results from B-F19/B-F22 (`{'client', 'timestamp', 'itc': {status, data}, 'sales': {status, data}}`) and the run folder.
- **Output:** `<run folder>/PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, and the log line "  Report saved: <name>" (or "  Report error: <text>").

#### User interaction

None; automatic at the end of Step 4. The user opens the file from the run folder; nothing on screen shows its contents or opens it.

#### Origin

The consolidator, including the look-up "first field whose name contains the label", dates from **edac6a3** (31-08-2025, shared - before the split on 30-10-2025). The look-up became harmful when **47a1492** (30-10-2025, shared - before the split on 30-10-2025) added three output sheets whose column labels are single words ("IGST", "CGST", "SGST") (READ notes PQ, LT §4.6; DC S-01). **c258047** (07-02-2026, B only (after the split)) only removed a duplicated header block and narrowed bare `except:` lines; **b0e7d0e** (09-02-2026, B only (after the split)) moved the file from the target folder into the run folder (READ `git show b0e7d0e -- gui/controllers/extraction_controller.py`: `create_report(results, self.output_folder)` became `create_report(results, self.level1_folder)`).

#### How the code decides, step by step (READ)

1. For each of the nine output sheets (`:45`): create the sheet, write the title and the header row (`:49-86`).
2. For each result (`:92`), write the version label in column A (`:94`). Take the report of this sheet's type (`'itc'` or `'sales'`, `:97`) and its `status['success']` (`:98`). This is the **refresh** status of B-F20 — not "were the figures read".
3. ✗ (refresh failed, report not ticked, or report not found): every figure column becomes `'Missing'` (`:116-117`), whatever was read.
4. ✓: for each configured column, walk through the extracted dictionary **in the order the cells were read** and take the first entry whose name contains the column label as a piece of text (`:130-133`); then, only if that value is `None`, try the exact name `<sheet>_<label>` with spaces turned into `_` (`:136-138`).
   - Labels with a space ("Diff IGST", "IGST Diff", "Taxable Value", "Table 6J Match"…) can never be a piece of a field name, because field names have `_` where the label has a space, so they always fall through to the exact configured field name (the one the configuration names for that column).
   - The one-word labels "IGST", "CGST", "SGST" of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" **are** pieces of the first ITC field names, `GSTR-3B_Difference_Diff_IGST` / `_CGST` / `_SGST`, which are read first (B-F22 table, first row). So those nine columns receive the "Diff GSTR-3B" J15/K15/L15 values whenever those cells hold anything — including 0, because only `None` triggers the fallback.
5. `None` or `''` → `'Missing'` (`:141-142`); 0 stays 0; text stays text; a text that starts with `=` is turned into a live formula by openpyxl (RAN in B-F22).
6. Whole or decimal numbers other than 0 get the format `#,##0.00` (`:149-150`); `Decimal` values, 0 and text do not.
7. Summary sheet (`:208-235`): "Extraction Summary", "Generated: <YYYY-MM-DD HH:MM:SS>", "Total Clients: <number of results>", "ITC Reports Processed: <x>/<n>" and "Sales Reports Processed: <y>/<n>", where x and y count **refresh successes**, then "Sheets Created:" and a bullet per sheet.
8. Name `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, saved into the run folder (`:163-167`); a file of the same name (same minute) is overwritten. Any error is logged and passed on (`:172-174`); the worker writes "  Report error: <text>" to the Step 4 log and carries on (`extraction_controller.py:94-95`).

#### Concrete trace (RAN `scripts/t23_consolidated.py`, output `out2/t23.txt`; real `ReportProcessor` with stand-in COM, real `DataConsolidator`, real file)

**Scenario.** Run folder `Annual Statement-050126 0900`, three versions, invented figures (each client's cells carry distinct numbers so that every figure can be traced back to its cell):

- `ClientA-DL`: ITC and Sales ticked, both fine. ITC cells J15/K15/L15 = 1101/1102/1103, Reco C37/D37/E37 = 2101-2103, K29 = 3101, G6 = 4101, **IMS_Reco AA1/AD1/AG1 = 5101/5102/5103, ExtraInPurchase O1/P1/Q1 = 6101-6103, ExtraInGSTR2B S1/T1/U1 = 7101-7103**; Sales N1-Q1 = 8101-8104, I1 = 9101.
- `ClientB-MH`: ITC ticked, but its refresh check fails (Info!BB2 empty, B-F20); its Sales report exists but is **not ticked**. Figures = ClientA's + 10000.
- `ClientC-TN`: ITC report **without the sheet "Diff GSTR-3B"**; Sales with O1 = the text `12,345`, P1 empty, Q1 = 0. Figures = ClientA's + 20000.

**Statuses produced by B-F19/B-F20:** ClientA ITC and Sales success; ClientB ITC `{'success': False, 'error': 'Refresh validation failed: Cell BB2 is empty'}`, Sales `{'success': False, 'error': 'Not selected for processing'}`; ClientC ITC and Sales success (program log: "⚠️ Sheet 'Diff GSTR-3B' not found").

**The workbook written** (`PQ_Extraction_Report_20261023_1324.xlsx`, in the run folder; sheets in this order: Summary, GSTR-3B Difference, GSTR-1 Difference, Turnover GSTR-1, Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B; each output sheet: A1 "`<sheet> Report`" merged A1:C1, headers on row 3, data from row 4):

| Sheet | ClientA-DL row | ClientB-MH row | ClientC-TN row |
|---|---|---|---|
| GSTR-3B Difference (Diff IGST/CGST/SGST) | ✓ 1101, 1102, 1103 | ✗ Missing ×3 | ✓ **Missing ×3** (sheet absent) |
| GSTR-1 Difference (Diff Taxable Value/IGST/CGST/SGST) | ✓ 8101, 8102, 8103, 8104 | ✗ Missing ×4 (**not ticked**) | ✓ 28101, `'12,345'` (text, no number format), Missing, 0 (no number format) |
| Turnover GSTR-1 (Taxable Value) | ✓ 9101 | ✗ Missing | ✓ 29101 |
| Reco Diff (IGST/CGST/SGST Diff) | ✓ 2101, 2102, 2103 | ✗ Missing ×3 | ✓ 22101, 22102, 22103 |
| Table 6J Match | ✓ 3101 | ✗ Missing | ✓ 23101 |
| Table 8D Limits | ✓ 4101 | ✗ Missing | ✓ 24101 |
| **IMS Reco (IGST/CGST/SGST)** — configured AA1/AD1/AG1 = 5101/5102/5103 | ✓ **1101, 1102, 1103** | ✗ Missing ×3 | ✓ 25101, 25102, 25103 |
| **Extra In Purchase** — configured O1/P1/Q1 = 6101/6102/6103 | ✓ **1101, 1102, 1103** | ✗ Missing ×3 | ✓ 26101, 26102, 26103 |
| **Extra In GSTR2B** — configured S1/T1/U1 = 7101/7102/7103 | ✓ **1101, 1102, 1103** | ✗ Missing ×3 | ✓ 27101, 27102, 27103 |

All whole numbers other than 0 carried the format `#,##0.00`. **Summary:** "Total Clients: 3", "ITC Reports Processed: 2/3", "Sales Reports Processed: 2/3" (ClientB's untouched Sales counts the same as a failure), then the nine sheet names.

**What the rows show.** For ClientA the three single-word sheets carry the "Diff GSTR-3B" figures 1101/1102/1103 instead of their configured cells; for ClientC — whose report lacks "Diff GSTR-3B" — the same sheets carry their configured cells, because the first matching field was `None` and the exact-name fallback was used. The same configuration therefore produces right or wrong figures depending on whether another sheet exists. (RAN by LT §4.6 a second way: with the extracted entries fed in reverse order, the three sheets showed the "Extra In GSTR2B" figures instead.)

**One unusual value** (same script, last part). ClientA's `Reco_Diff_IGST_Diff` was replaced by a time-zone-aware date (what recent pywin32 versions may return for a date cell, MEDIUM). `create_report` raised `TypeError: Excel does not support timezones in datetimes…`, and **a 4,579-byte file with the report's normal name was left behind that is not a valid Excel file** (openpyxl: "BadZipFile: File is not a zip file") (RAN). In the program the worker would write only "  Report error: Excel does not support timezones…" to the log; all figures of the run would be lost except in `run_history.json` (B-F25).

**Same minute, second run** (RAN `scripts/t22_extract.py` part 3, `out2/t22_extract.txt`): two reports created within one minute had the same name; only the second remained (its A4 held the second run's client).

#### Traced output versus what the feature claims

1. **9 of 22 figure columns show a different cell from the one configured** ("IMS Reco", "Extra In Purchase", "Extra In GSTR2B" ← "Diff GSTR-3B" J15/K15/L15) whenever those cells hold a value (RAN) — B-G1 in Section 7. [GST RULE] ⚠: which cell should feed these columns is the author's decision; the defect is only that the output does not follow the configuration.
2. **A failed refresh check turns every figure into "Missing"** although it was read; **"not ticked", "not found" and "failed" all show ✗ and "Missing"**; a missing sheet shows **✓ with "Missing"** (RAN) — B-G10.
3. **"Reports Processed" counts refresh successes** — an unticked report counts as not processed (RAN Summary) — B-G10, B-G71.
4. **Values are written unchecked**: text in a figure cell stays text, a text starting with "=" becomes a formula, a COM error number becomes a figure (RAN in B-F22) — B-G11, B-G65.
5. **One unusual value loses the whole workbook and leaves an unopenable file under the report's name** (RAN) — B-G43.
6. **A run in the same minute overwrites the earlier workbook** (RAN) — B-G27.
7. Column A shows "`<client folder> / <version folder>`", not a client name; one row per ticked version, so several versions of one client appear as separate rows (READ `:94`; RAN in B-F18 Select All) — technical point, the author's call.

#### Older documents versus code

- README.md:176 "Generate consolidated Excel report" — true; README.md:178-204 (output structure) does not list `PQ_Extraction_Report_*.xlsx` or say that it lands in the run folder (DO R67).
- README.md:15, 25 "Extracts specific cell values … into consolidated sheets" — partly true: three sheets carry other cells (DO R8, R14).
- README.md:251 `'missing_data_text': 'Missing'` — true; README does not say that a failed refresh check, an unticked report and a missing report all show "Missing".

---
### B-F24: Extraction progress, on-screen log and Stop

#### What it does (plain English)

While Step 4 runs, the screen shows a progress bar, a status line under it and a dark log box with coloured lines; a **Stop** button asks the run to stop. The work itself runs in a background "thread" (a second line of work inside the program, so the window keeps responding; B-F34) and reports back through Qt "signals" (messages from one part of the program to another). At the end, the version list is rebuilt from disk (B-F18) and the status bar at the bottom of the window says "Extraction complete".

#### Where it lives

| What | File and lines |
|---|---|
| Progress bar 0-100, status line "Ready for extraction", log box (dark, Consolas) | `gui/views/extraction_view.py:57-75`, `:77-86`, `:89-102` |
| Start Extraction, Stop (theme name "danger", disabled at start), Clear Log | `extraction_view.py:254-270`, `:272-276`, `:278-280` |
| Start/Stop swap | `extraction_view.py:392-395`, wired at `gui/pyqt_main_window.py:155-156` (started → running), `:550` (completed → idle), `:564` (error → idle) |
| Progress values: `idx / total × 100` with "Processing i/N: <name>" before each version; 95 "Creating report..."; 100 "Extraction complete" | `gui/controllers/extraction_controller.py:65-67`, `:89`, `:97` |
| The bar takes the whole-number part of the percentage | `extraction_view.py:397-400` |
| Log lines from the worker (own lines with a level) | `extraction_controller.py:43-46` (`_log`), `:62`, `:68`, `:84` ("Result: …" success/error), `:88`, `:93`, `:95`, `:98` |
| Log lines from the refresh/extraction code — **always level "info"** | `power_query_extractor/core/report_processor.py:30-37` (96 `self._log(` calls, none passes a level; RAN grep) |
| How a line is shown: inserted as HTML `<span style="color:…">…</span><br>`; colours normal `#d4d4d4`, info `#569cd6`, success `#4ec9b0`, warning `#dcdcaa`, error `#f44747` | `extraction_view.py:402-416` |
| Stop: button → flag → checked only before the next version | `extraction_view.py:276` → `gui/pyqt_main_window.py:169-170` → `extraction_controller.py:167-170`; check at `:60-63` |
| End of run: Start enabled, tab dot green, status bar "Extraction complete", run history, tree rebuilt | `gui/pyqt_main_window.py:549-561` |
| Error of the whole run: message box "Extraction Error" | `extraction_controller.py:36-41` → `gui/pyqt_main_window.py:563-565` |
| Theme rule that makes Stop red, with no disabled look | `gui/theme/theme_manager.py:104-111` (light), `:287-294` (dark) |
| Rules | B-R112, B-R113, B-R137, B-R138, B-R139 |

#### Input → Output

- **Input:** the worker's progress and log signals.
- **Output:** bar value, status text, coloured log lines; at the end the status-bar text and a rebuilt tree. No message box on normal completion.

#### User interaction

The user presses **Start Extraction** (it greys out and Stop becomes active), watches the bar and the log, may press **Stop** or **Clear Log**. The rest of the window stays usable throughout (B-F34).

#### Origin

Created in **c258047** (07-02-2026, B only (after the split)); restyled in **7e09e5c** (08-02-2026, B only (after the split)); **b0e7d0e** (09-02-2026, B only (after the split)) replaced the emoji result lines by "OK/FAIL" texts and began collecting the lines for the log file (READ notes MW, VW). The shared Tkinter extractor window before the split had its own progress display (not traced; it was deleted from B in a62d3cc).

#### How the code decides, step by step (READ)

1. `extraction_started` → the window disables Start and enables Stop (`pyqt_main_window.py:155-156` → `extraction_view.py:392-395`). Nothing else is disabled.
2. Before each version the worker sends `idx / total × 100` and "Processing i/N: <client folder> / <version folder>" (`extraction_controller.py:65-67`): the bar shows 0 % while the first version runs and reaches 100 × (N−1)/N while the last one runs. Inside a version (about 39 s per report at the default wait, B-F19) the bar does not move.
3. Every line that the refresh/extraction code writes reaches the screen as "info" (blue), because `ReportProcessor._log` never passes a level (`report_processor.py:30-37`). Only the worker's own lines carry other levels: "Result: …" is green if either report succeeded, red otherwise (`extraction_controller.py:84`); "Extraction stopped by user" is yellow (`:62`); "Report saved" and "Extraction complete" are green (`:93`, `:98`); "Report error" is red (`:95`).
4. Each line is inserted as HTML (`extraction_view.py:414`). In HTML, leading spaces and line breaks inside the text are dropped and anything between `<` and `>` is taken as a tag and disappears. All lines sit in one growing text block, which Qt re-lays-out at every insertion (RAN by VW: 500 / 1,000 / 2,000 lines took 3.1 / 12.1 / 48.6 s of window time).
5. Stop sets `stop_requested = True` on the worker (`extraction_controller.py:167-170`). The flag is looked at only at the top of the loop (`:60-63`), so the version in progress — with all its fixed Excel waits — runs to the end; then "Extraction stopped by user" is logged, the consolidated workbook and the log file are still written for the versions done, and the run finishes **like a normal one** (`:86-107`).
6. On completion (`pyqt_main_window.py:549-561`): Start enabled, tab 4 dot green whatever happened, status bar "Extraction complete", results stored in run history (B-F25), tree rebuilt from disk (all ticks lost, B-F18). No message box and no count of failures.

#### Concrete trace

**Run 1 — three versions, no Stop** (RAN `scripts/t24b_progress.py` → `out2/t24b.txt`; real window and worker thread, stand-in COM, waits counted). Run folder with `ClientA-DL`, `ClientB-MH`, `ClientC-TN`, one ITC report each; ClientB's refresh check fails (BB2 empty). ITC Only → Start Extraction.

| What | Observed |
|---|---|
| Progress updates sent | (0.0, "Processing 1/3: ClientA-DL / Version-050126 0900"), (33.33, "Processing 2/3: ClientB-MH / …"), (66.67, "Processing 3/3: ClientC-TN / …"), (95.0, "Creating report..."), (100.0, "Extraction complete") |
| Bar values shown | 0, 33, 66, 95, 100 |
| Fixed waiting the run would really have taken | 117 s (3 reports × 39 s), during which the bar showed 0 %, 33 % and 66 % without moving |
| Lines sent to the log | 541 — 536 "info", 4 "success", 1 "error" (about 180 lines per report) |
| ClientB's failure lines as sent | `❌ Cell BB2 is empty (None)`, `⚠️ Refresh validation failed: …`, `⚠️ Refresh completed with errors: …` — all **info (blue)**; only `Result: ITC: FAIL` was red |
| A line as sent → as displayed | `'  ✅ Cell BB2 populated successfully: <Client A>'` → `✅ Cell BB2 populated successfully: ` (indent and the bracketed text gone) |
| Last line | `Extraction complete: 3/3 clients processed` (green) although one of the three failed |
| End state | status line "Extraction complete", status bar "Extraction complete", bar 100, **no message box** |

**Run 2 — Stop pressed during the first version** (RAN `scripts/t24_window.py` → `out2/t24_window.txt`). Three versions, ITC and Sales ticked (Select All). The worker is held inside its first Excel wait so the screen can be inspected; Stop is pressed by setting the same flag the Stop button sets, while ClientA's second report is being finished.

| Moment | Observed |
|---|---|
| While running | Start disabled, Stop enabled; the tree, the option boxes, Step 3's "Start Processing" and Step 2's "START" all still enabled |
| Start Extraction pressed again | nothing on screen; program log only: "Extraction already in progress" |
| Progress | (0, "Processing 1/3: ClientA-DL / …"), (95, "Creating report..."), (100, "Extraction complete") |
| Log (selected lines, level) | `Result: ITC: OK, Sales: OK` (success); `Extraction stopped by user` (warning); `Report saved: PQ_Extraction_Report_…xlsx` (success); `Extraction complete: 1/3 clients processed` (success) |
| Fixed waiting before Stop could act | 78 s (ClientA's two reports ran to the end) |
| After the run | status bar "Extraction complete", no box, **tab 4 dot green `#107c10`**, Start enabled, 0 ticks left after the rebuild; tree: ClientA "Refreshed (…)", ClientB and ClientC "Pending" |
| On-screen log, first characters | `Processing ClientA-DL / Version-050126 0900...` / `Creating Excel instance...` — every line flush left, no blank lines |

**Stop's look** (RAN `scripts/t28_theme.py` → `out2/t28_theme.txt`): the disabled Stop button drew the same two colours (`#d83b01` red, `#ffffff` white) as an enabled one, in light and in dark mode.

#### Traced output versus what the feature claims

1. **Error lines from the refresh are not red** — they arrive as "info" (RAN: 536 of 541 lines) — B-G71 in Section 7.
2. **The HTML log drops indentation, blank lines and any text in angle brackets** (RAN) and **slows down sharply as it grows** (about 180 lines per report; RAN timings by VW) — B-G65, B-G56.
3. **Stop takes effect only after the current version, and a stopped run ends like a normal one** ("Extraction complete", green tab dot) (RAN Run 2) — B-G70, B-G39.
4. **A disabled Stop looks the same as an enabled one** (RAN) — B-G70.
5. **No completion message and no failure count; "X/Y clients processed" counts versions attempted**, not succeeded (RAN Run 1: "3/3" with one failure) — B-G71.
6. **The bar stands still for the whole of each version** (RAN: 0 % for the first 39-78 s) — B-G57.
7. **The window stays fully clickable during the keystroke refresh**: a click on B's own window (Stop, a tab, a tick) can take the keyboard focus from Excel just when Alt+A, R, A are typed (READ `report_processor.py:216-249`; MEDIUM, not runnable here) — B-G62, B-G74.
8. **Processing (Step 3) can be started during an extraction** (RAN Run 2: button enabled; RAN in B-F34: it actually runs) — B-G74.
9. On a run error the status bar keeps its previous text (RAN in B-F29) — B-G71.

#### Older documents versus code

- README.md:172-176 lists the Step 4 actions but says nothing about Stop, the log or how long a run takes; README.md:279 ("Processing seems frozen → large files take time, check log") is the only hint — in Step 4 the bar does not move for roughly 39 s per report (READ + RAN arithmetic).

---

### B-F25: Extraction log, statistics and run history

#### What it does (plain English)

Three records of each Step 4 run are kept:

1. **A text log** `Extraction_Log_<YYYYMMDD_HHMM>.txt` in the run folder, holding every line that was sent to the Step 4 log box.
2. **Statistics** in the consolidated workbook's Summary sheet (B-F23): total versions, and how many ITC and Sales reports "were processed" (meaning: refreshed successfully).
3. **Run history**: every version's whole result (statuses and all figures) is added to `run_history.json` in the user's settings folder `~/.gst_organizer` (JSON is a plain-text format for structured data), under the key "`<client folder> / <version folder>`".

#### Where it lives

| What | File and lines |
|---|---|
| Every line sent to the screen is also collected | `gui/controllers/extraction_controller.py:34`, `:43-46` |
| The log file is written once, at the end of a run that did not crash | `extraction_controller.py:100-101`, `:109-121` — name `:112-113` (`%Y%m%d_%H%M`), UTF-8 (a text encoding that can store every character, ✓ and ❌ included) `:114`, header "PQ Extraction Log - <YYYY-MM-DD HH:MM:SS>" and 60 "=" `:115-116`, the lines `:117-118`, write error → program-log warning only `:120-121` |
| Summary sheet counts | `power_query_extractor/core/data_consolidator.py:208-235` (counts `:222-226`) |
| Run history: each result stored under `result['client']` | `gui/pyqt_main_window.py:554-558` → `models/run_history.py:73-95` → saved by `models/cache_manager.py:66-68` → `models/run_history.py:128-136` |
| Where the key comes from | `gui/views/extraction_view.py:384` ("`<client folder> / <version folder>`") → `power_query_extractor/core/report_processor.py:44` |
| What the "Last Run" column of Step 2 reads (processing key "`<client>-<state code>`") | `gui/views/validation_view.py:381-385` → `models/run_history.py:97-107` |
| Rules | B-R90, B-R119, B-R120, B-R122 |

#### Input → Output

- **Input:** the lines sent during the run; the list of results.
- **Output:** `<run folder>/Extraction_Log_<YYYYMMDD_HHMM>.txt`; the Summary sheet; new or updated records in `~/.gst_organizer/run_history.json`.

#### User interaction

None; all automatic. Nothing on screen reads the stored extraction records.

#### Origin

Run history: **c258047** (07-02-2026, B only (after the split)); already then keyed by `result['client']` (READ notes MW). Log file: **b0e7d0e** (09-02-2026, B only (after the split)) (READ `git show b0e7d0e -- gui/controllers/extraction_controller.py`: `_log_lines`, `_save_log_file` added). Summary sheet: **edac6a3** (31-08-2025, shared - before the split on 30-10-2025).

#### How the code decides, step by step (READ)

1. `ExtractionWorker._log` appends each message, without its level, to `_log_lines` and sends it to the screen (`extraction_controller.py:43-46`). Lines that the refresh code writes only to the program log (`logger.warning/error/info` without `_log`: 19 such calls in `report_processor.py`, plus 3 `logger.debug` calls at `:394, :440, :476` that are below the INFO level and are written nowhere — 22 `logger.` calls in all outside `_log` itself, RAN grep) — for example "Sheet '…' not found", "Excel busy, retrying", "⚠️ Error detected: …", "Clicking button: …" — are **not** collected.
2. After the consolidated workbook, `_save_log_file()` (`:101`) writes the file with mode `'w'` (a file of the same minute is replaced). It is reached only if `_process` got that far: a crash inside the run, or closing the window during the run, leaves no log file (RAN in B-F34).
3. The Summary counts (`data_consolidator.py:222-226`) count results whose refresh status is success; an unticked report counts as "not processed".
4. On completion the window calls `run_history.add_extraction_result(result['client'], result)` for each result (`pyqt_main_window.py:555-557`). That function (`run_history.py:73-95`) looks up the key; the key is "`<client folder> / <version folder>`", which no processing record ever uses (processing records use "`<client>-<state code>`", B-F17), so it creates a **new** record (`:86-95`). (The `client_key` field that the window puts into every job — the client folder name, `gui/pyqt_main_window.py:524` — is never read by the extraction code or by this step: RAN grep, no use of `client_key` in `power_query_extractor/`; the key used is the display name `result['client']`.) — or, for a version extracted before, **replaces the earlier extraction's data** in that record (`:81-85`). Either way it sets `pq_refreshed = True` without looking at the status (`:83`, `:91`). The whole result — statuses, BB2 text, every figure (twice), file paths, a time stamp turned into text — is stored (`pq_extraction_data`). The file is rewritten whole (`cache_manager.py:66-68`).
5. `client_history` is never trimmed (only the processing `runs` list is capped at 50, `run_history.py:69-71`); extraction adds no `runs` entry.

#### Concrete trace

**Run A — the stopped run of B-F24** (RAN `scripts/t24_window.py` → `out2/t24_window.txt`). Files in the run folder afterwards: `Extraction_Log_<YYYYMMDD_HHMM>.txt` and `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`. The log file's first lines: `PQ Extraction Log - <date time>`, a line of 60 "=", two empty lines, ` Processing ClientA-DL / Version-050126 0900...`, `  Creating Excel instance...`; 300 lines in all; it holds the figures, for example `      ✓ GSTR-3B_Difference_Diff_IGST (J15) = 1101` and `      ✓ IMS_Reco_IGST (AA1) = 5101`. `run_history.json` afterwards: `client_history` key `'ClientA-DL / Version-050126 0900'` with `pq_refreshed: True`, the ITC status "Validated - Client Name: <Client A>" and all figures stored; `runs: []`, `total_runs: 0`.

**Run B — three versions, one failure, an existing processing record** (RAN `scripts/t25_full.py` → `out2/t25_full.txt`). Before the run, `run_history.json` already holds the processing record `'ClientA-DL'` from 05-01-2026 09:00:05 (as Step 3 would have written it). ClientB's refresh check fails; ClientC's ITC report has no "T6 of R9" sheet.

| What | Observed |
|---|---|
| Program-log warning `⚠️ Sheet 'T6 of R9' not found` — in the log file? | **No** |
| Any level marker (error / warning) in the log file? | **No** |
| `client_history` keys after the run | `'ClientA-DL'` (processing record, unchanged: `pq_refreshed: False`), `'ClientA-DL / Version-050126 0900'`, `'ClientB-MH / Version-050126 0900'`, `'ClientC-TN / Version-050126 0900'` |
| ClientB's new record | `pq_refreshed: True` next to the stored status `{'success': False, 'error': 'Refresh validation failed: Cell BB2 is empty', …}` |
| Step 2 "Last Run" for `ClientA-DL` | still `2026-01-05 09:00:05` (the processing run; the extraction is invisible there) |
| Summary sheet | "Total Clients: 3", "ITC Reports Processed: 2/3", "Sales Reports Processed: 2/3" |

**Run C — window closed during an extraction** (RAN `scripts/t34c_close.py` → `out2/t34c.txt`, B-F34): no log file, no consolidated workbook and no run-history file were written.

#### Traced output versus what the feature claims

1. **Extraction records are stored under a key that no screen reads**, so Step 2's "Last Run" never reflects an extraction (RAN Run B) — B-G25 in Section 7. Even without the "/ <version>" part the client folder name can differ from the processing key (abbreviation, length cut; B-F4).
2. **`pq_refreshed` is True even when the refresh failed** (RAN Run B) — B-G25.
3. **The log file leaves out every warning that goes only to the program log, and carries no levels** (RAN Run B) — B-G44, B-G71.
4. **The log file is written only at the end**: a crash or closing the window loses it (RAN Run C) — B-G40.
5. **A run in the same minute replaces the earlier log file** (READ `extraction_controller.py:112-114`, mode `'w'`; the same holds for the workbook, RAN in B-F23) — B-G27.
6. **Client figures are copied into plain-text files outside the reports** — the run folder's log file, `run_history.json` in the user's profile (kept for ever; `client_history` never trimmed) — and into `gst_organizer.log` (B-F32) (RAN) — B-G64, B-G59.
7. **"Reports Processed" counts refresh successes**, not extractions; an unticked report counts as not processed (RAN) — B-G10.
8. A second extraction of the same version replaces the first one's data in its record (RAN by LT §5.2 #27); nothing reads `pq_extraction_data` at all (READ grep) — B-G25.

#### Older documents versus code

- README.md:289-291 names the logs `gst_organizer.log` and `pq_extractor.log`; B never writes `pq_extractor.log` (B-F31) and README does not mention `Extraction_Log_*.txt` (DO R94, R95).
- The commit message of c258047 says "Run history displayed in client tree ('Last Run' column)" — only processing runs are ever displayed (RAN).

---
### B-F26: Automatic closing of Excel error dialogs

#### What it does (plain English)

While waiting for a refresh, B looks over **every window on the whole Windows desktop** for something that looks like an error dialog (a small pop-up box with buttons). If it finds one, it records the dialog's text as the refresh error and closes the dialog by clicking a button whose caption contains "ok", "close", "yes" or "continue" — or, if there is none, by sending the window a "close" message. A dialog found this way makes the refresh count as failed (B-F20), but the copy is still saved (B-F19).

#### Where it lives

| What | File and lines |
|---|---|
| When it looks: up to 5 checks 1 s apart after the user's wait time; one more after the extended wait | `power_query_extractor/core/report_processor.py:260-278`, `:288-298` |
| What counts as an error dialog (`_find_error_window`): every visible, enabled top-level window (`win32gui.EnumWindows`); its title, lower-cased, contains one of `microsoft excel`, `error`, `problem`, `cannot`, `failed`, `unable`, `could not`, `exception`, `invalid`, `missing`; **and** its window class is `#32770` (the standard Windows dialog), `NUIDialog` or `bosa_sdm_msword`; the first such window is taken | `report_processor.py:368-404` (words `:376-387`, classes `:392`, first `:402`) |
| Its text (`_get_window_text`): the title plus every child "static"/"text" control longer than 5 characters, joined, spaces collapsed, cut to 200 characters; nothing → "Power Query error occurred" | `report_processor.py:406-441` |
| How it is closed (`_close_error_window`): the first child whose class contains "button" and whose caption, lower-cased, **contains** `ok`, `close`, `yes` or `continue` is clicked (`BM_CLICK`); none → `WM_CLOSE` posted; then 0.5 s | `report_processor.py:443-476` (captions `:457`, click `:459`, close `:469-471`) |
| What the user sees: "⚠️ Error detected: <text>" goes to the program log only; at the end of the refresh "⚠️ Refresh completed with errors: <text>" goes to the screen log | `report_processor.py:272`, `:295` (program log); `:321-322` (screen) |
| The text becomes the refresh error; the copy is saved anyway | `report_processor.py:269-274`, `:326-330`; save `:305` |
| Rules | B-R102, B-R103 |

#### Input → Output

- **Input:** the windows open on the desktop at the moment of each check (any program's).
- **Output:** the error text in the report's status (`'error'`), a click on one of that dialog's buttons or a close message to it, and log lines.

#### User interaction

None. It acts on its own, on any program's dialog that matches — the user only notices that a dialog vanished.

#### Origin

All of it is in the first commit **edac6a3** (31-08-2025, shared - before the split on 30-10-2025) (READ `git show edac6a3:power_query_extractor/core/report_processor.py`, `_find_error_window` at line 352, `EnumWindows` at 385, `_close_error_window` at 427). Since the split only bare `except:` lines were narrowed to `except Exception:` in **c258047** (07-02-2026, B only (after the split)) (READ `git diff 47a1492 B/main`).

#### How the code decides, step by step (READ)

1. After the keystrokes and the user's wait (B-F19), up to five times, one second apart: `_find_error_window()` (`:266-275`).
2. It asks Windows for every top-level window of every program (`EnumWindows`, `:401`) and keeps those that are visible, enabled, whose lower-cased title contains one of the ten words **and** whose class is one of the three dialog classes (`:372-393`). A document window such as "Error log.txt - Notepad" is not a dialog class and is ignored; a dialog of another program titled "Backup failed" or any box titled "Microsoft Excel" qualifies.
3. The first qualifying window is used (`:402`) — in whatever order Windows lists them, not necessarily B's Excel's.
4. Its text is read (`:271`), logged as a warning to the program log only (`:272`), and the dialog is closed (`:273`): the child windows are walked and the first button whose caption contains "ok", "close", "yes" or "continue" gets a click (`:451-466`). "Yes" therefore answers whatever question the box asked; "ok" also matches captions such as "Look up…" or "Book…"; a box with only "Save / Don't Save / Cancel" gets a close message instead (`:469-471`).
5. Checking stops after the first hit (`:274`). After the extended wait, one last check runs if nothing was found yet (`:288-298`).
6. The refresh continues: BB2 check (B-F20), **save** (`:305`), close. The returned status is failure with `error` = the dialog text, but it still carries the BB2 message, e.g. "Validated - Client Name: …" (`:326-330`).

#### Concrete trace (RAN `scripts/t26_dialogs.py` → `out2/t26_dialogs.txt`; B's own `_refresh_power_query_simple` with stand-in Windows functions that present a made-up list of desktop windows; this proves only B's Python logic, not real Windows)

One refreshed copy of `<Client A>`'s ITC report (BB2 filled). Five desktop situations:

| Situation (windows on the desktop, in list order) | Window B picked | What B clicked or sent | Status returned | Copy saved? |
|---|---|---|---|---|
| A: Outlook main window; **another program's** dialog "Backup failed" (text "The nightly backup could not reach the server.", button OK); then a real Excel query error dialog | **"Backup failed"** (first in the list) | clicked **OK** on the backup program's dialog; the Excel error dialog was left open | failure, error "Backup failed The nightly backup could not reach the server." — validation message "Validated - Client Name: <Client A>" | yes |
| B: an Excel **question** box titled "Microsoft Excel": "Do you want to replace the existing file?" with &Yes / &No | that box | clicked **&Yes** | failure, error "Microsoft Excel Do you want to replace the existing file?" | yes |
| C: dialog "Cannot find the source" with buttons "Look up..." and "Cancel" | that dialog | clicked **"Look up..."** (its caption contains "ok") | failure, error "Cannot find the source" | yes |
| D: a Notepad window titled "Error log 2025.txt - Notepad" (not a dialog class) | none | nothing | **success** ("✓ No errors detected", "✓ No errors found") | yes |
| E: "Microsoft Excel" box with Save / Don't Save / Cancel | that box | no matching button → **WM_CLOSE** sent | failure, error "Microsoft Excel" | yes |

In every case with a dialog (A, B, C, E) the program log received "⚠️ Error detected: <text>" and "Clicking button: …" / "No button found, sending close message"; the screen log showed only "Checking for error dialogs…", "Error check 1/5…", "Final error check…" and, at the end, "⚠️ Refresh completed with errors: <text>" (RAN again in this check, same lines).

What real Windows and Excel would add (READ, cannot run here): whether a Power Query error appears as a `#32770` box or inside Excel's own window (then it is not found and the refresh counts as successful if BB2 is filled) — MEDIUM; whether pressing a button by `BM_CLICK` works on Excel's modern dialogs (`NUIDialog`) — LOW-MEDIUM. How to check: on a Windows test PC, break one query in a test template (e.g. point it at a missing file), run Step 4 on a dummy client and watch which box appears and whether B closes it.

#### Traced output versus what the feature claims

The feature is meant to catch Excel's refresh errors. The trace shows:

1. **It looks at every program's windows and can close another program's dialog**, possibly leaving the real Excel error open (RAN, situation A) — B-G61 in Section 7.
2. **It answers questions with "Yes"** and clicks any button whose caption merely contains "ok" (RAN, situations B and C) — B-G61.
3. **Any box titled "Microsoft Excel" counts as a refresh error**, even a harmless question (RAN, situation B) — B-G61.
4. **The copy is saved whatever the dialog said** (RAN, all situations) — B-G6.
5. **The failed status keeps the message "Validated - Client Name: …"**, which then sits next to the error in `run_history.json` (RAN) — B-F20 item 8.
6. Correction to the registry wording: the dialog text does **not** go to the program log only — the immediate "⚠️ Error detected" line does, but the same text reaches the screen and `Extraction_Log_*.txt` at the end as "⚠️ Refresh completed with errors: <text>" (RAN; `report_processor.py:322`).
7. A dialog that keeps re-appearing, or a modal prompt that is not matched (for example a credentials prompt), is not handled; COM calls then wait without any time limit (READ; MEDIUM) — B-G39.

#### Older documents versus code

- README.md does not mention that B closes dialogs on its own (READ README.md:164-176, 255-291).

---

### B-F27: Settings cache

#### What it does (plain English)

B remembers the user's Step 1 choices (source folder, the two templates, target folder, processing mode, the "client name in brackets" tick and its length) and the Step 4 options (wait time, file suffix, Skip refresh), plus the dark-mode choice, in a per-user file `settings.json`. It also writes a second file `session_state.json` (recent source folders, the last open tab and the window size) whose contents nothing ever uses: B opens it only to merge new values in before rewriting it (`models/cache_manager.py:46`, `:78`), and the two readers `load_session_state` / `get_recent_source_folders` (`:82-89`) are never called. Both live in the folder `.gst_organizer` inside the user's home folder (on Windows normally `C:\Users\<user>\.gst_organizer`).

#### Where it lives

| What | File and lines |
|---|---|
| Settings folder fixed when the module loads; created whenever the window opens | `models/cache_manager.py:16-17`, `:31-37` |
| The 11 settings and their defaults (`AppSettings`) | `models/app_state.py:10-46` — `source_folder`, `itc_template`, `sales_template`, `target_folder` (`''`), `processing_mode` (`'fresh'`), `include_client_name_in_folders` (False), `client_name_max_length` (35), `dark_mode` (False), `pq_wait_time` (10), `pq_suffix_pattern` (`'_Refreshed_{timestamp}'`), `pq_skip_refresh` (False) |
| Save: settings → JSON (indent 2, UTF-8, non-JSON values turned into text); the recent-folder list updated only when the source folder is new | `models/cache_manager.py:41-55`, `:135-142` |
| Load: unknown keys dropped, missing keys defaulted, **no type checks**; unreadable JSON → defaults | `models/cache_manager.py:57-62`, `:144-152`; `models/app_state.py:41-46` |
| Session file written on close; readers never called | `models/cache_manager.py:76-89`; `gui/pyqt_main_window.py:604-607` |
| When B saves: File > Save Settings, every scan, end of processing, dark-mode toggle, window close — **not** at Start (Step 3) or Start Extraction | `gui/pyqt_main_window.py:101-103`, `:227`, `:381`, `:597`, `:603`; `_save_settings` `:210-214` (status bar "Settings saved" for 3 s) |
| What is saved = what the screen shows at that moment | `gui/pyqt_main_window.py:194-208` |
| Start-up: load, then fill the screens | `gui/pyqt_main_window.py:50-51`, `:174-190` |
| Rules | B-R115, B-R116, B-R117 |

#### Input → Output

- **Input:** the values in the Step 1 and Step 4 widgets and the dark-mode choice.
- **Output:** `~/.gst_organizer/settings.json` (11 keys) and `~/.gst_organizer/session_state.json`; at the next start, the widgets filled from `settings.json`.

#### User interaction

Automatic, plus File > Save Settings. The user never sees the files.

#### Origin

**c258047** (07-02-2026, B only (after the split)) created `models/cache_manager.py` and `models/app_state.py`, replacing the v3 Tkinter cache `~/.gst_organizer_cache.json` (which B imports once, B-F2) (READ notes PQ; `git show c258047 --stat`).

#### How the code decides, step by step (READ)

1. At window start the settings folder is created (`cache_manager.py:33`) — before anything is saved.
2. `load_settings` reads `settings.json`; a missing or unreadable file gives the defaults (a warning goes to the program log, `:150-151`); a file that is valid JSON but not a key/value object (for example a list) makes `from_dict` fail (`app_state.py:45`) and the window is not built.
3. Each value is put straight into its widget (`pyqt_main_window.py:174-187`). A number widget given text (for example `pq_wait_time: "10"`) raises an error and the window is not built; an out-of-range number is clamped by the widget (999 → 120).
4. `_save_settings` collects every widget's current value (`:194-208`) and overwrites `settings.json` in place (`cache_manager.py:139-140`) — not via a temporary file, so a crash mid-write can leave a truncated file, which the next start treats as unreadable (defaults) and the next save overwrites. Write errors go to the program log only (`:141-142`).
5. On close, `session_state.json` gets `active_tab` and `window_size` (`pyqt_main_window.py:604-607`); `load_session_state` and `get_recent_source_folders` are never called (READ grep), so the window always opens on Step 1 at 1400 × 900.

#### Concrete trace (RAN `scripts/t27_settings.py` → `out2/t27_settings.txt`; offscreen window, home folder redirected)

| # | Action | Observed |
|---|---|---|
| 1 | open the window for the first time | `.gst_organizer` folder created, empty |
| 2 | type a target folder into Step 1's box (not saved), press Step 4 "Refresh List" | "No target folder configured." — Step 4 uses the **saved** value (B-F18) |
| 3 | set wait 30, suffix `_PQ_{timestamp}`, tick Skip; File > Save Settings | `settings.json` = the 11 keys with `pq_wait_time: 30`, `pq_suffix_pattern: "_PQ_{timestamp}"`, `pq_skip_refresh: true`; no `session_state.json` yet (the source folder is empty, so no recent-folder entry) |
| 4 | switch to Step 4, resize to 1600 × 1000, close | `session_state.json` = `{"active_tab": 3, "window_size": [1600, 1000]}` |
| 5 | reopen | tab **0**, size **1400 × 900** (session file ignored); wait 30, suffix `_PQ_{timestamp}`, Skip ticked (restored) |
| 6 | two windows open at once; one saves wait 45, the other then saves 60 | the file holds **60** — the last save wins |
| 7 | `pq_wait_time` hand-edited to the text `"10"` | window **not built**: `TypeError: setValue(self, val: int): argument 1 has unexpected type 'str'` (through `main.py` this ends the program with "Critical error", B-F1) |
| 8 | `pq_wait_time: 999` | window built, box shows 120 |
| 9 | file contains a JSON list `[1, 2]` | window **not built**: `AttributeError: 'list' object has no attribute 'items'` |
| 10 | file cut short (invalid JSON) | program log "Failed to load settings.json: Unterminated string…"; defaults used; the next save overwrites the damaged file (old values lost) |
| 11 | unknown key plus missing keys | unknown key dropped, missing ones defaulted, target `X:/somewhere` kept |

#### Traced output versus what the feature claims

1. **A wrongly typed value or a non-object file stops the window from opening** (RAN #7, #9) — B-G36 in Section 7.
2. **A damaged file silently resets every setting**, and the next save makes that permanent (RAN #10) — B-G36.
3. **The session file is write-only** (in effect: it is opened only to be merged and rewritten, never to use its contents): last tab and window size are never restored, the recent-folder list is never shown (RAN #4-5). This is dead (write-only) state (registry B-R117; noted in Section 7's writer/reader row, 7.7) rather than a numbered gap; listed here for Section 2's dead-code list.
4. **Options in use are not saved when a job starts**; a crash or a closed window during processing or extraction can lose them (READ `pyqt_main_window.py:345-371`, `:506-547` save nothing; closing does save, `:603`, but an abort inside a handler skips it, B-G35). Section 7 has no gap of its own for this; it is flagged for Section 7 in this part's return value.
5. **Step 4 reads the saved target, not the one on screen** (RAN #2) — B-G28.
6. **Two copies of the program overwrite each other's settings** (RAN #6) — B-G82.
7. **The folder is created at every start**; a file of that name stops the start (RAN by LT §5.2 #23: `FileExistsError` in the constructor, `cache_manager.py:33`), and a home folder that cannot be written would stop it the same way (READ `cache_manager.py:33`, no `try`; HIGH) — B-G37.
8. Not atomic writes (READ `cache_manager.py:135-142`) — B-G36.

#### Comparison with Standard 13 (facts only)

| Standard 13 (rebuild_pipeline_FINAL.md) | B |
|---|---|
| "Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`), never in the shared app folder" (lines 1039-1040) | per-user, but in `%USERPROFILE%\.gst_organizer\` (`cache_manager.py:17`), a dot-folder in the profile root |
| Passwords never in files (lines 1041-1045) | none stored |

#### Older documents versus code

- README.md:27 "Session Memory: remembers last used folders and settings" — true for settings; the "last used folders" list is kept but never shown (DO R16).
- README.md:171 "The extractor auto-loads target folder from cache" — true in B through `settings.json`, with the defects of B-F18.
- The class docstring "Stores … Session state (selected clients, active tab)" (`cache_manager.py:21-29`) — selected clients are never stored; the active tab is stored but never used (READ).

---
### B-F28: Dark mode / theme

#### What it does (plain English)

The menu item **View > Dark Mode** (a tick) switches the whole window between a light and a dark look and remembers the choice for the next start. The look is set by one "style sheet" for the whole program (QSS: a CSS-like text that tells Qt the colours and fonts of every kind of widget).

#### Where it lives

| What | File and lines |
|---|---|
| Menu item, ticked from the saved setting | `gui/pyqt_main_window.py:111-118` |
| Toggle: store the choice, apply the sheet, save settings | `gui/pyqt_main_window.py:594-597` |
| Applied at start-up | `gui/pyqt_main_window.py:44`, `:66` |
| `ThemeManager.apply_theme`: puts the light or dark sheet on the whole application | `gui/theme/theme_manager.py:408-419` |
| The two sheets (fonts "Segoe UI" 10 pt, "Consolas" 9 pt for logs) | `theme_manager.py:25-199` (light), `:201-405` (dark) |
| Red "danger" buttons (both Stop buttons) — no disabled variant | `theme_manager.py:104-111`, `:287-294` |
| Never used: the colour table `COLORS`, the `#info` button rules, `toggle()` | `theme_manager.py:10-22`, `:136-144` and `:319-327`, `:421-424` (RAN grep: no reader, no widget named "info", no caller) |
| Why parts stay light: widgets carry their own style sheets, which win over the application sheet — e.g. the content area of every coloured section (`ColoredSection`, the white panels of Step 1 — not the fold-away cards of B-F33, whose content area takes the card's own light colour, `gui/views/setup_view.py:62`) has `QWidget { background: white; }` | `gui/widgets/colored_section.py:62`; Step 4 status line colour `#323130` `gui/views/extraction_view.py:77-86`; Step 4 bar background `#f0f0f0` `extraction_view.py:61-74`; Step 3 status line `gui/views/processing_view.py:44-53`; Step 1 and 2 column backgrounds `gui/views/setup_view.py:101-111`, `gui/views/validation_view.py:31-44` |
| Stored as | `dark_mode` in `settings.json` (`models/app_state.py:20`, B-F27) |
| Rules | B-R132, B-R144 |

#### Input → Output

- **Input:** the tick.
- **Output:** the application style sheet; `"dark_mode": true/false` saved in `settings.json`.

#### User interaction

View > Dark Mode. Takes effect at once; restored at the next start.

#### Origin

**c258047** (07-02-2026, B only (after the split)) introduced the two style sheets, replacing the v3 Tkinter "dark mode manager" that walked every widget (212 lines, deleted in a62d3cc, B only (after the split)). **7e09e5c** (08-02-2026, B only (after the split)) added `COLORS` and the coloured button rules, and gave the Setup and Validation screens their own light colours (READ notes MW, VW; `git show 7e09e5c --stat`).

#### How the code decides, step by step (READ)

1. The menu action's `triggered` signal passes the new tick state to `_toggle_dark_mode` (`pyqt_main_window.py:117`, `:594`).
2. `self.settings.dark_mode` is set, `apply_theme` puts `DARK_STYLESHEET` or `LIGHT_STYLESHEET` on the running application (`theme_manager.py:414-419`), and `_save_settings()` writes all settings (`:597`).
3. Qt then redraws every widget. A widget that has its own style sheet keeps the colours written there (its own sheet is more specific than the application's), so only widgets without their own colours turn dark.

#### Concrete trace (RAN `scripts/t28_theme.py` → `out2/t28_theme.txt`; the window drawn offscreen and the two most common pixel colours of each widget counted)

| Widget | Light mode | Dark mode | Reading |
|---|---|---|---|
| Step 4 status line "Ready for extraction" | `#efefef` background, `#323130` text | **`#1e1e1e` background, `#323130` text** | dark grey on near-black: nearly unreadable |
| Step 3 status line | `#efefef`, `#323130` | **`#1e1e1e`, `#323130`** | same |
| Step 4 progress bar | `#f0f0f0`, `#cccccc` | `#f0f0f0`, `#cccccc` | stays light (its percentage text turns white in dark mode, RAN by VW: nearly invisible) |
| Step 4 "Skip refresh" tick box | `#efefef`, `#000000` | `#1e1e1e`, `#d4d4d4` | follows the theme |
| Step 1 source-folder box | `#ffffff`, `#cccccc` | `#ffffff`, `#3c3c3c` | stays white |
| Step 1 "Browse" (source) button | only `#ffffff` | only `#ffffff` | invisible in both modes (white text on white, B-F3) |
| Step 2 scan-summary box | `#f8f9fa`, `#cccccc` | `#f8f9fa`, `#cccccc` | stays light |
| Step 4 Stop button, **disabled** | `#d83b01`, `#ffffff` | `#d83b01`, `#ffffff` | the same as Stop **enabled** (RAN in dark mode: `#d83b01`, `#ffffff`) |

After the toggle `settings.json` held `"dark_mode": true`; a newly opened window had the menu ticked and the dark sheet applied (RAN).

#### Traced output versus what the feature claims

1. **Dark mode is only partly dark**: Steps 1 and 2 stay light, and the Step 3 and Step 4 status lines become dark text on a near-black background (RAN) — B-G69 in Section 7.
2. **A disabled Stop looks enabled** in both modes (RAN) — B-G70.
3. **Dead parts**: `COLORS`, the `#info` rules and `toggle()` are never used (RAN grep) — registry B-R132 (dead code, Section 2).
4. An error inside the toggle ends the program, because window handlers have no error hook (B-F1) — B-G35. Example: a `session_state.json` that is valid JSON but not a key/value object (for example the list `[1, 2]`) makes `session.get(...)` at `models/cache_manager.py:47` raise inside the settings save (RAN in this check, `/root/merge-scratch/run/VERIFY-B-S3b/scripts2/v_toggle.py`: View > Dark Mode → `AttributeError: 'list' object has no attribute 'get'`, program ended with exit code 134 — the exit code is the number a program hands back to Windows when it ends; 0 means a normal end; 134 is how this Linux machine reports an aborted program — on Windows the number would differ). A `session_state.json` that is not valid JSON at all does **not** do this: `_load_json` catches it and treats the file as empty (`cache_manager.py:144-152`; RAN `scripts2/v_session.py`: save succeeded).

#### Older documents versus code

- README.md:26 "Dark Mode: toggle between light and dark themes" — the toggle exists and is remembered; the result is only partly dark (DO R15).
- The module docstring "Uses QSS stylesheets instead of recursive widget walking" (`theme_manager.py:1-4`) is accurate; it explains why widgets with their own colours are not reached.

---

### B-F29: Step-status dots on tabs and automatic tab switching

#### What it does (plain English)

Each of the four tabs can carry a small coloured dot: orange = "ready to do next", green = "done". B also moves the user to the next tab by itself: to Step 2 after a scan, to Step 3 for a dry run or a processing start, to Step 4 after processing.

#### Where it lives

| What | File and lines |
|---|---|
| The dot: a 12-pixel filled circle | `gui/pyqt_main_window.py:569-580` |
| Colours: pending grey `#a0a0a0` (never used), ready orange `#FF8C00`, complete green `#107C10`; unknown → grey | `pyqt_main_window.py:582-590` |
| After a scan: Step 1 green, Step 2 orange, switch to Step 2 | `pyqt_main_window.py:304-307` |
| Dry Run: switch to Step 3 | `pyqt_main_window.py:332` |
| Start processing: switch to Step 3 | `pyqt_main_window.py:364` |
| After processing: Step 3 green, Step 4 orange; switch to Step 4 after the message box | `pyqt_main_window.py:390-392`, `:401` |
| After extraction: Step 4 green | `pyqt_main_window.py:551` |
| Error handlers change no dot | `pyqt_main_window.py:318-322`, `:412-415`, `:563-565` |
| All callers of `_update_tab_state` (RAN grep): lines 305, 306, 391, 392, 551 — nothing ever sets Step 2 to "complete" or any tab back to "pending" | — |
| Rules | B-R140 |

#### Input → Output

- **Input:** the end of a scan, a dry run, a processing start or end, an extraction end.
- **Output:** tab icons and the current tab.

#### User interaction

None; automatic. The user sees the dots and is moved between tabs.

#### Origin

Tab switching after a scan and at processing start came with the PyQt5 window in **c258047** (07-02-2026, B only (after the split)); the dots and the move to Step 4 after processing were added in **b0e7d0e** (09-02-2026, B only (after the split)) (READ `git show b0e7d0e -- gui/pyqt_main_window.py`: the new function `_update_tab_state` plus five added calls of it, and the added `self.tabs.setCurrentIndex(3)`).

#### How the code decides, step by step (READ)

1. `_on_scan_completed` sets tab 0 "complete" and tab 1 "ready" and switches to tab 1 — whatever the scan found (`:304-307`).
2. `_on_dry_run` and `_on_start_processing` switch to tab 2 (`:332`, `:364`); no dot changes.
3. `_on_processing_completed` sets tab 2 "complete" and tab 3 "ready", shows the message box, then switches to tab 3 (`:390-401`) — also after Stop, and also when every client failed in its reports (the worker still reports "completed", B-F16).
4. `_on_extraction_completed` sets tab 3 "complete" (`:551`) — also after Stop and when every refresh failed.
5. The error handlers change neither dots nor tab (`:318-322`, `:412-415`, `:563-565`).

#### Concrete trace (RAN `scripts/t29_dots.py` → `out2/t29_dots.txt`; offscreen window; the scan is real — an **empty** source folder through the real scan worker — and the later events call the window's own handlers with the values a worker would send)

| Event (in order) | Dots for Step 1 / 2 / 3 / 4 | Current tab | Other effects |
|---|---|---|---|
| window opens | none / none / none / none | Step 1 | — |
| scan of an **empty** folder | **green** / orange / none / none | Step 2 | status bar "Scan complete: 0 clients, 0 files, 0.0 B" |
| processing **error** | green / orange / none / none | Step 2 | box "Processing Error" |
| processing "completed" with 0 of 0 clients | green / orange / **green** / orange | Step 4 | box "Processing Complete" |
| a later scan **error** | green / orange / green / orange (unchanged) | Step 4 | — |
| extraction "completed" with no results | green / orange / green / **green** | Step 4 | status bar "Extraction complete" |
| extraction **error** | unchanged | Step 4 | status bar still says "Extraction complete" |

Also RAN in B-F24 Run 2 and B-F20 Scenario C: after a **stopped** extraction and after an extraction in which one refresh failed, Step 4's dot turned green (`#107c10`).

#### Traced output versus what the feature claims

1. **Green means "the step ran", not "it succeeded"**: an empty scan, a processing run with no successful client and an extraction that failed or was stopped all turn the dot green (RAN).
2. **Dots are never reset**: a later failure leaves the earlier green dot; a new scan does not reset Steps 3 and 4 (RAN).
3. **Step 2's dot never turns green** — no code sets it to "complete" (RAN grep) — and the grey "pending" colour is never used.
4. **After an extraction error the status bar keeps the previous run's "Extraction complete"** (RAN).

None of these four has its own gap number in Section 7 (Section 7.8 row S4 mentions only that the dots are never reset, under B-G56); they are flagged for Section 7 in this part's return value (closest existing gap: B-G71, feedback gaps).

#### Older documents versus code

- README.md does not describe the dots or the automatic tab moves (READ README.md:135-176).

---
### B-F30: Standalone extractor launcher

#### What it does (plain English)

`launch_extractor.py` is a second start file. It opens the same main window as `main.py` (B-F1), but switched to the "Step 4: Extract & Refresh" tab, for users who only want to refresh and extract reports that Step 3 made earlier.

#### Where it lives

| What | File and lines |
|---|---|
| Docstring "Convenience launcher - opens the main app on the Extract & Refresh tab (Tab 4)." | `launch_extractor.py:1-3` |
| Adds its own folder to the **end** of Python's import path | `launch_extractor.py:8` |
| `main()`: import Qt and the window, create the application and the window, switch to tab index 3, show, run | `launch_extractor.py:11-20` (tab at `:18`) |
| Any error → `print(f"Failed to launch: {e}")`, exit code 1 | `launch_extractor.py:22-24` |
| Entry | `launch_extractor.py:27-28` |
| Rules | B-R127 |

What `main.py` does and this file does not (READ, both files in full):

| `main.py` | `launch_extractor.py` |
|---|---|
| re-wraps the console output as UTF-8 (`main.py:11-13`) | no |
| configures logging: `gst_organizer.log` + console, level INFO (`main.py:15-23`) | **no logging set-up at all** |
| puts the program folder **first** on the import path (`main.py:26-29`) | appends it **last** (`:8`) — harmless when started as `python launch_extractor.py`, because Python already puts the script's folder first (READ, MEDIUM) |
| sets the application name "GST File Organizer" and version "4.0" (`main.py:41-42`) | no |
| a separate message for a missing module, naming `pip install PyQt5 openpyxl` (`main.py:49-54`) | one generic "Failed to launch: …" |

#### Input → Output

- **Input:** none (the saved settings, B-F27).
- **Output:** the main window, opened on Step 4 — or a console line and exit code 1.

#### User interaction

The user runs `python launch_extractor.py` from a console (README.md:166-169). There is no shortcut or `.bat` file.

#### Origin

Created in **edac6a3** (31-08-2025, shared - before the split on 30-10-2025) as the launcher of the separate Tkinter extractor: it imported `PowerQueryExtractor` from `power_query_extractor/extractor_main.py` (B-F31) and showed a Tk error box "Failed to launch: …" on failure (READ `git show edac6a3:launch_extractor.py`). **c258047** (07-02-2026, B only (after the split)) rewrote it to open the unified PyQt5 window on tab 4; the error box became a `print` (READ notes PQ).

#### Concrete trace (RAN `scripts/t30_entries.py` → `out2/t30.txt`; each start in its own empty start folder, home folder redirected)

| Run | Exit code | Console | Files in the start folder | Settings folder |
|---|---|---|---|---|
| `python -B launch_extractor.py`, **no pywin32** (as on any PC where pywin32 is not installed) | 1 | `Failed to launch: No module named 'pythoncom'` | none | not created (the window was never built) |
| the same with the stand-in pywin32 modules, through a wrapper that closes the window instead of waiting for a user | 0 | (wrapper) window title "GST File Organizer v4.0", current tab **3** "Step 4: Extract && Refresh" (shown on screen as "Step 4: Extract & Refresh") | **none — no log file** | `settings.json`, `session_state.json` (written on close) |

In the second run the window's start-up message "GST File Organizer v4.0 initialized" and two typical refresh lines that the wrapper made B's own `ReportProcessor._log` write (one containing "✓") appeared **nowhere** — not in a file and not on the console (RAN: stdout and stderr — the console's normal and error output — empty apart from the wrapper's line; RAN again in this check with the same result). On Windows with pywin32 installed, `core/excel_handler.py:26` calls `logging.info(...)` on the root logger at import time; with no logging set-up, Python then installs its own console handler at level WARNING (READ; general Python behaviour, HIGH) — so warnings and errors would go to the console, INFO messages would still be lost, and still no file is written.

#### Traced output versus what the feature claims

1. **Started this way, the program writes no log file at all** and loses every INFO message (RAN) — B-G44 in Section 7.
2. **Every start-up failure is only `print`ed**; without a console nothing is visible (RAN: the only trace of a missing pywin32 is one console line) — B-G34.
3. **It needs pywin32 exactly like `main.py`** (RAN) — B-G32.

#### Comparison with Standard 13 (facts only)

| Standard 13 (rebuild_pipeline_FINAL.md) | `launch_extractor.py` |
|---|---|
| users start the app through `Run [App Name].bat`, which runs `pythonw.exe -B -s main.py` with `start "" /D "%~dp0"` (lines 951-962) | started with `python launch_extractor.py` from a console; no `.bat`, no `pythonw` |
| `sys.excepthook` / `threading.excepthook` write an error file and show a message box (lines 995-1025) | none; `print` only |
| "There is no console: `print()` shows nothing. Report problems in the window, and log them to files in `%LOCALAPPDATA%\[App Name]\`" (lines 1035-1036) | no log file; errors printed |

#### Older documents versus code

- README.md:166-169 "After processing, run `python launch_extractor.py`" — it still works (with pywin32) but opens the main window, not a separate extractor (DO R61).
- README.md:283-287 tells the user to enable debug logging by editing `logging.basicConfig(…)` "in main.py or launch_extractor.py" — `launch_extractor.py` has no `basicConfig` (READ; DO R93).

---

### B-F31: Old stand-alone extractor entry point (extractor_main.py)

#### What it does (plain English)

`power_query_extractor/extractor_main.py` was the start file of the old separate Tkinter extractor window. In B that window's file has been deleted, so this start file **cannot even be loaded**; nothing in B calls it.

#### Where it lives

| What | File and lines |
|---|---|
| Import of the deleted window `power_query_extractor/gui/extractor_window.py` | `power_query_extractor/extractor_main.py:9` |
| Its own logging set-up: `pq_extractor.log` in the **current** folder | `extractor_main.py:12-19` |
| Class `PowerQueryExtractor`: remembers a target folder or looks one up | `extractor_main.py:24-69` |
| Cache look-up: `gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json` (all relative to the current folder); keys `target_folder`, `target_path`, `output_folder`, `last_target` | `extractor_main.py:36-60` |
| `run()`: Tk `mainloop()` | `extractor_main.py:62-69` |
| `main()` and entry | `extractor_main.py:72-79` |
| The now-empty package (one docstring line) | `power_query_extractor/gui/__init__.py` |
| Leftovers that still refer to it | `.gitignore:41` (`pq_extractor.log`), `.gitignore:63` (`config/last_session.json`) |
| Rules | B-R128 |

#### Input → Output

- **Input:** none. **Output:** an import error.

#### User interaction

None in the program. Someone could try to run it directly with `python -m power_query_extractor.extractor_main`; no document in B tells them to (README.md's extractor instructions, :164-176, name `launch_extractor.py`, B-F30).

#### Origin

Created in **edac6a3** (31-08-2025, shared - before the split on 30-10-2025); called by `launch_extractor.py` until **c258047** (07-02-2026, B only (after the split)) rewrote that launcher (B-F30); broken by **a62d3cc** (07-02-2026, B only (after the split)), which deleted `power_query_extractor/gui/extractor_window.py` (READ notes PQ; `git log B/main --oneline`: a62d3cc "Remove old Tkinter GUI files replaced by PyQt5 in v4.0"). Even before the split, its cache look-up read files in the **current** folder that the main program never wrote (the v3 main program wrote `~/.gst_organizer_cache.json` in the **home** folder) — a writer/reader mismatch that already existed at 47a1492 (READ notes PQ, DO R62).

#### Concrete trace (RAN `scripts/t30_entries.py` → `out2/t30.txt`)

| Run | Exit code | Error | Files created in the start folder |
|---|---|---|---|
| `python -B -m power_query_extractor.extractor_main` with the stand-in pywin32 modules | 1 | `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'` at `extractor_main.py:9` | none — **no `pq_extractor.log`** (line 12 is never reached) |
| the same without stand-ins | 1 | the same error (it fails before anything needs pywin32) | none |

No file in B imports or calls it (RAN grep for `extractor_main` over the whole project, README.md and the file itself included: no hit at all — not even README.md names it).

#### Traced output versus what the feature claims

1. **The entry point is broken and dead** (RAN) — B-G45 in Section 7. Its files (`pq_extractor.log`, the three cache names) are never created in B; the `.gitignore` lines for them are leftovers.

#### Older documents versus code

- README.md:91-94 shows `power_query_extractor/` with a `gui/` folder (now holding only a one-line `__init__.py`); README.md:164-171 describes a stand-alone extractor, started with `python launch_extractor.py`, that "auto-loads target folder from cache" — in B that launcher opens the main window instead (B-F30), whose Step 4 does load the saved target (B-F18, so the "auto-loads" sentence is true in B; DO R62), while the old stand-alone window and this file's own cache look-up no longer work; README.md:291 names `pq_extractor.log` as the extractor's log, which B never writes (DO R36, R95).

---

### B-F32: Application log file

#### What it does (plain English)

When the program is started with `main.py`, every program message of level INFO or higher (INFO = information; WARNING; ERROR) from every part of B is written to a text file **`gst_organizer.log` in the folder the program was started from**, and also to the console. The file is added to at every start and never trimmed.

#### Where it lives

| What | File and lines |
|---|---|
| Logging set-up: level INFO; format "`<time> - <module> - <level> - <message>`"; `FileHandler('gst_organizer.log')` (a **relative** name, so the current folder; no encoding given, so the PC's default text encoding); a console handler | `main.py:15-23` |
| UTF-8 wrapper for the console only (not for the file) | `main.py:11-13` |
| Every module's own logger feeds it: 13 modules call `logging.getLogger`, but one of them, `power_query_extractor/extractor_main.py:21`, can never be loaded (B-F31), so 12 can actually write | e.g. `power_query_extractor/core/report_processor.py:18`, `gui/pyqt_main_window.py:29` (RAN grep by DC P-03; re-counted in this check) |
| The refresh/extraction lines — all at INFO, including client folder names, BB2 text and every extracted figure | `report_processor.py:30-37` (`_log` → `logger.info`), `:541`, `:557-562` |
| Links values written in Step 3, first 50 characters | `core/excel_handler.py:165` |
| Direct calls on the root logger | `core/excel_handler.py:26`, `:28` (at import); `main.py:50`, `:57` |
| Not set up when started through `launch_extractor.py` | B-F30 |
| Rules | B-R121, B-R123 |

#### Input → Output

- **Input:** every program message of level INFO and above.
- **Output:** lines appended to `<current folder>/gst_organizer.log`, and the same lines on the console.

#### User interaction

None; README.md:289-290 and :333 tell the user to look at this file when something goes wrong.

#### Origin

The same set-up, with the same file name, is in **edac6a3** and at the split point **47a1492** (both shared - before the split on 30-10-2025) (READ `git show 47a1492:main.py`, lines 19-26). A real `gst_organizer.log` was committed in edac6a3 and deleted in cbfa4a5 (shared - before the split on 30-10-2025) — history only, its content was not opened. **c258047** (07-02-2026, B only (after the split)) rewrote `main.py` and kept the set-up.

#### Concrete trace (RAN `scripts/t30_entries.py` → `out2/t30.txt`; each start in its own empty start folder; the wrapper `scripts/t32_wrapper.py` runs `main.py` as the real program, makes B's own `ReportProcessor._log` write two typical refresh lines — "✓ File is accessible: ITC_Report_ClientA_DL_050126_0900.xlsx" and "Path length: 120 characters" — and closes the window)

| Run | Exit code | Where the log went | What the log holds |
|---|---|---|---|
| `python -B main.py`, **no pywin32** | 1 | `gst_organizer.log` (83 bytes) **in the start folder** | one line: `root - ERROR - Import Error: No module named 'pythoncom'`; the console also showed "Missing module: No module named 'pythoncom'", "Please install required packages:", "  pip install PyQt5 openpyxl" |
| `main.py` with the stand-ins | 0 | `gst_organizer.log` (341 bytes) in the start folder — **not** in the program folder (RAN: the program copy's folder held no log afterwards) | 3 lines: `gui.pyqt_main_window - INFO - GST File Organizer v4.0 initialized`; `power_query_extractor.core.report_processor - INFO - ✓ File is accessible: ITC_Report_ClientA_DL_050126_0900.xlsx`; `… - INFO - Path length: 120 characters` (the same three lines on the console) |
| the same, **simulating a Windows PC whose default text encoding is cp1252** (the wrapper gives the file handler `encoding='cp1252'` where B gives none; cp1252 is the usual Western-European Windows character set and has no "✓") | 0 | the start folder (199 bytes) | **only 2 lines: the "✓" line is missing**; the console's error channel showed "--- Logging error ---" and a traceback |
| `main.py` started a second time in the same folder | 0 | the same file | 6 lines — appended, never trimmed |
| `launch_extractor.py` (B-F30) | 0 | **nowhere** | — |

Whether a given office PC writes the file in cp1252 depends on its Windows language settings and on Python's UTF-8 mode, which is off by default in Python 3.14 on Windows (MEDIUM; RAN by DC T7 with the same simulation). How to check: on an office PC, run one Step 4 extraction and look in `gst_organizer.log` for lines that start with "✓"; if there are none but the screen showed them, lines are being lost.

#### Traced output versus what the feature claims

1. **The log goes to whatever folder the program was started from** (RAN) — with the Standard 13 launcher that would be the shared program folder, written by every user at once — B-G44 in Section 7; and if that folder cannot be written, the program stops before any message (RAN by the Section 7 writer) — B-G33.
2. **Lines containing symbols such as ✓ can be dropped silently on a cp1252 PC** (RAN, simulated) — B-G44.
3. **No file at all when started through `launch_extractor.py`** (RAN) — B-G44.
4. **Never trimmed** (RAN: appended at every start) — B-G44.
5. **It holds client folder names, BB2 text, Links paths and every extracted figure** in plain text (READ `report_processor.py:541, 557-562`; `core/excel_handler.py:165`) — B-G64.

#### Comparison with Standard 13 (facts only)

| Standard 13 (rebuild_pipeline_FINAL.md) | B |
|---|---|
| "Find the app's own files from `APP_DIR`, never from the current folder" (line 1034) | the log uses the current folder (`main.py:20`) |
| Log to files in `%LOCALAPPDATA%\[App Name]\`; settings and logs never in the shared app folder (lines 1035-1040) | the current folder, which the Standard's launcher sets to the app folder (`start "" /D "%~dp0"`, line 962) |
| "Never keep … any file several people write at the same time, in the shared folder" (lines 1046-1047) | started from a shared folder, every user would append to the same `gst_organizer.log` |
| An error hook writes `error_<date>.txt` under `%LOCALAPPDATA%\<App>\logs` (lines 997-1012) | none (B-F1) |

#### Older documents versus code

- README.md:289-290 "Logs saved to `gst_organizer.log` — main application" — true only for `main.py`, and the file lands in the start folder (DO R94).
- README.md:308 "All operations are logged" — not when started through `launch_extractor.py`; and the screen-only Step 4 lines are all at INFO, so error lines are not marked as errors in the file (READ `report_processor.py:30-37`).

---
### B-F33: Help texts, tooltips, scrolling and collapsible sections

#### What it does (plain English)

B's on-screen help is fixed text: a welcome banner and two fold-away "cards" on Step 1 (INSTRUCTIONS with seven steps, and EXPECTED FILE NAMES / MODES), a note under the "client name in brackets" option, and five instruction lines on Step 2. The left columns of Steps 1 and 2 scroll. Clicking a card's header folds or unfolds it, and both cards fold away by themselves at every scan to make room for the scan log. There are no tooltips (the small pop-up hints that appear when the mouse rests on a control).

[GST RULE] ⚠ — the EXPECTED FILE NAMES card shows the names of the GST return/data files the program expects (GSTR-2B Reco, IMS Reco, GSTR-3B, Sales, Sales Reco, Annual Report); return names and formats can change over time. Only the agreement between the card and the code is checked below; whether the list is right is the author's call.

#### Where it lives

| What | File and lines |
|---|---|
| The fold-away card widget (`CollapsibleCard`): header button "`  v  <title>`" when open, "`  >  <title>`" when folded | `gui/views/setup_view.py:17-82` (`_toggle` `:69-72`, header text `:74-76`; `addLayout` `:81-82` never called; `_bg`, `_border`, `_text_color` `:23-25` never read) |
| Step 1 layout: two columns 3:2, the left one scrolls | `setup_view.py:95-146` (scroll area `:107-111`, ratio `:143-144`) |
| Welcome banner "Welcome to GST File Organizer!" / "This tool helps you organize GST files and create Excel reports automatically. Let's get started!" (the frame's unscoped `QFrame { … }` rule also draws boxes around the two labels) | `setup_view.py:150-188` (rule `:153`, title `:162`, subtitle `:175-176`) |
| INSTRUCTIONS card: "1. Select folder with GST files" … "7. Start processing in Processing tab" | `setup_view.py:469-486` |
| EXPECTED FILE NAMES / MODES card: six name patterns and three mode lines | `setup_view.py:488-517` (patterns `:492-499`, modes `:510-512`) |
| Note "Names >10 chars will show warning. This overrides individual client settings if checked." | `setup_view.py:385-388` |
| Fold both cards (only folds; never unfolds) | `setup_view.py:590-594`, called at every scan from `gui/pyqt_main_window.py:228` |
| Step 2: scrolling left column; "VALIDATION INSTRUCTIONS" with five lines | `gui/views/validation_view.py:25-68` (scroll `:37-53`), `:72-115` (lines `:96-102`) |
| A second fold-away widget that nothing uses (`CollapsibleGroup`) | `gui/widgets/collapsible_group.py:9-42` (loaded at start through `gui/widgets/__init__.py`) |
| Tooltips | none (RAN grep `setToolTip`: 0 hits) |
| Rules | B-R141, B-R44 |

#### Input → Output

- **Input:** none (display only); a click on a card header; the start of a scan.
- **Output:** text on screen; cards folded or unfolded.

#### User interaction

Reading; clicking a card's header; scrolling the left column.

#### Origin

The cards, the banner, the coloured sections, the Step 2 instruction box and the note text came with the restyle in **7e09e5c** (08-02-2026, B only (after the split)); the note's wording was carried over from the Tkinter Setup tab at the split point (`gui/tabs/setup_tab.py:238-239@47a1492`, shared - before the split on 30-10-2025), where the per-client setting and the long-name warning it mentions still existed (READ notes VW). The card's mode lines kept the older wording of **c258047** (07-02-2026, B only (after the split)) while the radio buttons got new wording in 7e09e5c. **b0e7d0e** (09-02-2026, B only (after the split)) kept the two cards on the screen object and added the automatic folding at each scan.

#### How the code decides, step by step (READ)

1. At start the two cards are open (`_collapsed = False`, `setup_view.py:22`); `_update_header` writes "`  v  <title>`" (`:67`, `:74-76`).
2. A header click calls `_toggle` (`:57`, `:69-72`): the content is hidden or shown and the arrow changes.
3. At every scan `_on_scan_requested` calls `collapse_info_cards()` (`pyqt_main_window.py:228`) before the scan starts; it folds every card that is open (`setup_view.py:592-594`), including a card the user had just opened.
4. The texts are typed into the screen code by hand (`setup_view.py:473-481`, `:492-499`, `:510-512`; `validation_view.py:96-102`); they are not generated from the real name patterns (`utils/constants.py:13-62`) or from the mode logic (`core/file_organizer.py`).

What the help says, set against what the code does (READ; patterns RAN by LT §1):

| Help text | What the code does |
|---|---|
| Card: "GSTR-2B-Reco-Client-State-Period", "ImsReco-Client-State-DDMMYYYY", "GSTR3B-Client-State-Month", "Sales-Client-State-Start-End", "SalesReco-Client-State-Period", "AnnualReport-Client-State-Year" ⚠ | the six patterns agree in shape (B-R6 to B-R11), but the card does not say that `.xls` is also accepted, that the client and state parts may not contain a hyphen (a hyphen silently shifts the split, B-G17), that the IMS date must be exactly 8 digits (any 8 digits pass), that GSTR3B names may start with "(1) ", that Sales needs exactly four parts, or that `.xlsm` files are never recognised |
| Card: "Fresh: New timestamped structure" | agrees (new run folder each run) |
| Card: "Re-Run: Overwrite existing"; radio button: "Re-Run - Add new version folders to existing structure" | the two texts disagree; the code adds a new Version folder (nothing is overwritten except within the same minute) (B-F12) |
| Card: "Resume: Skip already processed"; radio button: "Resume - Continue previously interrupted processing" | the code also adds a new Version folder and copies everything again; it skips a file only when run again within the same clock minute (B-F12) |
| Note: "Names >10 chars will show warning. This overrides individual client settings if checked." | no such warning exists, and there are no individual client settings (B-F4, B-F5) |
| Step 2: "Check client list - green = complete, orange = missing files" | there is a third colour, red-orange, for "Has duplicates" (B-F8) |
| Step 2: "Click 'Start Processing' when ready" | the Step 2 button reads "START" ("Start Processing" is the Step 3 button) |

#### Concrete trace (RAN `scripts/t33_help.py` → `out2/t33_help.txt`; the window drawn offscreen at its opening size)

| What | Observed |
|---|---|
| Card headers at start | `  v  INSTRUCTIONS`, `  v  EXPECTED FILE NAMES / MODES` |
| INSTRUCTIONS lines | "1. Select folder with GST files" … "5. Click 'Scan Files'", "6. Review clients in Validation tab", "7. Start processing in Processing tab" |
| Card pattern and mode lines | the six patterns above, then "Fresh: New timestamped structure", "Re-Run: Overwrite existing", "Resume: Skip already processed" |
| Radio buttons | "Fresh - Create new folder structure (recommended for first time)", "Re-Run - Add new version folders to existing structure", "Resume - Continue previously interrupted processing" |
| Step 2 buttons | "DRY RUN", "EXPORT LIST", "START" |
| **Is SCAN FILES visible without scrolling?** | window 1400 × 900; the left column's visible height is 815 pixels; SCAN FILES starts at y = 989 → **no**, the user must scroll to find the main button |
| Click on the INSTRUCTIONS header | `  >  INSTRUCTIONS`, content hidden |
| After SCAN FILES | both cards folded |
| The user unfolds one card and scans again | both folded again |
| Widgets with a tooltip | **0 of 205** |
| `CollapsibleGroup` widgets in the window | 0 |

#### Traced output versus what the feature claims

1. **The help contradicts the code and itself** (Re-Run / Resume, the note about warnings and per-client settings, the colour legend, the button name) (RAN + READ) — B-G72 in Section 7.
2. **No tooltips anywhere** (RAN) — B-G72.
3. **The main SCAN FILES button is below the visible part of Step 1 at the opening window size** (RAN) — B-G68.
4. **The file-name card is typed by hand and leaves out rules that make real files fail** (READ; the rules RAN by LT) — B-G47, B-G17.
5. The cards fold again at every scan even when the user opened them (RAN) — a small usability point, part of B-G72.
6. Dead pieces: `CollapsibleGroup`, `CollapsibleCard.addLayout`, three unused card attributes (RAN grep) — Section 2 dead code.

#### Older documents versus code

- README.md:124-133 gives the same six naming patterns with examples; it is consistent with the card and has the same omissions (DO R43-R50).
- README.md:142-145 describes Re-Run as "Add new versions to existing structure" (matches the code) and Resume as "Continue interrupted processing" (the code does not do that; DO R53).
- README.md does not describe the cards, the banner or the Step 2 instructions.

---

### B-F34: Background jobs and run guards

#### What it does (plain English)

The three long jobs — scanning (Step 1), processing (Step 3) and extraction (Step 4) — each run in a background thread, so the window keeps responding. Each job's controller refuses to start a second job **of the same kind** while one is running; it tells only the program log, not the user. Jobs of **different** kinds are not stopped from running at the same time. Workers never touch the window directly; they send signals, which Qt delivers to the window's own thread.

#### Where it lives

| What | File and lines |
|---|---|
| Scan: refuse if running ("Scan already in progress", program log only), new thread + worker, connections, start, clean-up; **no Stop** | `gui/controllers/scan_controller.py:57-80` (guard `:59-61`), `:88-94` (no stop method: RAN grep) |
| Processing: the same, plus a Stop flag | `gui/controllers/processing_controller.py:256-279` (guard `:259-261`), `:281-284`, `:292-298` |
| Extraction: the same, plus a Stop flag | `gui/controllers/extraction_controller.py:138-165` (guard `:142-144`), `:167-170`, `:178-184` |
| A worker's own catch-all: error → signal → message box in the window | `scan_controller.py:24-41`, `processing_controller.py:35-40`, `extraction_controller.py:36-41` → `gui/pyqt_main_window.py:318-322`, `:412-415`, `:563-565` |
| What is disabled while a job runs: only that job's own Start (and Stop enabled), and SCAN FILES during a scan | `gui/views/processing_view.py:132-135`, `gui/views/extraction_view.py:392-395`, `gui/pyqt_main_window.py:231` |
| Window close: saves settings and session, accepts at once, no check for running jobs | `gui/pyqt_main_window.py:601-608` |
| Data shared between threads without locks: the Skip setting; the window's scan data, changed by the processing worker | `extraction_controller.py:54-55`; `core/file_organizer.py:216-229` via `processing_controller.py:30` |
| Rules | B-R111, B-R112 |

#### Input → Output

- **Input:** a request to start a scan, a processing run or an extraction.
- **Output:** a background job whose progress, log lines and result reach the window by signals — or, for a second request of the same kind, nothing but a program-log line.

#### User interaction

Automatic whenever a job is started (SCAN FILES / RE-SCAN; START / Start Processing; Start Extraction); Stop for processing and extraction; closing the window.

#### Origin

The QThread-worker pattern was introduced with the PyQt5 window in **c258047** (07-02-2026, B only (after the split)) (READ notes MW). Before the split the Tkinter program (shared - before the split on 30-10-2025) also ran processing in a background thread, but that thread changed window elements directly, which shared commit 6237597 only partly wrapped (READ notes SH; not traced here).

#### How the code decides, step by step (READ)

1. Start request → the controller checks `self._thread and self._thread.isRunning()`; if so, it writes a warning to the program log and returns (e.g. `extraction_controller.py:142-144`). The window has by then already done its own part — for example switched tab or, for a re-scan, cleared the scan log (`pyqt_main_window.py:227-231`).
2. Otherwise it creates a new `QThread` (Qt's object for one background thread) and a worker object, moves the worker into the thread, connects its signals — eight for extraction (`extraction_controller.py:155-163`) and processing (`processing_controller.py:269-277`), seven for the scan (`scan_controller.py:70-78`, which has no log signal) — and starts the thread (e.g. `extraction_controller.py:146-165`).
3. The worker's `run()` catches every error and sends it as an `error` signal (`extraction_controller.py:36-41`); the window shows "Scan Error" / "Processing Error" / "Extraction Error".
4. Nothing ties the three controllers together, so a scan or a processing run can start while an extraction runs, and the other way round. Processing and extraction both drive Excel; Step 3's report creation first attaches to any running Excel (`core/excel_handler.py:110-112`), which could be Step 4's visible one (READ, MEDIUM; B-F13).
5. Stop sets a flag that is read only between items (B-F16, B-F24); scanning has none.
6. Closing the window runs `closeEvent` (`pyqt_main_window.py:601-608`): it saves settings and the session and accepts; a running worker is neither stopped nor waited for.
7. When a thread finishes, `_cleanup` deletes whatever thread and worker the controller holds **at that moment** (`extraction_controller.py:178-184`); if a new job of the same kind was started in between, the new running thread is deleted — RAN by MW (e13, run on the scan controller, whose `_cleanup` `scan_controller.py:88-94` is identical; timing widened on purpose): "QThread: Destroyed while thread is still running", exit code 134 (LOW likelihood).

#### Concrete trace

**Run 1 — three jobs at once** (RAN `scripts/t34b_jobs.py` → `out2/t34b.txt`; offscreen window; the extraction worker is held inside its first Excel wait as if Excel were refreshing; Step 3 takes B's openpyxl path because this is not Windows). Target folder with one earlier run folder `Annual Statement-050126 0900` (one ITC report of `ClientA-DL`); a source folder with two files of an invented `ClientD` in Gujarat; two small templates.

| Step | What the user does | What happened |
|---|---|---|
| 1 | ticks all, presses **Start Extraction** | extraction thread running |
| 2 | presses **Start Extraction** again | no message box; program log only: "Extraction already in progress" |
| 3 | presses **SCAN FILES** | the scan ran to the end in its own thread **while the extraction thread was still running**; client found `ClientD-GJ`; the window switched to Step 2 |
| 4 | presses **START** on Step 2 | processing ran to the end in a third thread, still while the extraction was running; box "Processing Complete — Successfully processed 1/1 clients"; a new run folder `Annual Statement-231026 1326` was made in the same target |
| 4b | (automatic) | the processing end rebuilt the Step 4 tree: it now showed **"Source: Annual Statement-231026 1326"**, with 0 ticked rows — while the extraction was still working on `Annual Statement-050126 0900` |
| 5 | (the extraction is released) | it finished and wrote `Extraction_Log_20261023_1326.txt` and `PQ_Extraction_Report_20261023_1326.xlsx` into the **old** run folder; the tree still showed the new one |
| — | threads used (RAN) | window: `MainThread`; extraction: `Dummy-1`; scan: `Dummy-2`; processing: `Dummy-3` (Python's names for threads started by Qt) |

**Run 2 — closing the window during an extraction** (RAN `scripts/t34c_close.py` → `out2/t34c.txt`; the real event loop and the same exit path as `main.py`; each fixed wait shortened to 0.05 s but really waited). The user ticks one version (ITC and Sales) and presses Start Extraction; 0.3 s later closes the window.

| What | Observed |
|---|---|
| At the moment of closing | extraction thread running |
| `closeEvent` | accepted at once, window gone; `settings.json` written; `run_history.json` **not** written |
| The program's exit | `app.exec_()` returned 0 and the process **ended with exit code 0, with no message** while the worker was still inside the ITC refresh |
| Left on disk | the ITC report, an **unrefreshed copy** `ITC_Report_ClientA_DL_050126_0900_Refreshed_231026_1326.xlsx` (same size as the report), the Sales report untouched; **no consolidated workbook, no extraction log, no run history** |

On Windows with real Excel, the visible Excel that the refresh had started is a separate program and would most likely stay open with the copy loaded (READ, MEDIUM — how to check: on a test PC close B during a Step 4 refresh and look for EXCEL.EXE in Task Manager). (The Excel notes, XL §5.2 step 8, expected Qt to abort the program in this case (MEDIUM); in this run it did not — it ended normally with exit code 0.)

**Run 3 — Stop and second start during an extraction**: see B-F24 Run 2 (RAN): the second Start press is refused with a program-log line only; Step 3's Start Processing and Step 2's START stay enabled.

#### Traced output versus what the feature claims

1. **A second job of the same kind is refused with no on-screen word** (RAN) — B-G71 in Section 7.
2. **Jobs of different kinds run at the same time**: a scan and a processing run completed during an extraction, and the processing end switched the Step 4 list to another run folder mid-extraction (RAN Run 1) — B-G74.
3. **Closing the window during a job is accepted at once**: the job is cut off, half-made files stay (an unrefreshed copy), and no report, log or history is written (RAN Run 2) — B-G40.
4. **A late clean-up can delete a newer running thread and abort the program** (RAN by MW e13; LOW likelihood) — B-G41.
5. **Stop acts only between items; scanning cannot be stopped at all** (READ; RAN in B-F24) — B-G39, B-G70.
6. **Shared data is changed across threads without locks** (the Skip setting; the window's scan data) — RAN by DC T5e for the scan data — B-G75, B-G76.
7. Good: **no worker touches a window element directly** (RAN: thread names above; RAN by VW/MW thread maps).

#### Older documents versus code

- README.md:279 "Processing seems frozen → large files take time, check log" — in B the window does not freeze during the three jobs (they run in threads); the Step 4 list-building, however, walks the target folder on the window's own thread (B-F18, B-G58).
- README.md says nothing about running jobs at the same time or about closing the window during a run.

---
*End of the second part of Section 3 (B-F18 to B-F34).*
