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
| — folder-read error handling | `gui/pyqt_main_window.py:437-440` (listing the target), `:496-497` (walking clients: warning in the program log only) |
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

The disk-driven tree was added in **b0e7d0e** (09-02-2026, B only (after the split)). It replaced the list built from run history that 7e09e5c (08-02-2026, B only (after the split)) had introduced, which itself replaced the list fed from Step 3's results in c258047 (07-02-2026, B only (after the split)) (READ, notes H2 for b0e7d0e). Before the split, the separate Tkinter extractor window (`power_query_extractor/gui/extractor_window.py`, created in edac6a3, shared - before the split on 30-10-2025; deleted from B in a62d3cc) also scanned the disk, but chose the run folder by newest **modification time** and showed one version per client (`47a1492:power_query_extractor/gui/extractor_window.py:517, 525, 529`, READ via the H2 notes).

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
3. **"Refreshed" does not mean the refresh worked:** a copy from a failed refresh shows green "Refreshed" (RAN here; also RAN in B-F24's second run, where a version whose refresh check failed turned "Refreshed"). A version holding only a refreshed copy also counts as having a report (READ `:461`, the pattern matches copies; RAN by the MW reader, e6b). A custom file suffix always shows "Pending" (see B-F21).
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
| — COM start (COM = Windows' mechanism for one program to drive another), Excel by `Dispatch` | `:171, 174` |
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
| 8 | `Dispatch("WScript.Shell")`; `AppActivate("Excel")` (Excel's window caption) | `:216, 222` |
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
| Excel answers "Call was rejected by callee" (Excel busy) to every open | 3 tries × 3 opening attempts = 9 opens for the refresh (4 per try on Windows when the file has a separate 8.3 short name, `:709-716`), then it **carried on with no workbook**: the 5 s wait, `AppActivate` and the three keystrokes were still sent — into whatever window has the keyboard focus — and the result was `{'success': False, 'error': "'NoneType' object has no attribute 'Save'"}`. Extraction then failed too (no data). |
| The open fails with "Open method of Workbooks class failed" | No keystrokes; result `{'success': False, 'error': 'Open method of Workbooks class failed'}`. The special branch at `:125-137` that should say "Cannot open file for refresh … Proceeding with extraction without refresh" **never ran**, because `_refresh_power_query_simple` catches every error itself (`:339-344`) and returns a status instead of raising. |
| `AppActivate` fails all 3 times | Only a program-log warning (`:229-230`); **the three keystrokes were still sent**; the stand-in then reported success. |

**What real Excel and Windows would do (READ, cannot run here):**

- `Dispatch` (rather than `DispatchEx`, which always starts a separate Excel) may hand back an Excel the user already has open (MEDIUM). Then the user's Excel is made visible and maximised, receives the keystrokes, and is closed by `Quit()` (`:357`) with alerts switched off, so unsaved work in it is not offered for saving. How to check: on a Windows test PC open Excel with an unsaved Book1, run Step 4 on one dummy client, and see whether Book1's Excel is reused or closed.
- `AppActivate("Excel")` may bring forward the user's own Excel window instead of B's (MEDIUM); the refresh would then run in the user's workbook, and B's copy would be saved unrefreshed.
- Alt+A, R, A are Excel's ribbon "key tips" for Data → Refresh All in an English Excel; another Excel version or display language may map them differently (MEDIUM).
- B never asks Excel whether the queries finished: no `RefreshAll()`, no `BackgroundQuery = False`, no `CalculateUntilAsyncQueriesDone()`, no check of a connection's `Refreshing` (READ, grep finds none of these in B). A slow query may still be running when B saves at `:305` (MEDIUM; depends on the templates, which are not in the repository).
- The copy is **always saved** at `:305`, also after an error dialog or a failed BB2 check (RAN stand-in in B-F26 and B-F20).
- Stop is not honoured inside these waits (see B-F24).

#### Comparison with Standard 12 "Excel through COM" (facts only)

| Standard 12 says (rebuild_pipeline_FINAL.md) | B's refresh does | Where |
|---|---|---|
| Separate hidden Excel with `DispatchEx`, `Visible = False`, `DisplayAlerts = False`; never take over the user's Excel (lines 852-855) | `Dispatch`, `Visible = True`, maximised, `DisplayAlerts = False` | `report_processor.py:174, 181-183` |
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
