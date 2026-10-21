### A-F16: Processing summary workbook and organisation report

#### What it does

During Step 3 (the "Processing" step, which copies each ticked client's files into folders and builds the two Excel reports) the program writes two kinds of record of the run:

- after each client, a small text file `_Organization_Report_<client>_<state>.txt` in that client's **Version folder** (the folder `Version-<DDMMYY HHMM>` that this run creates inside the client's folder). It lists the folders, the files of each type and three counts;
- after the last client, one Excel workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` in the **Level-1 folder** (the run folder `Annual Statement-<DDMMYY HHMM>`), with five sheets: Summary, Client Status, File Mapping, Errors and Variations.

Both are written by plain Python; the workbook is made with **openpyxl** (a Python library that writes `.xlsx` files without Excel), never through Excel.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Organisation text report: file name, contents, the three counts, writing | `core/file_organizer.py:317-379` (name `:330`; folder list `:346-350`; files per type `:355-360`; counts `:362-368`; written as UTF-8 (a text encoding that can hold every character) `:371-372`; any error → logged, `None` returned `:377-379`) |
| The run-wide lists behind two of the counts | `core/file_organizer.py:50-51` (`created_folders`, `copied_files`), filled at `:100, :108, :139` (folders), `:172` / `:195` (Level 1), `:231` (every copy result) |
| Called once per client; the return value is ignored | `gui/handlers/processing_handler.py:414` |
| Data collected for the workbook during the run | `gui/handlers/processing_handler.py:209-226` (the `summary_data` dictionary, a lookup table of named values); report counters `:315-317, :321, :325, :337, :346, :382-384, :395, :401, :409`; one client row per SUCCESSFUL client `:416-429`; one file row per copy result `:431-438`; one error row per failed client `:442-453` |
| When the workbook is written | `gui/handlers/processing_handler.py:455-475` — only `if processed > 0 and level1_folder` (`:456`) |
| Workbook builder | `core/excel_handler.py:371-409` (`create_summary_report`: new workbook `:381`, default sheet removed `:384-385`, five sheets `:391-395`, column widths `:398-399`, saved straight to its final place `:402`, success = "the file exists" `:405`); styles `:411-430`; sheets: Summary `:432-462`, Client Status `:464-491`, File Mapping `:493-519`, Errors `:521-552`, Variations `:554-585`; widths `:587-601` |

#### Input → Output

The run's results — counters, one record per successful client, the copy result of every file, one record per failed client, and the scan's list of unrecognised files ("variations", see A-F9 in Section 3, part 1) → one `_Organization_Report_<client as written in the file names>_<state as written>.txt` per processed client, and `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` in the Level-1 folder (only if at least one client was processed).

#### User interaction

None: automatic during and at the end of Step 3, after the user presses "🚀 START" (Step 2) or "🚀 START PROCESSING" (Step 3).

#### How the traces in this part were run (applies to A-F16 to A-F31)

- **Scratch copy.** All runs used `/root/merge-scratch/run/A-S3b/A`, a copy of the snapshot without `__pycache__` folders (Python's compiled-file caches); `diff -rq` against `/root/merge-scratch/A-copy` was empty and the sha256 checksums (fingerprints of the file contents) of both were identical before the runs (`out/acopy_checksums_before_v2.txt`, `out/copy_checksums_before_v2.txt`). Runs used `/root/merge-scratch/venv-A/bin/python -B` (Python 3.14.6, Tk 9.0 — Tk is the window toolkit built into Python, used here through Tkinter —, openpyxl 3.1.5) with `PYTHONDONTWRITEBYTECODE=1`, and `HOME` and `TMPDIR` pointed into the scratch folder, so the settings file and Excel's staging folder (the temporary local folder in which Step 3 builds a report before copying it to its place) were only written there.
- **Scripts and outputs.** Scripts are in `/root/merge-scratch/run/A-S3b/scripts/`, outputs in `/root/merge-scratch/run/A-S3b/out/` (same name, `.txt`). Some scripts were first written by an earlier, interrupted run of this same task; every one of them was read again and re-run for this text (all outputs quoted here are from the re-runs on 21-10-2026, 13:33-13:44). Scripts new in this run: `t19b_scan_realistic.py`, `t28_dark.py`, `t29_launcher.py`, `t29b_getattr.py`, `t30_logfile.py`, `t30b_basicconfig.py` (run 21-10-2026 after the others), `t31_help.py`. A verification pass re-ran every script once more (copies in `/root/merge-scratch/run/VERIFY-A-S3b/`, 21-10-2026, 14:43-14:51) and got the same results apart from clock times, made-up file sizes and the virtual screen size (details at the end of this part).
- **Windows.** Window traces ran under **Xvfb** (a virtual screen, so Tk windows can be built without a monitor). Message boxes were replaced by recorders so no click is needed.
- **Excel through COM** (COM = Windows' way for one program to drive another; here Python drives Microsoft Excel through the **pywin32** package) cannot run on this Linux machine. Those parts are traced **by reading**. For Step 4 the real Python code was additionally driven through `scripts/fakecom.py`, a **fake COM layer** (a stand-in that records every call and reads workbooks with openpyxl): it shows the order of the Python calls and the branches taken, **never** what real Excel would do.
- **Invented clients.** `<Client A>` = the made-up "ClientA Traders", state Maharashtra; `<Client B>` = the made-up "ClientB Industries Private Limited", Gujarat; `<Client C>` = the made-up "ClientC Services LLP", Karnataka. They appear inside made-up file names because the program reads the client from the file name; they are not real clients. `<target>` stands for the scratch target folder.

#### Concrete trace (RAN, me: `scripts/t16_17_step3.py` → `out/t16_17_step3.txt`)

**Input:** a source folder with three clients — `<Client A>` and `<Client B>` complete (7 files each: GSTR-2B Reco, IMS Reco, two GSTR-3B months, Sales, Sales Reco, Annual Report), `<Client C>` missing its Sales file (6 files) — plus `notes for ClientC.xlsx`, which matches no pattern. Two templates with a `Links` sheet (the sheet into which Step 3 writes folder paths and file names that the template's queries read). The **real** Step 3 worker `ProcessingHandler.process_files_thread` (`gui/handlers/processing_handler.py:182-530`) is driven by a stub window object (a stand-in for the main window that records log lines, progress and message boxes, and whose `root.after` runs a call at once). On Linux the reports are built by the openpyxl path (A-F14, A-F15 in part 1). The run time stamp was fixed per run so that later runs get a new Level-1 folder.

**Run R1 (fresh, stamp `211026 1030`), step by step:**

1. Scan (`core/file_parser.py`, A-F7): `<Client A>` 'Complete' 7 files 100.0 %; `<Client B>` 'Complete' 7 files 100.0 %; `<Client C>` 'Missing 1 files' 6 files, missing `['Sales']`, 83.3 % [GST RULE] (the completeness rule A-R28/A-R29, described only); variations: `('notes for ClientC.xlsx', 'No matching pattern found', 'Check file naming convention', 4831)` (the last value is the made-up file's size in bytes; it depends on the generated file: 4831 in the writer's run, 4833 when the verifier re-ran the same script on 21-10-2026, 14:43).
2. `processing_handler.py:209-226`: `summary_data` starts with `total_clients = 3`, all counters 0, the processing mode, the target folder text, and — at `:224` — a **reference** to the very list `app.variations` (not a copy).
3. For each client (`:228-453`): folders created (`:253`; the first client's Level-1 folder is kept, `:255-256`), files copied (`:271-273`), ITC report → `reports_generated`, `itc_reports_created` +1 (`:315-317`), Sales report → `sales_reports_created` +1 (`:382-384`), then `create_organization_report(folders, client_info)` (`:414`):
   - `core/file_organizer.py:330`: name `_Organization_Report_ClientA Traders_Maharashtra.txt` — client and state **exactly as written in the file names** (with spaces; the state NAME, not the code) [GST RULE] (state text, A-R23, described only);
   - `:346-350`: every folder except Level 1 and Level 2, relative to Level 1: `VERSION: ClientA Traders-MH/Version-211026 1030`, `GSTR3B: …/GSTR3B`, `ITC: …/ITC`, `SALES: …/Sales`;
   - `:355-360`: one block per file type with the (already cleaned) names, e.g. `GSTR-3B Export (2 files):` then `GSTR3B-ClientA_Traders-Maharashtra-Apr.xlsx`, `…-May.xlsx`;
   - `:363-368`: "Total Files" = `client_info['file_count']` (this client); "Successfully Copied" = the number of 'Success' entries in `self.copied_files`; "Folders Created" = the length of `self.created_folders`. Both lists belong to the organiser object, which lives for the **whole run** (`:50-51`).
   - Result (RAN): `<Client A>` "Total Files: 7 / Successfully Copied: 7 / Folders Created: 6"; `<Client B>` "7 / **14** / **11**"; `<Client C>` "6 / **20** / **16**".
   - Then the client row is added with status `'Success'` (`processing_handler.py:416-429`) and one File Mapping row per copied file (`:431-438`).
4. After the loop, `processed = 3 > 0` and a Level-1 folder exists (`:456`) → progress 95 % → name `GST_Processing_Summary_211026_1030.xlsx` (`:462-464`) → `create_summary_report` (`core/excel_handler.py:371-409`) → saved **directly** into the Level-1 folder (`:402`, no temporary copy) → `True` → log "✅ Success: GST_Processing_Summary_211026_1030.xlsx" and "Sheets: Summary, Client Status, File Mapping, Errors, Variations".

The workbook read back (RAN):

| Sheet | Contents |
|---|---|
| Summary | A1 "GST FILE PROCESSING SUMMARY" (merged A1:B1); then Processing Date `'211026 1030'` (the run stamp as text) · Total Clients 3 · Successful 3 · Failed 0 · Total Files 20 · Reports Generated 6 · ITC Reports Created 3 · Sales Reports Created 3 · Report Creation Errors 0 · Processing Mode `fresh` · Target Folder `<target>` · Include Client Name in Folders `No` |
| Client Status | header `Client, State, Status, Files, Missing, Complete %`; `ClientA Traders, Maharashtra, Success, 7, 0, '100.0%'`; `ClientB Industries Private Limited, Gujarat, Success, 7, 0, '100.0%'`; `ClientC Services LLP, Karnataka, Success, 6, 1, '83.3%'` — the percentage is **text** (`:485`) |
| File Mapping | 20 rows `Original File, Client, Type, Destination, Status`, e.g. `AnnualReport-ClientA Traders-Maharashtra-2024-25.xlsx, ClientA Traders-MH, Annual Report, <target>/Annual Statement-211026 1030/ClientA Traders-MH/Version-211026 1030/AnnualReport-ClientA_Traders-Maharashtra-2024-25.xlsx, Success` ("Client" is the internal key `<client>-<state code>`) |
| Errors | `No errors encountered` (merged A2:E2) |
| Variations | `notes for ClientC.xlsx, No matching pattern found, Check file naming convention, 4831, Rename file` (size in bytes) |

**Run R2 (stamp `211026 1045`): one client fails.** The trace put a FILE where `<Client C>`'s Level-2 folder must be created. `create_client_structure` raised `[Errno 17] File exists: '<target>/Annual Statement-211026 1045/ClientC Services LLP-KA'` → caught at `processing_handler.py:442-453` (failed +1, error row, log "❌ Failed ClientC Services LLP-KA: …"). Workbook (RAN): Summary "Successful 2 · Failed 1 · Total Files 14 · Reports Generated 4"; Client Status lists **only** `<Client A>` and `<Client B>`; Errors sheet: `13:33:20 | ClientC Services LLP-KA | Processing Error | [Errno 17] File exists: '<target>/…/ClientC Services LLP-KA' | Review`. Message box: "Successfully processed 2/3 clients".

**Run R3 (stamp `211026 1050`): every client fails** (all three Level-2 names blocked). `processed = 0` → no workbook at all (`:456`); no "📂 Output location" line (`:489`); no folder opened (`:505`); the log still ends "🎉 ════════ PROCESSING COMPLETE ════════" and "✅ Processed: 0/3 clients", and the box says "Successfully processed 0/3 clients" (RAN). The Level-1 folder holds only the three blocking files.

**Run R4 (stamp `211026 1055`, workflow 'full'): Stop after the first client.** When "✅ Completed ClientA Traders" was logged, the trace set `app.stop_requested = True` — exactly what the STOP button does (`processing_handler.py:179`). The loop stopped before `<Client B>` (`:229-230`). Workbook (RAN): "Total Clients **3** · Successful **1** · Failed **0**" — `<Client B>` and `<Client C>` appear on no sheet; box "Successfully processed 1/3 clients"; the hand-over to Step 4 still happened (A-F5, part 1).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G19 in Section 7): the per-client text report shows run-wide totals.** "Successfully Copied" and "Folders Created" count the whole run so far, not the client (`core/file_organizer.py:363-368` reads the run-wide lists `:50-51`): `<Client C>`, with 6 files, reads "Successfully Copied: 20, Folders Created: 16" (RAN). Only the first client's numbers are right. HIGH.

> **BUG (for Section 7; not yet in Section 7): after Stop, the summary workbook does not account for the clients that were not processed.** "Total Clients 3, Successful 1, Failed 0" (RAN R4): the two unprocessed clients are in no sheet and no count, and nothing says the run was stopped (`processing_handler.py:229-230` breaks without recording them; the box wording is A-G69). HIGH (mechanism).

> **BUG (for Section 7; the box wording matches A-G69, now RAN; the missing workbook itself is not yet in Section 7 — neither A-G19 nor A-G69 mentions it): when no client succeeds there is no workbook, yet the run "completes successfully".** No `GST_Processing_Summary_…` file, no "Output location", no folder opened, but "PROCESSING COMPLETE" and "Successfully processed 0/3 clients" (RAN R3). The reasons for the failures exist only in the on-screen log and the log file. HIGH.

- **Failed clients appear only on the Errors sheet** (R2, RAN): Client Status lists successful clients only (`:416-429` runs only on success), so the Client Status sheet alone does not show that a client is missing.
- **"Success" means "no exception", not "both reports made"** (READ): the client row is always `'Success'` (`:423`) once the per-client `try` ends normally, even if both reports failed (they only add to "Report Creation Errors", `:321, :325, :337, :346, :395, :401, :409`). The flags `itc_report_created` / `sales_report_created` are collected (`:427-428`) and never written (`core/excel_handler.py:479-491` writes six columns only). In this trace every report succeeded, so the case was not run (READ, HIGH; A-G19).
- **"Complete %" is text** (`'83.3%'`, `core/excel_handler.py:485`), so Excel cannot add or sort it as a number; "Processing Date" is the run stamp text `211026 1030` (day-month-year and minute), not a date (RAN).
- **The workbook is written straight to the share** (`:402`, no temporary file) and a second run in the same minute overwrites it with only its own clients (RAN TRACE T6 run 2: 5 clients → 1 client; A-G9); an interrupted save leaves a broken workbook under the final name (A-G12).
- **The Variations sheet reads the live list** (`processing_handler.py:224` passes a reference): a Re-scan pressed during the run empties it, and the sheet then says "No file variations found" (RAN GUI exp8; A-G13).
- **The organisation report's "FILES ORGANIZED" list is every file of the client, not every file copied** (READ `core/file_organizer.py:355-360` loops over `client_info['files']`, not over the copy results): a file whose copy failed or was skipped is still listed (under its original name, because only successful copies are renamed, `:233-246`). LOW. (Not yet in Section 7.)
- The text report's file name uses the raw client and state text (`:330`); a write error is only logged and the run goes on (`:377-379`; return value ignored at `processing_handler.py:414`) (READ).
- Text beginning with "=" (e.g. a variation file name) is written into a cell as given and openpyxl stores it as a formula (A-G58 in Section 7; not re-run here).

Rules involved: A-R61 (summary workbook layout), A-R62 (organisation text report), A-R28 and A-R29 (completeness % and status) [GST RULE], A-R31 (scan statistics), A-R33 (run stamps), A-R60 (order and success counting) [GST RULE], A-R11 (variations), A-R23 (state text in the text report's name) [GST RULE], A-R122 (logging).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:237-245` (F5.1-F5.4): five sheets at Level 1, name `GST_Processing_Summary_DDMMYY_HHMM.xlsx`, "automatic at end of processing" — still true, except that "automatic" holds only when at least one client succeeded (`processing_handler.py:456`). The document does not mention the run-wide counts in the text report. Its location "create_summary_report method, lines 319-357" (`PROJECT_AUDIT_ANNUALSTATEMENT.md:239`) was right when it was written (`core/excel_handler.py:319@f7eb8f1`, A only (after the split), READ) but is stale now: the method is at `core/excel_handler.py:371-409` at the snapshot (the file grew in A's later commits).
- `README.md:182, 195` (RM26): summary file `GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx` — never true (the stamp is `DDMMYY_HHMM`, `processing_handler.py:186, 462-464`). `README.md:190, 204` (RM29): `_Organization_Report_ClientName_State.txt` in the Version folder, "Processing log for each client" — the name is still true (raw client name, full state name); as a "processing log" it is misleading, because two of its three counts are run-wide (see above). `README.md:196-200` (RM31): the five sheet names — still true.
- `README.md:43, 261, 310` (RM6) "Files are verified before finalization": the summary workbook is checked only for "the file exists" (`core/excel_handler.py:405`).

#### History

Both builders are unchanged since the first commit: `git blame` (a git command that shows, for every line, the commit that last changed it; RAN) attributes all 231 lines of `core/excel_handler.py:371-601` and all 63 lines of `core/file_organizer.py:317-379` to `edac6a3` (shared - before the split on 30-10-2025); the leftover `# NEW` markers at `core/excel_handler.py:449-451` are from that commit. The data collection in the worker (`gui/handlers/processing_handler.py:413-475`) is also from `edac6a3`; the only later change in `:413-530` is the workflow hand-over call (`510f7b4`, A only (after the split)).

### A-F17: Processing progress and log (Start / Stop, ETA, completion)

#### What it does

Step 3 is the screen the user watches while files are organised and reports are built. It shows a progress bar, a percentage label ("📊 Progress: X.X%"), a current-operation line ("Client i of n: <client> | ETA: …", ETA = estimated time still needed), a dark, colour-coded log of every step with a `[HH:MM:SS]` time, and two buttons, "🚀 START PROCESSING" and "⏹ STOP". The work runs on a **worker thread** (a second line of work beside the window, so the window can keep redrawing). At the end a box "Processing Complete" appears, the Level-1 folder is opened in the file manager, and (in the 'full' workflow) the program switches to Step 4.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Step 3 screen: bar, labels, buttons (wired with `.bind('<Button-1>')`), log box and its colour tags | `gui/tabs/processing_tab.py:17-117` (Start `:64-74`, Stop `:76-87` created `state='disabled'`, log `:103-110`, tags `:113-117`) |
| Start: guard against a second run, input checks, long-name question, button states, start of the worker (a **daemon thread**: one that is killed at once, without clean-up, when the program exits) | `gui/handlers/processing_handler.py:69-134` (guard `:75-76`; worker `:129-134`) |
| Stop: sets a flag and logs a line | `gui/handlers/processing_handler.py:177-180` |
| The worker: progress, ETA, per-client log lines, summary, completion box, opening the folder, hand-over, clean-up | `gui/handlers/processing_handler.py:182-530` (flag checked `:229-230`; ETA `:235-236`; bar `:239-241, 251-252, 266-269, 287-289, 457-459, 478-482`; completion `:484-502`; open folder `:505-516`; hand-over `:519`; fatal error `:522-525`; `finally` `:527-530`) |
| Log and progress helpers (both schedule the screen change with `after(0)`, Tk's "run this on the window's thread as soon as possible") | `gui/main_window.py:283-292`; `gui/utils/ui_helpers.py:16-40`; the worker's own scheduling helpers `gui/handlers/processing_handler.py:23-36` |
| Time estimate | `utils/helpers.py:476-523` (`ProgressTracker`, `format_duration`) |
| Start-up lines in the log; asking before closing | `gui/main_window.py:391-398`; `:344-355` |

#### Input → Output

The worker's progress (client number, file number, report step, results) → the bar and two labels, log lines in five colours (success `#4CAF50`, warning `#FF9800`, error `#F44336`, info `#2196F3`, normal `#FFFFFF` on `#1E1E1E`), the "Processing Complete" box, a file-manager window on the Level-1 folder.

#### User interaction

"🚀 START" on Step 2 or "🚀 START PROCESSING" on Step 3 starts a run (both call the same `start_processing`); "⏹ STOP" asks the worker to stop; the window's close button asks "Processing is still running. Stop and exit?" while Step 3 runs.

#### Concrete trace — progress, ETA and completion (RAN, me: `scripts/t16_17_step3.py`, the same three clients as A-F16)

To make the time estimate visible, the trace replaced the clock used by `ProgressTracker` with a fake one that moves forward 90 s each time a client finishes. The log's own "Total processing time" line uses the real clock (`processing_handler.py:185, 492-496`), so it said "0 minutes 0 seconds" in this trace; on a real run both clocks agree.

| Moment | What the code does | Bar / text shown (RAN, R1) |
|---|---|---|
| Client 1 starts | `progress_tracker.update(0)` (`:235`, first client adds 0) → ETA text "Calculating..." (`:236`); bar = 0/3 × 100 (`:239`) | `0.0`, "Client 1 of 3: ClientA Traders \| ETA: Calculating..." |
| Each file copied | `file_progress` (`:266-269`): bar = 0 + (file/7) × (50/3) — the first half of the client's share | `2.4`, `4.8`, … `16.7`, "Client 1/3: Copying AnnualReport-ClientA Traders-Maharashtra-2024-25.xlsx \| ETA: Calculating..." |
| Reports | `:287-289`: bar = 0 + 50/3 | `16.7`, "Client 1/3: Creating Excel reports \| ETA: Calculating..." |
| Client 2 starts (90 s later) | `update(1)` → `current = 1`; `get_eta()` = (3 − 1) ÷ (1 ÷ 90 s) = 180 s (`utils/helpers.py:497-510`) | `33.3`, "Client 2 of 3: ClientB Industries Private Limited \| ETA: 3m 0s" |
| Client 3 starts (180 s) | `current = 2` → (3 − 2) ÷ (2 ÷ 180) = 90 s | `66.7`, "… \| ETA: 1m 30s" |
| Summary (270 s, all clients done) | `current` is still 2 — the last client is never counted — so ETA = (3 − 2) ÷ (2 ÷ 270) = 135 s (`:457`) | `95`, "Creating summary report \| ETA: 2m 15s" |
| End | elapsed from the tracker (`:478-482`) | `100`, "Processing complete \| Total time: 4m 30s" |

Then (RAN): log "🎉 ════════ PROCESSING COMPLETE ════════", "✅ Processed: 3/3 clients", "📁 Files organized: 20", "📊 Reports created: 6", "📂 Output location: <target>/Annual Statement-211026 1030", "⏱️ Total processing time: …"; box `('Processing Complete', 'Successfully processed 3/3 clients\n\nFiles organized in:\n<target>')` — it names the **target** folder, not the new Level-1 folder; then `xdg-open "<target>/Annual Statement-211026 1030"` (Linux's "open with the default program"; `:512`; on Windows `os.startfile`, `:508`) and "📂 Opened output folder" logged although the stub returned the "command not found" code 32512 (the `os.system` result is not checked); buttons: Start `normal`, Stop `disabled`, `is_processing False` (`:527-530`). In R4 (Stop after client 1) the summary step showed "ETA: Unknown" (`current` stayed 0) and the box said "Successfully processed 1/3 clients".

#### Concrete trace — the real window (RAN, me, Xvfb: `scripts/t17_gui.py`, `scripts/t17_close.py`)

The real `GSTOrganizerApp` was built, the four paths set, the source scanned (three clients), all ticked. Each report was slowed by 0.6 s so that a run lasts a few seconds.

1. **STOP while idle.** Before any run the Stop button is `disabled`; a click on it (`event_generate('<Button-1>')`) still ran `stop_processing`: `stop_requested = True` and the line "🛑 Stop requested by user" appeared in the log (RAN). Harmless because Start resets the flag (`processing_handler.py:123`), but misleading.
2. **Start** (`root.after(100, …)` → `start_processing`): immediately after, Start `disabled`, Stop `normal`, `is_processing True`, 2 threads (window + worker) (RAN).
3. **Second Start click during the run** (700 ms): the click still reaches `start_processing` (the button only looks disabled, because `.bind` ignores the state), but the guard at `:75-76` returns at once — still 2 threads (RAN).
4. **Stop click during client 1** (1000 ms): bar at 16.7 %, "Client 1/3: Creating Excel reports". The worker finished the whole of `<Client A>` (its Sales report was written after the Stop line), then left the loop at `:229-230`. Log (RAN, abridged): "[13:35:56] 🛑 Stop requested by user" (from step 1) … "🚀 ════════ PROCESSING STARTED ════════" … "🛑 Stop requested by user" … "✅ Success: Sales_Report_ClientA_Traders_MH_…xlsx" … "✅ Completed ClientA Traders" … "📊 Creating Level 1 summary report..." … "✅ Processed: 1/3 clients" … "📂 Opened output folder" … "ℹ️ Workflow is 'organise only'. …".
5. **The completion box was raised from the worker thread**: the recorder saw thread `Thread-1 (process_files_thread)` call `showinfo('Processing Complete', …)` (RAN) — Tk is designed to be used from the window's own thread only (A-G70).
6. **Second run in the same session:** the log was not cleared — "PROCESSING STARTED" appears twice and the start-up lines are still there (RAN). The "[DEBUG]" lines (A-F15, part 1: 8 per client in a normal run, 14 in the code) appear in the user's log; the first one, checked by the trace, carries the plain `normal` colour tag (RAN). In the code 9 of the 14 use `normal`; `:353` and `:393` use `warning`, and `:403`, `:404`, `:411` use `error` — those 5 appear only when something goes wrong (READ). [Corrected during verification — originally stated "The 14 '[DEBUG]' lines … carry the plain `normal` colour tag"; the trace checked one line and the code gives 5 of the 14 other tags.]
7. **Closing the window during a run** (`t17_close.py`): `on_closing` asked "Processing Active / Processing is still running. Stop and exit?" (recorded answer Yes) → `stop_requested = True` and `root.after(1000, root.destroy)` (`gui/main_window.py:346-353`); the main loop returned while the worker was **still alive** (RAN: `is_alive() True`). In the real program `main.py` then ends and the daemon worker is killed wherever it is (READ; A-G27). Closing during a **Step 4** run asks nothing at all, because `on_closing` looks only at Step 3's flag (`:346`) (READ).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G69): "Successfully processed p/t clients" also after a Stop or when clients failed** (RAN: 1/3 after Stop; 0/3 when all failed, A-F16 R3), and nothing on screen says "stopped". The box names the target folder, not the Level-1 folder it then opens. HIGH.

> **BUG (for Section 7; matches A-G62): the greyed-out STOP and START still react to clicks**, and neither works from the keyboard (`.bind('<Button-1>')`, `gui/tabs/processing_tab.py:74, 87`). A Stop clicked while idle writes "🛑 Stop requested by user" into the log (RAN). HIGH.

> **BUG (for Section 7; matches A-G27): closing the window mid-run kills the worker** about 1 s later, mid-client (RAN: worker alive when the window loop ended); for Step 4 there is no question at all (READ). HIGH.

- **Stop is coarse** (RAN): it is checked only before each client (`:229`), so the current client — including its slow Excel work — always finishes.
- **ETA** (RAN): during the run the estimate covers the clients not yet finished, including the current one — correct for an even workload; at the summary step it still shows one client's time ("ETA: 2m 15s" after all three were done), because the tracker never counts the last client (`:235` adds 0 for the first client and 1 for each later one; A-G72). With one client, or after a Stop during the first client, it shows "Unknown". The registry (the ID list of this audit, `/root/merge-scratch/audit/A_registry.md`) wording "the ETA is one client behind" means this end-of-run effect.
- **Two clocks** (READ): the bar's "Total time" comes from the tracker (`:478`), the log's "Total processing time" from `datetime` (`:492-496`); they measure the same thing twice.
- **Only successful copies are listed** in the log (`:282-284`); failed or skipped copies get no line (READ; A-G67).
- **"📂 Opened output folder" is logged even when opening failed** on macOS/Linux (RAN: exit code 32512 ignored, `:510-514`); on Windows `os.startfile` errors are caught and logged as "Could not open folder" (READ). The folder path is placed inside a shell command on macOS/Linux (A-G57).
- **The worker touches the window directly** (READ): 45 direct `log_message` calls, two `update_progress` calls, the two message boxes and the button changes in `finally` run on the worker thread (`:319-530`; A-G70). It worked here on Linux/Tk 9 (RAN); on Windows not verifiable (MEDIUM).
- **The log is never cleared and never trimmed** (RAN, step 6); it grows by one line per copied file (A-G52).
- The 'full' hand-over runs also after a Stop (RAN R4; A-F5 in part 1).

Rules involved: A-R135 (run control), A-R140 (log and completion texts), A-R141 (progress and ETA), A-R154 (workflow hand-over), A-R60 (success counting) [GST RULE].

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:289` (F8.3) "Start button, daemon thread, Stop button, progress & log via `root.after(0, …)`" — the thread and Stop are as described; Stop acts only between clients.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:291` (F8.4) "All UI updates from the processing thread use `self.app.root.after(0, lambda: …)`" (a **lambda** is a small unnamed function; `root.after(0, …)` hands it to the window's own thread) — **never true**: at `47a1492` (the last commit shared - before the split on 30-10-2025) only 24 calls in the whole file were wrapped — 14 of them in the worker (10 log lines, 4 progress updates), the other 10 in the dry run (9, lines 36-52) and in `start_processing` (1, line 104) (`gui/handlers/processing_handler.py@47a1492`, RAN `grep` on `git show`) — while the worker made 45 direct `log_message` calls; at the snapshot 45 direct `log_message` calls, the message boxes and the button changes still run on the worker (DOCS F8.4, X6).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:811-815` (UI5): progress bar, labels, START/STOP, dark log — still true (file unchanged since `edac6a3`, shared - before the split on 30-10-2025).
- `README.md:160-162` (RM22) "Click Start Processing; monitor progress bar and log; wait for completion message" — still true. `README.md:158-159` (RM21) "Step 3: Processing Tab — 1. Click 'Dry Run'" — never true (Dry Run is on Step 2).

#### History

`gui/tabs/processing_tab.py`, `gui/utils/ui_helpers.py` and `ProgressTracker` (`utils/helpers.py:476-523`) are unchanged since `edac6a3` (shared - before the split on 30-10-2025) (RAN `git log`, `git blame`). In the worker, `6237597` (shared - before the split on 30-10-2025) wrapped part of the log calls in `root.after(0, lambda: …)`; A only (after the split): `0890022` (23-07-2026) replaced those lambdas with `_log_async` / `_progress_async`, which freeze the text when it is scheduled (the late-binding fix: a lambda in a loop otherwise reads the loop variable when it finally runs), `33d0baf` (23-07-2026) added the re-entrancy guard (a check that ignores a second Start while a run is going, `:71-76`), `510f7b4` (26-07-2026) added the hand-over (`:519`). `git blame` of `:182-412`: 217 lines `edac6a3`, 14 lines `0890022` (RAN).

### A-F18: Step 4 tab hosting the extractor (folder follows Step 1; fallback notice)

#### What it does

The fourth tab of the main window, "🔄 Step 4: Extract", shows a short explanation strip and, below it, the **Power Query Extractor panel** (Power Query = Excel's data-import tool; see A-F21) — the same screen the separate extractor program shows (A-F29), with its own two sub-tabs "Setup" and "Processing". Inside the main window the panel is built without its own blue banner and without its own status bar, and it does **not** scan when the program starts. The panel's folder box is filled with Step 1's target folder and keeps following it until the user puts a different folder into the box. If the extractor code cannot be loaded at all, the tab shows an orange notice instead of the panel.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Guarded import of the panel (a failure of any kind is caught) | `gui/tabs/extract_tab.py:21-30` |
| The tab, the explanation strip, the panel with `show_header=False, show_status_bar=False, auto_load=False` | `gui/tabs/extract_tab.py:48-82` |
| The orange "unavailable" notice | `gui/tabs/extract_tab.py:84-105` |
| Following Step 1's target folder (a **trace**: an automatic call made every time a Tk variable is written) | `gui/tabs/extract_tab.py:108-131` (`_pushed_folder` = the value this tab last put in, `:43, :82, :131`; "user chose their own" test `:127-128`) |
| Tab created when the main window is built | `gui/main_window.py:144-146` |
| What the three flags do inside the panel | `power_query_extractor/gui/extractor_window.py:107-139` (no scan when `auto_load=False`, `:132-134`), banner `:150-161`, status bar `:177-179, 573-585` |
| The 'full' workflow writes the panel folder directly | `gui/handlers/processing_handler.py:161-175` (`:166-168`) |

#### Input → Output

Step 1's target folder (and every later change of it) → the Step 4 folder box. If the extractor cannot be imported → the notice text plus the error message.

#### User interaction

The user opens Step 4. The folder follows Step 1 automatically; the user presses "🔍 Scan" (A-F19) when ready. In the 'full' workflow Step 3 opens this tab itself (A-F5).

#### Concrete trace — the folder box following Step 1 (RAN, me, Xvfb: `scripts/t18_tab.py`)

Input: a settings file (A-F27) whose `target_folder` is folder `t1`; the real main window is built; then the user changes the Step 1 target and types into Step 4's box.

| Step | What the code does | Result (RAN) |
|---|---|---|
| 1 Start-up | `ExtractTab` is built while the paths are still empty (`gui/main_window.py:146` runs before `apply_cached_values`, `:55`); the cached target is then written into `app.target_folder` → trace → `_on_target_folder_changed` (`extract_tab.py:121-131`): box empty → copy `t1` in, remember it | box = `t1`, `_pushed_folder = t1`, no client rows, no `level1_folder` yet (no scan) |
| | status bar | `show_status_bar = False`; status text "Ready" |
| 2 Step 1 target → `t2` | box still equals `_pushed_folder` → follow | box = `t2` |
| 3 User types `mine` into Step 4's box, then Step 1 target → `t4` | `current ('mine') != _pushed_folder ('t2')` → leave it (`:127-128`) | box = `mine` (kept, as intended) |
| 4 The 'full' hand-over (`processing_handler.py:161-175`) | selects tab 3, writes the Step 1 target into the box **directly** (`:166-168`) — `_pushed_folder` is not updated — and scans | box = `t4`, `_pushed_folder` still `t2`, tab 3 selected |
| 5 Step 1 target → `t5` | `current ('t4') != _pushed_folder ('t2')` → treated as the user's own choice | box stays `t4`: **Step 1 changes are no longer followed** |
| | the hand-over's scan of the empty folder `t4` | Step 4's own log: "Scanning folder for clients...", "No Annual Statement folders found"; status text "No processed folders found"; Step 3's log: "✅ Step 4 ready - review the client list and press Process to refresh." |
| | which widget shows the panel's status text? (a walk over every widget) | **none** |

Tab texts (RAN): `📁 Step 1: Setup`, `✅ Step 2: Validation`, `🚀 Step 3: Processing`, `🔄 Step 4: Extract`; panel sub-tabs `Setup`, `Processing` — a set of tabs inside a tab.

#### Concrete trace — the "unavailable" notice (RAN, me, Xvfb: `scripts/t18_notice.py`)

On this machine pywin32 is missing, yet the panel was built (`EXTRACTOR_IMPORT_ERROR = None` in the first trace), because the extractor guards its own pywin32 imports (`power_query_extractor/core/report_processor.py:20-36`). To see the notice at all, the trace made the import fail for a different reason: a simulated typo in `power_query_extractor/config/cell_mappings.py` (`SyntaxError: '(' was never closed`). The tab then showed (RAN): "⚠️  Power Query refresh is unavailable" / "This step needs pywin32 (Excel automation), which is not installed for the Python running this app. … Install it with:    pip install pywin32 … Details: '(' was never closed (cell_mappings.py, line 120)" / "Steps 1-3 are unaffected: organising files and creating the reports works without pywin32." The 'full' hand-over then failed with "Could not open Step 4 automatically: 'NoneType' object has no attribute 'folder_path'" (the guard at `processing_handler.py:153-155` tests for the tab, not for the panel) (RAN).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G65 — now RAN, was READ MEDIUM there): after the 'full' hand-over, Step 4 stops following Step 1.** The hand-over overwrites a folder the user chose in Step 4 and leaves `_pushed_folder` stale, so every later change of Step 1's target is ignored (RAN steps 4-5). The comment's promise "never clobber a folder the user chose in this tab" (`extract_tab.py:125-126`) is broken by the hand-over. HIGH.

> **BUG (for Section 7; matches A-G65): the panel's status messages are invisible inside the main window** — no widget displays the status text (RAN); "No processed folders found", "Error scanning folder" etc. appear only in the log on the other sub-tab ("Processing"). HIGH.

> **BUG (for Section 7; matches A-G66 and A-R129): the notice can no longer appear for the reason it names, and blames pywin32 for any other failure** (RAN: shown for a configuration typo, with the pywin32 text). Its last line, "Steps 1-3 are unaffected … works without pywin32", disagrees with `be6418c` (A only (after the split)), which makes pywin32 required on Windows (A-F1). HIGH.

- **"✅ Step 4 ready" is logged even when the Step 4 scan found nothing** (RAN step 5) — the panel's scan reports into its own log and returns normally (`extractor_window.py:625-628`). The text says "press Process" although the button reads "🚀 Start Processing" (A-G66). (Not yet in Section 7 as such.)
- The first Step 4 folder value at start-up comes from the settings file indirectly (through Step 1's variable), never from the panel's own cache reader (`load_from_cache` runs only when `auto_load=True`, `extractor_window.py:138-139`) (READ + RAN step 1).
- Step 4's work is invisible to Step 3 and to the close question (`gui/main_window.py:346` checks Step 3's flag only), and Step 3 does not check Step 4 either (READ; A-G71).

Rules involved: A-R129 (Step 4 loading guard), A-R143 (Step 4 panel behaviour), A-R144 (Step 4 folder follows Step 1), A-R154 (workflow hand-over).

#### Older documents against the code

- `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:27-30, 39-41` (P3, P5, P9): a fourth tab hosting `ExtractorPanel` — built as planned (`bfcb24a`, A only (after the split)), plus the guarded import and notice (`c463256`, A only (after the split)).
- Design note `:85-87` (P18) "pre-filled from Setup's target folder, follows it until the user overrides it" — built, except that the 'full' hand-over breaks it (RAN above).
- Design note `:54-56` (P12) "`logging.basicConfig` moves into the standalone entry point" (`basicConfig` = Python's one-time set-up of where log lines go; see A-F30) — not moved; avoided by importing `extractor_window` instead (`extract_tab.py:8-10`) (A-F30).
- `bfcb24a`'s message (A only (after the split)): "a notebook inside a notebook. Flattening is the next decision" — still nested at the snapshot (RAN tab list).
- `power_query_extractor/core/report_processor.py:11-13` calls pywin32 "genuinely optional" — out of date since `be6418c` (A only (after the split)) (DOCS X12).

#### History

All of `gui/tabs/extract_tab.py` is A only (after the split): `bfcb24a` (26-07-2026, 87 lines: the tab, the folder sync, `auto_load`) and `c463256` (26-07-2026, 44 lines: the guarded import and the notice, after hosting the extractor had made the whole program fail to start without pywin32) (RAN `git blame`). The flags `show_header`, `show_status_bar` come from `a677670` (25-07-2026, the split into `ExtractorPanel` + window shell), `auto_load` from `bfcb24a`.

### A-F19: Find each client's reports for Step 4 (latest run folder, version folder)

#### What it does

When the user presses "🔍 Scan" (or picks a folder with "📁 Browse") on Step 4's Setup sub-tab, the panel looks **inside** the folder in its box for entries named `Annual Statement-*`, takes the one with the newest "date modified" (**mtime**: the time a file or folder was last changed), and remembers it as the run folder for this Step 4 run. Inside it, every sub-folder whose name does not start with `_` is a client; for each, the `Version-*` folder with the newest date modified is used, and the panel notes whether that folder holds an ITC report and a Sales report. The list is then shown (A-F20).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Scan | `power_query_extractor/gui/extractor_window.py:611-668` (empty box `:613-615`; search `:622-623`; none found `:625-628`; newest by mtime and stored `:631-632`; clients `:638-659`; newest Version by mtime `:641-648`; error `:665-668`) |
| "Has ITC / has Sales" in one folder listing | `power_query_extractor/gui/extractor_window.py:20-42` |
| Browse (then scan at once) | `:601-609` |
| Standalone start-up scan from the settings file | `:587-599` (A-F29) |
| Called also by the 'full' hand-over | `gui/handlers/processing_handler.py:169` |

#### Input → Output

The folder in the box → `self.level1_folder` (the run folder later used for the consolidated workbook, A-F24) and a list of clients `{name = client folder name, latest_version, has_itc, has_sales}` → the client rows (A-F20); log lines on the "Processing" sub-tab; the status text (visible only in the standalone window).

#### User interaction

"🔍 Scan" or "📁 Browse" on Step 4; automatically at standalone start-up (A-F29) and after a 'full' Step 3 (A-F5).

#### Concrete trace (RAN, me, Xvfb: `scripts/t19b_scan_realistic.py`)

**Input:** a target folder holding (a) an older run `Annual Statement-181026 0900` (one client, both reports, three days old) and (b) the **real output of Step 3 run R1** from A-F16 (`Annual Statement-211026 1030` with `<Client A>`, `<Client B>`, `<Client C>`, each with an ITC and a Sales report, plus the summary workbook). The real `ExtractorPanel` is built as Step 4 builds it.

1. **Scan** (`:611-668`): box not empty → status "Scanning folder..." and log "Scanning folder for clients..." (`:617-618`) → `glob("Annual Statement-*")` (a wildcard search for names starting with "Annual Statement-") finds both run folders (`:623`) → newest by mtime = `Annual Statement-211026 1030` (`:631`), stored as `level1_folder` (`:632`), log "Using folder: Annual Statement-211026 1030" → `iterdir()`: `GST_Processing_Summary_211026_1030.xlsx` is not a folder → skipped (`:639`); the three client folders → each has one `Version-*` → `_classify_version_reports` (`:20-42`): both reports present → `display_clients` → log "Found 3 clients", Start enabled.

   Result (RAN): `ClientA Traders-MH`, `ClientB Industries Pvt Ltd-GJ`, `ClientC Services LLP-KA`, each `Version-211026 1030`, ITC tick ✓, Sales tick ✓, "Never refreshed". Note that `<Client C>` — the client with **no Sales file** — is offered with its Sales report ticked: that report is the unfilled template copy described in A-F15 (part 1).
2. A refresh of `<Client A>`'s ITC report fails before Excel saves anything (see A-F20/A-F22).
3. **Someone opens the OLDER run's summary workbook in Excel.** Excel writes its owner file `~$GST_Processing_Summary_181026_0900.xlsx` directly into that run folder; adding an entry changes the folder's date modified (RAN: old run folder now 13:42, new run folder 13:33).
4. **Scan again:** newest by mtime is now `Annual Statement-181026 0900` → Step 4 silently switches to the **older batch**: one client, `Version-181026 0900`, log "Using folder: Annual Statement-181026 0900", "Found 1 clients" (RAN).

**Further cases (RAN, me, Xvfb: `scripts/t19_20_25_panel.py`):**

| Case | Result (RAN) |
|---|---|
| A client folder holds two Version folders; the newer one (`Version-211026 1045`) has only an ITC report, the older one (`…1030`) both | the newer is used: ITC ✓, **Sales tick disabled** (the Sales report in the older Version folder is not offered) |
| A client folder with no `Version-*` folder; a folder `_archive` | both left out silently (`:639-642`) |
| Box pointed at a run folder itself (`…/Annual Statement-211026 1030`) | "No Annual Statement folders found"; the **previous list, run folder and Start button stay** (`:625-628` returns early) |
| A newer FILE named `Annual Statement-notes.txt` in the target | "Error scanning folder: [Errno 20] Not a directory: '…/Annual Statement-notes.txt'" (RAN) — the file was already stored as `level1_folder` at `:632` before `iterdir()` failed at `:638` (READ: the trace printed only the log line, not `level1_folder`; HIGH from the code order) |
| A scan while a run is in progress | Start button `normal` again (`:756-757`) |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G45): "newest" means "most recently changed", so ordinary actions switch Step 4 to an older batch.** Opening an older run's summary workbook in Excel was enough (RAN with a real file created in the folder, step 3-4). Nothing on the Setup sub-tab says which batch was chosen; the log line is on the other sub-tab. HIGH (code); MEDIUM that Windows updates folder dates the same way for every such action (NTFS, the disk format of Windows, updates a folder's date when entries are added, renamed or removed directly in it).

> **BUG (for Section 7; matches A-G45): pointing the box at a run folder finds nothing and silently keeps the old list** (RAN). The design note (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md:87-88`) says "pointing at an `Annual Statement-...` folder from a previous run still works", and the `bfcb24a` message (A only (after the split)) says "an older 'Annual Statement-...' batch can still be refreshed without touching Step 1" (RAN `git show -s`). Pointing at the batch folder itself does not work; pointing at a parent folder in which the older batch is the newest-dated one does (READ `:623, :631`). HIGH.

- **Only the newest Version folder of a client is looked at** (RAN): a report that exists only in an older Version folder of the same client is not offered. (Not yet in Section 7 as such.) MEDIUM.
- **A FILE named like a run folder breaks the scan and leaves `level1_folder` pointing at the file** (RAN); a later Start would write the consolidated workbook "into" a file and fail (A-G29).
- **The scan runs on the window's own thread** (READ: button `command=` → `scan_folder`), so the window freezes while a slow network share is listed (A-G48).
- **Folder order is the file system's order** (READ `:638`): the client list is not sorted (on a Windows NTFS disk listings are normally alphabetical — MEDIUM; on Linux the PQ reader saw `CLIENTB` listed before `CLIENTA`, RAN PQ).
- Two clients that share one Level-2 folder (A-F13 in part 1) appear as one client here (RAN TRACE T9).

Rules involved: A-R42 (which run and which version), A-R59 (report names recognised; the names carry the state code) [GST RULE], A-R43 (same "newest by date modified" rule in Re-run/Resume).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:413-417` (S4.7, S4.8): "scans for `Annual Statement-*` → latest; for each client latest `Version-*`" — still true. "Latest" run folder has always meant "newest date modified" (`power_query_extractor/gui/extractor_window.py:373@edac6a3`, shared - before the split on 30-10-2025; unchanged at `:631`, RAN `git blame`); "latest" Version folder means "newest date modified" only since `c1583d6` (A only (after the split); before: by name, `extractor_window.py:385@edac6a3`, which sorts day-first names wrongly across months). [Corrected during verification — originally said "'latest' is by date modified since `c1583d6`", which holds for the Version folder only.]
- Design note `:87-89` (P19) "pointing at an `Annual Statement-…` folder from a previous run still works" — not as built (RAN).

#### History

The scan is from `edac6a3` (shared - before the split on 30-10-2025: 47 of its 58 lines, RAN `git blame`); A only (after the split): `275d9a4` (23-07-2026, one listing per version folder in `_classify_version_reports`), `c1583d6` (23-07-2026, newest Version by date modified instead of by name).

### A-F20: Step 4 client and report selection with last-refresh status

#### What it does

After a scan, Step 4's right-hand column "Select Clients & Reports" shows one row per client: the client folder name, an **ITC** tick and a **Sales** tick (each enabled and pre-ticked only if that report exists in the client's newest Version folder), and a "Last Refresh Status" text such as "ITC: 2026-10-21 13:33 | Sales: Never". Four buttons change the ticks: "✓ Select All", "✗ Deselect All", "ITC Only", "Sales Only". The status column is re-read when a run ends.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Column, buttons, header row (widths 25 / 8 / 8 / rest), scrolling list | `power_query_extractor/gui/extractor_window.py:343-432` |
| Rows | `:670-757` (`display_clients`; ticks `:693-694`; disabled when the report is absent `:724, :734`; name label width 30 `:714`; Start enabled `:756-757`) |
| "Last Refresh Status" | `:45-82` (`_latest_refresh_times`), `:784-790`, `:792-797` (uses the CURRENT suffix box text), `:799-808` (texts) |
| Re-read after a run | `:759-782` (`refresh_status_display`), called from `:977-986` |
| Buttons | `:810-836` |
| What Start takes from the ticks | `:845-861` |
| Name test used for "refreshed" | `power_query_extractor/core/refresh_naming.py:43-55` (A-F22) |

#### Input → Output

The scan's client list → rows with two ticks and a status text; the ticks → the list Start sends to the worker (`{… 'process_itc': …, 'process_sales': …}` per client with at least one tick).

#### User interaction

Click the ticks or the four buttons; press "🚀 Start Processing".

#### Concrete trace (RAN, me, Xvfb: `scripts/t19_20_25_panel.py` and `scripts/t19b_scan_realistic.py`)

1. **Rows** (t19 case a): `ClientA Traders-MH` (newest Version has ITC only): ITC ✓ `normal`, Sales ☐ `disabled`, "Never refreshed"; `ClientB Industries Pvt Ltd-GJ` (Sales only, plus a refreshed Sales copy dated "2026-10-20 14:17"): ITC ☐ `disabled`, Sales ✓, "ITC: Never | Sales: 2026-10-20 14:17" (the script dates its made-up files relative to the moment it runs, so the times differ between runs: the verifier's re-run on 21-10-2026 showed "2026-10-20 15:25" here and "2026-10-20 14:45" in step 4).
2. **Buttons** (RAN): "ITC Only" → A (✓, ☐), B (☐, ☐); "Sales Only" → A (☐, ☐), B (☐, ✓); "Deselect All" → all ☐; "Select All" → A (✓, ☐), B (☐, ✓) — absent reports are never ticked (`:810-836`).
3. **Start with nothing ticked** → box "No Selection" / "Please select at least one report to process" (`:859-861`) (RAN).
4. **Changing the suffix box** to `_R{timestamp}` and re-reading (no file changed): `<Client A>` now reads "ITC: 2026-10-20 13:37 | Sales: Never" — the date of its ORIGINAL report, because the marker `_r` occurs inside every `itc_report_…` name (RAN; A-F22).
5. **A refresh that failed before saving** (t19b step 2): the leftover copy `ITC_Report_ClientA_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`, made by `shutil.copy2` (which copies the content AND the original's date), has the original's date (RAN: both 13:33); after `refresh_status_display` the row reads **"ITC: 2026-10-21 13:33 | Sales: Never"** — a "last refresh" that never happened (RAN).
6. **Alignment:** header widths `[25, 8, 8, …]`, row name width 30 (RAN) — the ticks sit to the right of their headings.

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G2): a failed refresh is shown as the last refresh**, dated with the original report's time (RAN step 5). The status column trusts the file NAME for "refreshed" and the file DATE for "when". HIGH.

> **BUG (for Section 7; matches A-G16): with some custom suffixes every original report counts as refreshed** (RAN step 4: `_R{timestamp}`). HIGH (mechanism).

- **A client with no Sales-type files still gets a pre-ticked Sales row** (RAN t19b step 1, `<Client C>`), because Step 3 always makes both reports (A-F15 in part 1; A-G5).
- **Ticks are offered only from the newest Version folder** (RAN step 1, A-F19).
- **Any scan re-enables "🚀 Start Processing", also during a run** (RAN; `:756-757`), and Start has no "already running" check (`:845-883`) — a second worker can be started (A-G65).
- **The status column uses the current suffix box text** (`:792-797`), which starts at the default after every restart (A-F22), so files made with a custom suffix in an earlier session read "Never refreshed" (READ; A-G16).
- An unreadable folder reads "Never refreshed", not "unknown" (`:74-75` swallows a listing error and `:784-790` returns `(None, None)` on any other error) (READ).
- Rows built while dark mode is on stay light (A-F28).

Rules involved: A-R65 (which reports are processed), A-R145 ("Last Refresh Status" text), A-R68 (recognising a refreshed file), A-R143 (panel behaviour).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:831-845` (UI9): headers Client Name / ITC / Sales / Last Refresh Status, the four buttons, a scrolling list — still true.
- The design note's testing section (`:108-109`, P26) says `_classify_version_reports` and `_latest_refresh_times` are unit-tested — never true at the note's date; today `_latest_refresh_times` has three tests (`tests/test_refresh_naming.py:88-121`, class `TestStatusDetectionUsesTheConfiguredSuffix`, added `2aaa2b1` (25-07-2026 23:11, A only (after the split)), 32 minutes after the note's commit `b276542` (22:39, A only (after the split)); RAN `git grep` at `b276542`: no test named either function), `_classify_version_reports` none.

#### History

The two-column Setup sub-tab with per-report ticks and the "Last Refresh Status" column is from `bd176bf` (30-10-2025, shared - before the split on 30-10-2025; 93 of the 167 lines of `:670-836`, RAN `git blame`). A only (after the split): `275d9a4` (one listing per folder), `2acffed` (25-07-2026, `refresh_status_display` after a run: before it the column kept the state from before the run until the user re-scanned), `2aaa2b1` (25-07-2026, recognition by the configured suffix; before: hard-coded `_refreshed_`).

### A-F21: Refresh reports through Power Query

#### What it does

**Power Query** is Excel's data-import tool; the ITC and Sales templates carry queries that read the client's copied files through the folder and file names written into the `Links` sheet in Step 3. "Refreshing" re-runs those queries. For each ticked report, Step 4 makes a copy of the report under a "refreshed" name (A-F22), opens that copy in a **visible, maximised** Excel, brings Excel to the front, presses the keys **Alt+A, R, A** ⚠ by **SendKeys** (pressing keys by program into whatever window has the keyboard; in English Excel: Data → Refresh All; the key letters depend on Excel's version and display language), waits fixed times while looking for error dialogs (A-F26), checks one cell (`Info!BB2` ⚠ — a template cell position) [GST RULE], **saves the copy in every case**, closes it and quits that Excel. Then the values are read from the copy (A-F23).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| pywin32 guard (no fallback: without pywin32 every report fails) | `power_query_extractor/core/report_processor.py:20-36`, refusal `:66-77` |
| Per client: ITC then Sales, or "Not selected for processing" | `:62-106` |
| Per report: find the original, name the copy, Skip Refresh, copy, refresh, wait 3 s, extract | `:108-207` (find `:112-125`; name `:127-130`; skip `:132-159`; copy `:164`; refresh `:168`; dead `except` `:169-181`; wait `:183-185`; extract `:188`) |
| The refresh session | `:209-415` (`Dispatch` `:218` — "give me an Excel", which may attach to one already running; `DispatchEx` would always start a separate one; Visible/alerts/maximised with 5 tries `:221-234`; open with 3 tries `:236-249`; load wait `:251-255`; bring to front `:258-276`; keys `:278-293`; wait `:297-302`; dialog checks `:305-322`; extended wait `:325-330`; final check `:333-342`; BB2 check `:345`; save ALWAYS `:348-351`; close `:360-363`; result `:365-374`; inner error `:376-381`; unreachable `return` `:383-386`; outer error `:388-393`; clean-up `:394-415`) |
| The only success test | `:638-685` (`_validate_refresh`) |
| Opening with fallbacks | `:687-799` (`_safe_open_workbook`); "already open in this Excel?" `utils/excel_com.py:19-29` |
| Wait-time box (5-60 by the arrows, default 10) and what Start hands over | `power_query_extractor/gui/extractor_window.py:259-287`; `:863-865`; worker `:885-966` |

#### Input → Output

`ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` in the client's newest Version folder → `<stem><suffix>.xlsx` next to it, refreshed (if the keys reached Excel) and saved; a status `{'success', 'error', 'validation_message'}` per report.

#### User interaction

"🚀 Start Processing" on Step 4 (or in the standalone window, A-F29), with the wait time and suffix boxes and the Skip Refresh tick set beforehand.

#### Concrete trace — Excel through COM (READ; control flow RAN with the fake COM layer: `scripts/t21_23_26_com.py` → `out/t21_23_26_com.txt`)

**Input:** `<Client A>`'s `ITC_Report_ClientA_Traders_MH_211026_1030.xlsx` in `…/Annual Statement-211026 1030/ClientA Traders-MH/Version-211026 1030/`, ITC ticked, Sales not ticked; defaults: wait 10 s, suffix `_Refreshed_{timestamp}`, Skip Refresh off; Windows with Excel and pywin32; nobody touches the PC.

1. `extractor_window.py:845-883` (window thread): selected = `[{… 'process_itc': True, 'process_sales': False}]`; `processor.wait_time = 10`, `suffix_pattern = '_Refreshed_{timestamp}'` (`:864-865`); switch to the "Processing" sub-tab; Start disabled; log cleared; a daemon worker starts.
2. Worker `_process_clients_thread` (`:885-938`): header lines; "📋 Processing ClientA Traders-MH - Reports: ITC"; progress 0 %; `process_client` (`report_processor.py:62`).
3. `:66` pywin32 present → `:87` ITC ticked → `_process_report(…, 'ITC')`; `:97` Sales not ticked → status "Not selected for processing".
4. `:112-125`: `glob("ITC_Report_*.xlsx")` in the Version folder → one hit → `original_file`.
5. `:128-130`: copy name `ITC_Report_ClientA_Traders_MH_211026_1030_Refreshed_<DDMMYY_HHMM>.xlsx` (RAN: `…_Refreshed_211026_1338.xlsx`) ⚠ (the stamp is the clock time of the refresh, minute resolution).
6. `:134-135`: Skip Refresh off → `:164` `shutil.copy2(original, copy)` — on the share, network to network.
7. `_refresh_power_query_simple(copy)` (`:209-415`), recorded call order (RAN, fake COM):

```
pythoncom.CoInitialize()                         # :215  switch COM on for this worker thread
Dispatch('Excel.Application')                    # :218  may attach to an Excel already running (MEDIUM)
excel.Visible = True                             # :225  "Keep it visible to avoid issues"
excel.DisplayAlerts = False                      # :226
excel.WindowState = -4137                        # :227  maximised
Workbooks.Open('<…>_Refreshed_211026_1338.xlsx') # :241 -> :737 (first attempt, no options)
   (wait 5 x 1 s "Loading...")                   # :253-255
Dispatch('WScript.Shell'); AppActivate('Excel')  # :260-266  bring Excel's window to the front by its title
Shell.SendKeys('%a'); SendKeys('r'); SendKeys('a')   # :282-292  Alt+A, R, A with 1 s gaps
   (wait 10 x 1 s)                               # :300-302  wait_time
   (5 dialog checks, 1 s apart)                  # :310-319  A-F26
   (wait max(10, wait_time) = 10 x 1 s)          # :326-330
   (final dialog check)                          # :333-342
Range('Info!BB2').Value -> '<Client A>'          # :345 -> :656  the only success test
wb.Save()                                        # :349  ALWAYS
wb.Close(SaveChanges=True)                       # :361
wb.Close(SaveChanges=False)  -> raises, ignored  # :399  finally: wb was not set to None
excel.Quit()                                     # :406
pythoncom.CoUninitialize()                       # :412  while 'excel' and 'wb' still hold COM objects
```
   Result: `{'success': True, 'error': None, 'validation_message': 'Validated - Client Name: <Client A>'}`. Fixed waiting in this one report: **39 s** in 36 sleeps (5 + 1 + 2 + 10 + 5 + 10 + 2 + 1 = 36 s in the refresh, + 3 s release wait at `:185`) (RAN count).
8. `:183-185` "Waiting for file to be released..." 3 s → `:188` `_extract_data(copy, 'ITC')` (A-F23) → result `{'status', 'data', 'extracted_values', 'original_file', 'refreshed_file'}`.
9. Back in the worker (`extractor_window.py:918-938`): "✅ ITC Report - Refreshed successfully", "⏭️ Sales Report - Skipped"; after all clients the consolidated workbook (A-F24).

**Variants (control flow RAN with the fake COM layer; what Excel itself does: READ):**

| Variant | What happened (RAN) |
|---|---|
| **No pywin32** (the real module on this PC) | `process_client` returns at `:66-77` for both reports: "pywin32 is not installed for this Python, so Power Query cannot be refreshed. Install it with: pip install pywin32"; no copy is made; the run still writes a consolidated workbook full of ✗ / "Missing" (TRACE T4 Part C) |
| **`Info!BB2` empty after the refresh** | the same calls, the copy is **saved anyway** (`:349`), then status `{'success': False, 'error': 'Refresh validation failed: Cell BB2 is empty'}`; the values are still extracted |
| **Excel busy: every open "rejected by callee"** | 9 `Workbooks.Open` calls (3 outer tries × 3 attempts inside `_safe_open_workbook`, the last with `UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False`), then — with **no workbook open** — the keys were still sent, then `wb.Save()` failed: status `{'success': False, 'error': "'NoneType' object has no attribute 'Save'"}`; 60 s of fixed waits; the unrefreshed copy stays under the "Refreshed" name and is then read by the extraction (A-G73) |
| **Excel could not be brought to the front** (`AppActivate` failed 3 times) | logger only (the log file, not the screen): "Could not activate Excel window, continuing anyway..." (`:273-274`); `SendKeys('%a')`, `'r'`, `'a'` sent anyway — on a real PC to whatever window has the keyboard; status **`success: True`**, because BB2 already held a value in the copied report |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G3): "Refreshed successfully" does not prove a refresh happened.** The only test is "`Info!BB2` ⚠ is not empty, not 0, not an error text" (`:638-685`; the test itself is A-R73 [GST RULE]; which cell proves a refresh is the author's call); a report whose BB2 already holds the client name before refreshing passes even when the key presses went to another window (RAN variant "not brought to the front") or Excel never got them. There is no `RefreshAll`, no wait for queries to finish, no "still refreshing" check (READ, grep). HIGH (code).

> **BUG (for Section 7; matches A-G2): a failed refresh is saved and kept under the "Refreshed" name** (`:349` saves after an error dialog or a failed check; a copy made at `:164` stays when Excel fails), and is later shown as the last refresh (A-F20) and reused by Skip Refresh as a success (A-F22). HIGH.

> **BUG (for Section 7; matches A-G55 and A-G54): keys go to whatever window is in front, and the Excel used may be the user's own.** `Dispatch` (not `DispatchEx`, which would always start a separate Excel) can attach to an Excel the user has open (MEDIUM — general pywin32 behaviour, cannot be tested here); that Excel is then maximised, keyed into and quit with alerts off, which closes it without asking to save (READ `:218-227, 404-409`; effect MEDIUM-HIGH). HIGH (code).

- **Open retries can run out silently** (RAN variant "Excel busy"): after three "Call was rejected by callee" failures the loop ends with `wb = None` and the code goes on pressing keys (`:239-249`; A-G73).
- **Fixed waiting is long and cannot be stopped** (RAN: 39 s per report at 10 s; READ: 136 s + 3 s at 60 s), with two Excel sessions per report and no Stop button (A-G50).
- **The `except` at `:169-181` can never run** (READ: `_refresh_power_query_simple` catches every error itself, `:376-393`), and the `return` at `:383-386` is unreachable (READ).
- **Open fallbacks**: attempt 2 closes, without saving, every workbook in that Excel with the same file NAME (`:744-753`) — if Excel is the user's own, a different workbook with the same name could be closed (READ, MEDIUM-LOW). For paths over 218 characters a `\\?\` prefix is tried (`:720-732`); whether `Workbooks.Open` accepts it is NOT SURE (LOW; how to check: open such a path once on Windows).
- **The key sequence assumes the English ribbon** ⚠ (Alt+A = Data tab, R, A = Refresh All; the key tips differ between Excel versions and display languages — NOT SURE which; how to check: press Alt in Excel on the office PCs and look at the letters shown) (READ `:282-292`; A-G44).
- **The wait-time box accepts typed values outside 5-60** (0 is accepted and skips the main wait — the extended wait still lasts `max(10, 0)` = 10 s, `report_processor.py:326`; text such as "abc" makes Start silently do nothing, because `IntVar.get()` (reading the box's whole-number Tk variable) raises inside the button's callback (the function the button runs)) (RAN PQ gui_probe; READ `extractor_window.py:864`; A-G26).
- All of this needs a logged-in, unlocked desktop where nobody types for the whole run (READ; Standard 12's "never a scheduled job" applies a fortiori).

Facts against the author's Standard 12, "Excel through COM" (`rebuild_pipeline_FINAL.md:844-878`; facts only, no recommendation): `Dispatch` instead of `DispatchEx` (`:218`); `Visible = True` (`:225`) where the standard asks for a hidden Excel; `DisplayAlerts = False` set (`:226`), but if all five tries fail the code continues anyway (`:234`); `Quit()` in `finally` — yes (`:404-409`); `IgnoreReadOnlyRecommended=True` only in the last-resort open (`:777-783`); `workbook.ReadOnly` never checked; no `BackgroundQuery = False`, `RefreshAll()`, `CalculateUntilAsyncQueriesDone()` or `Refreshing` check — the refresh is key presses plus fixed waits (`:278-330`); `RefreshDate` not used — yes; `CoInitialize` on the worker — yes (`:215`); `CoUninitialize` inside the same function while Excel objects are still held (`:412`); Excel needs a logged-in user — met (interactive only). The design note keeps the key-press method on purpose: "Rewriting the Power Query refresh itself (the SendKeys approach and its fixed time.sleep waits stay as they are)" is listed as out of scope (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md:125-126`).

Rules involved: A-R64 (no fallback without pywin32), A-R65 (which reports), A-R66 (which file is the original), A-R67 (copy name), A-R71 (key presses ⚠), A-R72 (waits and retries), A-R73 (success test, `Info!BB2` ⚠ template cell position) [GST RULE], A-R102 (open fallbacks), A-R103 (two Excel sessions per report), A-R104 (218-character threshold), A-R105 (dialogs, A-F26).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:321-338` (F11.1-F11.7): copy with suffix, visible Excel, Alt+A, R, A, wait 5-60 s (default 10), dialog check, BB2 check, new hidden Excel for extraction — still true; "waits configurable time" understates the real time (≈ 36 s + 3 s per report at the default).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:340` (F11.8) unreachable `return` — still present (moved to `:383-386`).
- `README.md:173` (RM25) "Click 'Refresh & Extract'" — never true; the button is "🚀 Start Processing" (`extractor_window.py:330`).
- `README.md:14, 24` "Refresh Power Query connections / Refreshes data connections in Excel (Windows only)" — the mechanism is key presses, which the README does not say.
- The commit message of `f363946` (A only (after the split)) says a live `Excel.Application` `Dispatch` was checked on the author's PC — the only sign in the history that this path was run against real Excel.

#### History

The refresh session, the dialog handling, the BB2 check and the open fallbacks are from `edac6a3` (shared - before the split on 30-10-2025): 199 of 207 lines of `:209-415`, all of `:417-525`, 159 of 167 lines of `:638-804` (RAN `git blame`). `bd176bf` (30-10-2025, shared - before the split on 30-10-2025) made the wait time and suffix configurable and added the per-report ticks. A only (after the split): `064c634` (removed a stray `print` above the method's docstring), `ba2a753` (shared "already open" helper), `2aaa2b1` (Skip Refresh by the configured suffix), `c463256` (26-07-2026, the pywin32 guard and the refusal at `:66-77`).

### A-F22: Skip refresh and refreshed-file naming

#### What it does

Two settings on Step 4's Setup sub-tab decide how refreshed copies are named and whether a refresh is done at all. **"Refreshed File Suffix"** (default `_Refreshed_{timestamp}`) is added to the original report's name, with `{timestamp}` replaced by the current `DDMMYY_HHMM`. **"Skip Refresh (Use existing refreshed files)"** reuses the newest existing refreshed copy of a report instead of refreshing it again. Small pure functions in `refresh_naming.py` decide which file names count as "refreshed"; the same test drives the "Last Refresh Status" column (A-F20).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Recognising a refreshed name: markers from the suffix, the legacy marker `_refreshed_`, "is a copy of this original" | `power_query_extractor/core/refresh_naming.py:14-66` (`refresh_markers` `:23-40`; `is_refreshed_name` `:43-55`; `is_refreshed_copy_of` `:58-66`) |
| Writing the copy name | `power_query_extractor/core/report_processor.py:51` (default), `:127-130` (only the lower-case `{timestamp}` is replaced) |
| Skip Refresh: find the newest existing copy, else refresh anyway | `:132-159` |
| The boxes and the tick; the tick writes a shared option | `power_query_extractor/gui/extractor_window.py:289-325` (suffix `:297`, tick `:315-325`), `:483-492` (writes `EXTRACTION_OPTIONS['skip_refresh']`), `:792-797` |
| The option's default | `power_query_extractor/config/cell_mappings.py:249` |

#### Input → Output

Suffix text + tick + the files already in the Version folder → the name of the new copy, or the choice of an existing copy (marked "Skipped - Using existing refreshed file", success).

#### User interaction

Type in "Refreshed File Suffix"; tick "Skip Refresh"; then "🚀 Start Processing". Neither setting is remembered between launches (A-R119).

#### Concrete trace (RAN, me: `scripts/t22_naming.py` → `out/t22_naming.txt`)

The real `_process_report` ran on fake Version folders; only the two Excel steps (`_refresh_power_query_simple`, `_extract_data`) were replaced by recorders, so the naming, copying, searching and Skip logic are real. Each original report was dated three days earlier.

| Case | Files before | What the code did (RAN) | Result |
|---|---|---|---|
| **1a** default suffix; the refresh **fails** (as when Excel cannot open the copy) | the original only | copy `…_1030_Refreshed_211026_1338.xlsx` made by `shutil.copy2`, refresh recorder returned failure, extraction ran on the copy | status `{'success': False}`; the copy stays, **dated like the original** (both 2026-10-18 13:38); the client list's "last refresh" = `('2026-10-18 13:38', None)` |
| **1b** next run with Skip Refresh ticked | original + the failed copy | the search `glob("ITC_Report_*.xlsx")` listed the copy FIRST on this disk, so the copy itself was taken as "the original"; `glob("<its stem>*.xlsx")` + `is_refreshed_copy_of` found the copy itself; the newest by date was reused; no refresh | status `{'success': True, 'validation_message': 'Skipped - Using existing refreshed file'}`, extraction from the unrefreshed copy |
| **2** suffix `_R{timestamp}` + Skip | the original only | markers `['_r']`; `'_r'` occurs inside `itc_report_…` → the ORIGINAL counts as its own refreshed copy | "Skipped - Using existing refreshed file", success, extraction from the **original** (never refreshed); the client list showed the original's date as "last refresh" |
| **3** suffix `_Refreshed_{TIMESTAMP}` (capitals) | the original only | `str.replace("{timestamp}", …)` does not match capitals | copy named literally `…_Refreshed_{TIMESTAMP}.xlsx`; a second run overwrote the same file |
| **4** empty suffix | the original only | copy name = the original's name → `shutil.SameFileError` (Python refuses to copy a file onto itself; caught at `:198-207`) | status `{'success': False, 'error': "… are the same file"}`, no extraction |
| **5** suffix `{timestamp}` only | the original only | copy `…_1030211026_1338.xlsx` | refreshed, but `refresh_markers` falls back to `_refreshed_`, so the copy is never recognised: list "(None, None)" = "Never refreshed" |
| **6** default suffix, an earlier copy `…_Refreshed_211026_1100.xlsx` present | original + copy | this time the original was listed first (directory order is not controlled, `:113, 125`) | new copy `…_Refreshed_211026_1338.xlsx`; with the other order, a copy of the copy would be refreshed (RAN TRACE T5: `…_Refreshed_211026_1040_Refreshed_211026_0422.xlsx`) |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G2): Skip Refresh reuses a copy whose refresh failed and reports success** (RAN 1a → 1b). Nothing records whether a copy was ever refreshed; the name and the date are all the code looks at. HIGH.

> **BUG (for Section 7; matches A-G16): the suffix setting can make originals look refreshed, or refreshed files invisible** (RAN 2, 3, 5), and an empty suffix fails every report (RAN 4). The suffix is not checked when typed. HIGH (mechanism).

> **BUG (for Section 7; matches A-G15): the "original" report can be an earlier refreshed copy** — the search pattern `ITC_Report_*.xlsx` also matches refreshed copies, and the first file the file system lists is taken (RAN 1b on Linux; on a Windows NTFS share the original usually sorts first because "." sorts before "_" — MEDIUM). MEDIUM.

- A client whose cleaned name contains the word "Refreshed" (so that the report name contains `_refreshed_`, e.g. the made-up `Refreshed Test Traders` → `ITC_Report_Refreshed_Test_Traders_…`) makes its original look refreshed under every suffix (RAN TRACE T5: legacy marker, `refresh_naming.py:55`); with a custom suffix `_v{timestamp}`, every client whose name starts with "V" (RAN TRACE T5).
- Two refreshes of one report in the same minute overwrite the earlier copy (minute stamp, `:128`; READ).
- The Skip tick changes a module-level dictionary (one copy for the whole running program) shared by every panel in the program (`extractor_window.py:486-487`), stays usable during a run and takes effect for the next report at once (RAN PQ gui_probe; READ). It always starts unticked, whatever `cell_mappings.py:249` says (`:315`); so if `:249` were set to `True`, the tick would show unticked while Skip Refresh is in fact on (the processor reads the dictionary, `report_processor.py:134-135`) until the tick is clicked on and off again (READ, HIGH).
- `EXTRACTION_OPTIONS['auto_refresh_if_missing']` ("Automatically refresh if no refreshed file found") is never read; the behaviour is hard-coded at `:156-159` (A-R70, dead).
- The suffix, wait time and tick are not saved anywhere (A-R119; RAN t27: the settings file holds no such key), so files made with a custom suffix in an earlier session read "Never refreshed" after a restart (READ; history note H-A2 finding 2).

Rules involved: A-R67 (copy name), A-R68 (recognising a refreshed file), A-R69 (Skip Refresh), A-R70 (dead option), A-R119 (not stored), A-R145 (status text).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:330` (F11.2): "copy with suffix (default `_Refreshed_{timestamp}`)" — still true.
- `README.md:247-253` (RM33): `EXTRACTION_OPTIONS` with `skip_refresh`, `auto_refresh_if_missing`, `show_timestamp`, `missing_data_text` — the text matches the code, but two of the four are never read, and the tick overwrites `skip_refresh` in memory.
- The module docstring of `refresh_naming.py:1-12` ("so all three agree") — the writer and the two readers agree on the default suffix; the report picker (`report_processor.py:112-125`) is a fourth reader that does not use the module and does not exclude copies.

#### History

Skip Refresh itself is older: at `edac6a3` (shared - before the split on 30-10-2025) it was already the option `EXTRACTION_OPTIONS['skip_refresh']` in `cell_mappings.py` with a hard-coded search `glob(f"{stem}_Refreshed_*.xlsx")` (`power_query_extractor/core/report_processor.py:76-93@edac6a3`, READ), and the copy name was a fixed `_Refreshed_{timestamp}` (`:71@edac6a3`). `bd176bf` (30-10-2025, shared - before the split on 30-10-2025) added the suffix box, the wait box and the on-screen Skip Refresh tick that writes that option (RAN `git show bd176bf`). [Corrected during verification — originally said `bd176bf` added Skip Refresh with the hard-coded search.] `2aaa2b1` (25-07-2026, A only (after the split)) added `refresh_naming.py` and made the client list and Skip Refresh follow the configured suffix ("Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one"), with 19 tests in `tests/test_refresh_naming.py` (18 test functions, one of them run for two inputs; RAN `pytest --collect-only`: 19) and 2 more in `tests/test_extractor_panel.py` — 21 collected in all; the message says "20 new", which is the number of test functions it added (RAN `git show 2aaa2b1`: 18 + 2). [Corrected during verification — originally stated "21 tests in `tests/test_refresh_naming.py`"; the history note H-A2 gives 21 = 19 + 2.]

### A-F23: Extract cells from refreshed reports

#### What it does

After each refresh (and also after a failed refresh, and in Skip mode) the refreshed copy is opened again in a **second, hidden** Excel, and the configured cells of the configured sheets are read. Each value is stored twice: under a field name built from the output sheet and column (e.g. `IMS_Reco_IGST`) and under the bare cell reference (e.g. `AA1`). The configuration is a Python file, `power_query_extractor/config/cell_mappings.py`, which names 22 cells in 8 report sheets.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Configuration: 9 output sheets, 22 mappings [GST RULE] ⚠ | `power_query_extractor/config/cell_mappings.py:33-242` |
| Lookup table built at import (report type → sheet → cell → field name) | `:259-278` (field name `:276`) |
| Extraction session | `power_query_extractor/core/report_processor.py:527-636` (hidden Excel `:533-539`; "read-only" log lines `:541-543`; open `:542`; sheet found by EXACT name `:564-568`; missing sheet → its fields `None` `:570-576`; cell read and stored twice `:581-590`; cell error `:591-594`; sheet error `:596-601`; summary log of every value `:603-611`; any other error → `{}` `:615-617`; clean-up `:618-636`) |

The configured cells (READ `cell_mappings.py:33-242`) [GST RULE] — which report cell feeds which output column is the author's rule; every sheet name, cell address and table number below can change with the templates or the return formats ⚠:

| Output sheet | Report | Report sheet ⚠ | Cells ⚠ → output columns |
|---|---|---|---|
| GSTR-3B Difference | ITC | `Diff GSTR-3B` | J15 → Diff IGST, K15 → Diff CGST, L15 → Diff SGST |
| GSTR-1 Difference | Sales | `Diff - GSTR-1` | N1 → Diff Taxable Value, O1 → Diff IGST, P1 → Diff CGST, Q1 → Diff SGST |
| Turnover GSTR-1 | Sales | `Diff - GSTR-1` | I1 → Taxable Value |
| Reco Diff | ITC | `Reco` | C37 → IGST Diff, D37 → CGST Diff, E37 → SGST Diff |
| Table 6J Match | ITC | `T6 of R9` | K29 → Table 6J Match |
| Table 8D Limits | ITC | `T8 of R9` | G6 → Table 8D Limits |
| IMS Reco | ITC | `IMS_Reco` | AA1 → IGST, AD1 → CGST, AG1 → SGST |
| Extra In Purchase | ITC | `ExtraInPurchase` | O1 → IGST, P1 → CGST, Q1 → SGST |
| Extra In GSTR2B | ITC | `ExtraInGSTR2B` | S1 → IGST, T1 → CGST, U1 → SGST |

(The names "Table 6J" / "Table 8D" ⚠ and the sheets "T6 of R9" / "T8 of R9" ⚠ refer, by their wording, to tables of GSTR-9; only the names are described here — table numbers can change with the return format.)

#### Input → Output

The refreshed copy + the configuration → a dictionary `{field name: value, cell: value, …}` per report, handed to the consolidated workbook (A-F24).

#### User interaction

None: automatic after each refresh (A-F21).

#### Concrete trace (READ; control flow RAN with the fake COM layer: `scripts/t21_23_26_com.py`, Part 1)

**Input:** the refreshed copy of `<Client A>`'s ITC report from A-F21, with invented marker numbers: `Diff GSTR-3B` J15/K15/L15 = 101/102/103; a sheet named **`reco`** (lower case) with C37/D37/E37 = 109/110/111; `T6 of R9`!K29 = `Match`; `T8 of R9`!G6 = 0; `IMS_Reco` AA1/AD1/AG1 = 201/202/203; `ExtraInPurchase` O1/P1/Q1 = 301-303; `ExtraInGSTR2B` S1/T1/U1 = 401-403.

1. `:533-539`: `CoInitialize`; a NEW `Dispatch("Excel.Application")`; `Visible = False`; `DisplayAlerts = False`; log "✓ Excel instance created (hidden mode)".
2. `:541-543`: log "Opening file in read-only mode: …" → `_safe_open_workbook` opens it **normally** (read-write) → log "✓ File opened in read-only mode" (RAN: both lines in the GUI log).
3. `:546`: `CELL_MAPPINGS['ITC']` = 7 sheets.
4. For each sheet (in configuration order) → loop over `wb.Sheets` comparing `sheet.Name == sheet_name` **exactly** (`:564-568`):
   - `Diff GSTR-3B` found → J15 → `GSTR-3B_Difference_Diff_IGST = 101` and `J15 = 101`; K15, L15 likewise (RAN);
   - `Reco` **not found** (the sheet is called `reco`) → logger warning "⚠️ Sheet 'Reco' not found" (file log only) and `Reco_Diff_IGST_Diff = None`, `Reco_Diff_CGST_Diff = None`, `Reco_Diff_SGST_Diff = None` (no bare cell keys) (`:570-576`) (RAN);
   - `T6 of R9`: `Table_6J_Match_Table_6J_Match = 'Match'`, `K29 = 'Match'`; `T8 of R9`: `… = 0`, `G6 = 0`; `IMS_Reco`, `ExtraInPurchase`, `ExtraInGSTR2B`: 201-203, 301-303, 401-403 (RAN).
5. `:603-611`: "✅ Extraction complete. Total values extracted: <n>" — n counts field names AND cell keys — then every key and value, through `_log` into the GUI log and `logger.info` into the log file (A-F30).
6. `finally` (`:618-636`): `wb.Close(SaveChanges=False)`, `excel.Quit()`, `CoUninitialize()`.

Other values seen through the fake layer (RAN TRACE T4 Part B2): a number stored as text (`'1,234.00'`) is returned as text; a never-calculated formula comes back empty through openpyxl (real Excel would calculate it — not verifiable here); a sheet named `diff - gstr-1` (lower case) is "not found", so all five Sales values become `None`.

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G17): a sheet whose name differs only in letter case or spaces yields no values at all** (RAN: `reco` vs `Reco`; TRACE `diff - gstr-1`), and the warning goes only to the log file — the screen shows nothing, and the consolidated row can still show ✓ (A-F24). By contrast the BB2 check finds `Info` in any letter case (`:646`). HIGH (mechanism).

- **"Read-only mode" is claimed but not used** (RAN: the log says it; `_safe_open_workbook` passes no read-only option; the last-resort attempt even passes `ReadOnly=False`, `:780`). The copy is closed without saving (`:622`), so no harm is done today (READ; A-G66).
- **Values are extracted even when the refresh failed** (`report_processor.py:187-188`), and the consolidated workbook then throws them away (A-F24; A-G17).
- **Excel error cells** (`#N/A`, `#REF!` …) probably arrive through COM as large negative whole numbers, which the consolidated workbook would show as figures (MEDIUM — general pywin32 behaviour, not testable here; how to check: put `=NA()` in a mapped cell and run Step 4 once) (A-G18).
- **A second, hidden Excel per report** (`Dispatch` again, not `DispatchEx`): if it attaches to the user's Excel, that Excel is hidden and then quit (READ; A-G54). Facts against Standard 12: COM is used here only to READ values, which the standard (`rebuild_pipeline_FINAL.md:827`) says openpyxl can do safely; `CoUninitialize` runs inside the same function (`:633`) (facts only).
- Extracted figures and paths are written to the log file at INFO level (the normal level of detail, which the program writes by default; `:56, :590, :603-611`; A-F30; A-G56).

Rules involved: A-R74 to A-R82 (the nine output sheets and their cells ⚠) [GST RULE], A-R83 (how the mappings are read) [GST RULE], A-R103 (two Excel sessions per report), A-R122 (what is logged).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:337-338` (F11.7): "a NEW hidden Excel instance for extraction; reads cells per CELL_MAPPINGS" — still true; the "read-only" wording of the log is misleading (DOCS F11.7).
- `README.md:229-245` (RM33): the example mapping (`'GSTR-3B Difference'`, ITC, `'Diff GSTR-3B'`, `J15`, `'Diff IGST'`) matches the first entry of `cell_mappings.py:35-58` (`:35-43`) exactly; the README shows one mapping, the code has three in that output sheet (J15, K15, L15) and eight more output sheets (READ). [Corrected during verification — originally said the example "matches `cell_mappings.py:35-58` exactly".]
- The comment "Backward compatibility" / "Build old format for compatibility" (`cell_mappings.py:259, 263`) is misleading: the table built there is the one the extraction actually uses (`report_processor.py:546`) (READ).

#### History

`_extract_data` is unchanged since `edac6a3` (shared - before the split on 30-10-2025; all 110 lines, RAN `git blame`). The configuration file had the first three output sheets in `edac6a3`; `47a1492` (30-10-2025, the last commit shared - before the split on 30-10-2025) added the other six (Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B) and deleted an older `cell_mappings copy.py`; `cell_mappings.py` has not changed in A since (126 lines `47a1492`, 152 lines `edac6a3`).

### A-F24: Consolidated extraction workbook

#### What it does

When every selected client has been through A-F21 and A-F23, Step 4 writes one new Excel workbook, `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, into the `Annual Statement-…` run folder it is working on. The workbook has a "Summary" sheet and one sheet per output group of the configuration (A-F23 table: "GSTR-3B Difference", "GSTR-1 Difference", "Turnover GSTR-1", "Reco Diff", "Table 6J Match" ⚠, "Table 8D Limits" ⚠, "IMS Reco", "Extra In Purchase", "Extra In GSTR2B"). Each data sheet has one row per processed client: client name, "Successfully Run" (✓ / ✗) and the configured figures. At the end a box asks "Open the consolidated report?".

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Whole workbook: one sheet per output group, header row, one row per client | `power_query_extractor/core/data_consolidator.py:20-186` (sheet loop `:45`; title "<sheet> Report" merged A1:C1 `:66-71`; a first, unformatted header loop that is overwritten straight after `:73-80`; real header `:85-98`; rows `:104-166`) |
| ✓ / ✗: the report's REFRESH status, nothing else | `:108-120` |
| Any status other than success → every column "Missing" | `:128-129` [GST RULE] |
| Which extracted value goes into a column: first key that CONTAINS the column title, else the exact field name | `:139-150` [GST RULE] |
| Empty or `None` → "Missing"; non-zero numbers get `#,##0.00` | `:152-162` |
| Calculated columns (none configured: every mapping has `'calculation': None`) | `:130-136`, `:188-218` (dormant; `cell_mappings.py:15, 42-…`) |
| Column widths | `:416-432` |
| Summary sheet | `:220-248` |
| File name (year first, minute resolution) and save straight to the share | `:174-180` ⚠ |
| Any error: logged and raised again (the run's figures are lost) | `:184-186` |
| Never called: an older one-sheet layout and an "Extraction Details" sheet | `:250-352` (`_write_section`), `:354-414` (`_create_details_sheet`) |
| Caller: which folder, the closing messages, "Open the consolidated report?" | `power_query_extractor/gui/extractor_window.py:940-955`, `:968-975` (`os.startfile`, Windows only) |

#### Input → Output

The list of per-client results built by A-F21/A-F23 + the run folder **as it is set at the moment the last client finishes** (`self.level1_folder`, `extractor_window.py:945`) → `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` in that folder.

#### User interaction

None until the end; then "Yes" on "Processing complete! Open the consolidated report?" opens the workbook with Windows' default program.

#### Concrete trace (RAN, me: `scripts/t24_consolidator.py` → `out/t24_consolidator.txt`)

The real `DataConsolidator.create_report` ran on result records built exactly as `_extract_data` builds them (field name AND bare cell key, in configuration order). Every configured cell got a different invented "marker" number, so each output cell shows where its value came from.

**Input:** three clients in `Annual Statement-211026 1030`:

| Client | ITC report | Sales report |
|---|---|---|
| `ClientA Traders-MH` (<Client A>) | refreshed, BB2 filled → success; J15/K15/L15 = 101/102/103, C37/D37/E37 = 109/110/0, K29 = `Match`, G6 = 98765.4321, AA1/AD1/AG1 = 201/202/203, O1/P1/Q1 = 301-303, S1/T1/U1 = 401-403 | success; N1/O1/P1 = 1001/1002/1003, Q1 empty, I1 = 5000 |
| `ClientB Industries Pvt Ltd-GJ` (<Client B>) | refresh check failed ("Cell BB2 is empty") but values were read (each marker + 1000) | success; N1 = the text `'1,234.00'`, O1 = 0, P1 = 7.5, Q1 = `True`, I1 = 9000 |
| `ClientC Services LLP-KA` (<Client C>) | success, but the sheet `Diff GSTR-3B` was not found (its three fields `None`) | not ticked ("Not selected for processing") |

**Walk-through (code path, each step RAN):**

1. `:24` imports `EXTRACTION_CONFIG` (9 groups) and `EXTRACTION_OPTIONS` (`missing_data_text` = "Missing").
2. `:45-71` for "GSTR-3B Difference": no `description` → A1 "GSTR-3B Difference Report", merged A1:C1; header row 3 "Client Name | Successfully Run | Diff IGST | Diff CGST | Diff SGST" (`:85-98`).
3. `:104-120` <Client A>: `result['itc']['status']['success']` → True → "✓" green. Column "Diff IGST": `:141-145` looks for a key containing the text "Diff IGST" — keys use underscores (`GSTR-3B_Difference_Diff_IGST`), so none matches; `:147-150` exact key `GSTR-3B_Difference_Diff_IGST` → 101 → `:160-162` format `#,##0.00`.
4. <Client B>: success False → `:128-129` "Missing" in all three columns, although 1101/1102/1103 were read.
5. <Client C>: success True → "✓", but the three fields are `None` → "Missing" ×3 (`:153-154`).
6. … same for every group. For **"IMS Reco" column "IGST"** (<Client A>): `:141-145` takes the FIRST key that contains the text "IGST": in configuration order that is `GSTR-3B_Difference_Diff_IGST` → **101** (from `Diff GSTR-3B!J15`), and `break` — the configured `IMS_Reco!AA1` value 201 is never used. Same for CGST/SGST, and for "Extra In Purchase" and "Extra In GSTR2B".
7. <Client C>, "IMS Reco" "IGST": the first matching key is again `GSTR-3B_Difference_Diff_IGST`, but its value is `None` → `:147-150` exact key `IMS_Reco_IGST` → 201 — the value of the cell the configuration names (`IMS_Reco!AA1`), but only because Client C's `Diff GSTR-3B` sheet was missing (which cell is meant is the author's call).
8. `:172` Summary sheet; `:175-180` name from the clock → `PQ_Extraction_Report_20261021_1338.xlsx`, saved.

**Output (RAN, abridged):**

```
Summary:  Total Clients: 3 | ITC Reports Processed: 2/3 | Sales Reports Processed: 2/3
GSTR-3B Difference  ClientA ✓ 101 102 103 | ClientB ✗ Missing Missing Missing | ClientC ✓ Missing Missing Missing
GSTR-1 Difference   ClientA ✓ 1001 1002 1003 Missing | ClientB ✓ '1,234.00'(text) 0(no format) 7.5 TRUE | ClientC ✗ Missing x4
Reco Diff           ClientA ✓ 109 110 0 | ClientB ✗ Missing x3 | ClientC ✓ 109 110 0
IMS Reco            ClientA ✓ 101 102 103   <- configured cells hold 201 202 203
Extra In Purchase   ClientA ✓ 101 102 103   <- configured cells hold 301 302 303
Extra In GSTR2B     ClientA ✓ 101 102 103   <- configured cells hold 401 402 403
                    ClientC ✓ 201.. / 301.. / 401..  (= the configured cells, by accident: its Diff GSTR-3B values were None)
configured cell vs value written, row ClientA: mismatches 9 of 22
```

Further cases (RAN, same script):

| Case | Result (RAN) |
|---|---|
| Second run in the same minute with one client | the same file name; the earlier 3-client workbook is replaced by a 1-client one ("Total Clients: 1") |
| The run folder is gone when the workbook is saved (share disconnected) | `FileNotFoundError` from `wb.save` (`:179`), logged and raised (`:184-186`); the caller shows "Processing failed: …" (`extractor_window.py:957-960`); no workbook, the figures of the whole run are lost — the refreshed copies stay on disk |
| A re-scan of ANOTHER target folder while the run is going (RAN `t19_20_25_panel.py`) | the workbook for the first batch's clients was written into the other folder's newest batch `other_target/Annual Statement-221026 0800/` ("Total Clients: 2"); nothing was written in the batch actually processed |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G1): 9 of the 22 configured columns are filled from another cell** [GST RULE] ⚠. The IGST / CGST / SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" show the `Diff GSTR-3B!J15/K15/L15` values instead of `IMS_Reco!AA1/AD1/AG1`, `ExtraInPurchase!O1/P1/Q1` and `ExtraInGSTR2B!S1/T1/U1` (RAN, 9 MISMATCH lines). This audit only states that the code does not do what its own configuration (`cell_mappings.py`) names; which cell is meant is the author's decision. HIGH.

> **BUG (for Section 7; matches A-G17): figures that were read are shown as "Missing", and a ✓ row can hold only "Missing"** [GST RULE]. ✓/✗ follows the REFRESH status only (`:108-120`), and a non-success status blanks every column (`:128-129`) — RAN <Client B> ITC (values read, all "Missing") and <Client C> ITC (✓ with three "Missing"). A dialog from another program during the wait also makes a good refresh "not success" (A-F26 Part 4) and so blanks the client's figures. Stated as a technical point only: the data was read but is not shown, and the ✓ mark does not say whether figures were found; whether figures from a report whose refresh check failed may be used is the author's call. HIGH.

> **BUG (for Section 7; matches A-G29 and A-G9): one error loses the whole run's figures; a second run in the same minute replaces the first workbook** (RAN). MEDIUM.

> **BUG (for Section 7; matches A-G65 / A-G71; RAN): a re-scan during a run sends the workbook to whatever folder was scanned last**, because the folder is read at the end of the run (`extractor_window.py:945`), not at Start. MEDIUM.

- **Values change type or look on the way in** [GST RULE] (RAN; A-G18): a number stored as text stays text (`'1,234.00'`, no number format, Excel will not add it up); zero is written without the number format (`value != 0`, `:161`), so it shows "0" while its neighbours show "0.00"-style figures; `True` gets the number format (a `bool` counts as a whole number in Python, RAN: `True [fmt #,##0.00]`) and Excel shows TRUE; `False` equals 0, so by the same test at `:161` it is written without the format (READ); an empty string becomes "Missing"; a real zero stays 0 — so "Missing" and 0 mean different things. Excel error values: see A-F23 (MEDIUM).
- **Summary counts are refresh successes over ALL clients** (RAN): "ITC Reports Processed: 2/3" counts <Client C>'s ITC as processed (✓ with "Missing" values) and does not say that <Client C>'s Sales report was simply not ticked (counted as not processed). A-R91; A-G19.
- **"Open the consolidated report?" uses `os.startfile`** (Windows only; elsewhere the failure is only logged, `extractor_window.py:972-975`) (READ; A-G31).
- Dormant code that a reader may take as live: `_calculate_value` (no mapping has a calculation), `_write_section` and `_create_details_sheet` (never called; READ grep), `EXTRACTION_OPTIONS['show_timestamp']` (never read). A-R92 to A-R95.
- The client name in column A is written as given; a name beginning with "=" would become a formula (A-G58; READ).

Rules involved: A-R86 (file name ⚠), A-R87 (layout), A-R88 (which value goes into a column) [GST RULE], A-R89 ("Missing") [GST RULE], A-R90 (number display) [GST RULE], A-R91 (Summary), A-R92 to A-R95 (dormant), A-R74 to A-R82 (the nine output groups) [GST RULE].

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:342-350` (F12): "Creates a multi-sheet Excel report … with sheets for each EXTRACTION_CONFIG entry plus a Summary sheet" — true; it does not mention the column mix-up, the ✓-only-on-refresh rule or the "Missing" rule.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:435` "DataConsolidator.create_report() → PQ_Extraction_Report_*.xlsx at Level 1" — true, but "Level 1" is the folder scanned LAST, not necessarily the one processed (RAN).
- `README.md:15, 25` "Extract specific cell values from multiple reports into consolidated sheets" — true for 13 of the 22 configured cells; 9 come from other cells (RAN).
- `README.md:29` "Comprehensive Summary: Excel summary report with 5 detailed sheets" refers to Step 3's workbook (A-F16), not this one; README never describes this workbook's layout.

#### History

`data_consolidator.py` is entirely from `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) (RAN `git blame`), when the configuration had only three output groups whose column titles all contain a space ("Diff IGST" etc.) — so the "contains the text" look-up never matched, and every value came from the exact-name fallback. The six groups added by `47a1492` (30-10-2025, shared - before the split on 30-10-2025) introduced single-word titles ("IGST", "CGST", "SGST"), which DO occur inside the earlier `GSTR-3B_Difference_Diff_IGST` keys (RAN `git show edac6a3:power_query_extractor/config/cell_mappings.py` in `/root/merge-scratch/combined`: the only titles then were "Diff IGST", "Diff CGST", "Diff SGST", "Diff Taxable Value" and "Taxable Value"; HIGH that the mismatch began with `47a1492`, because before it no title without a space existed). The caller lines `extractor_window.py:940-975` changed in `c1f425e` (25-07-2026, A only (after the split): messages and the "open?" box moved onto the window thread).

### A-F25: Extraction log, statistics and run history

#### What it does

While Step 4 runs, its "Processing" sub-tab shows a progress bar, a progress line and a time-stamped log of everything the worker does, in four colours (info black, success green, warning amber, error red). The only statistics are the progress percentage, the closing lines and the consolidated workbook's Summary sheet (A-F24). No history of runs is kept anywhere.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Progress frame, log box (`Text`, Consolas 9, `#f8f9fa`), scrollbar, wheel handler, colour tags | `power_query_extractor/gui/extractor_window.py:494-571` (wheel `:558-565`; tags `:568-571`) |
| What the worker writes, and in which colour | `:885-966` (header `:888-892`; per client `:907`; per report `:918-938`; end `:941-952`; error `:957-960`) |
| Log is emptied at every Start | `:875` |
| Writing a line from the worker: time stamp taken now, drawn later on the window thread | `:989-1029` (`_ui` `:989-1010`; `log_message` `:1021-1025`; `_append_log` `:1027-1029`) |
| Every processor line goes to the log FILE and to the screen in "info" colour | `power_query_extractor/core/report_processor.py:53-60` |
| Lines that go to the log FILE only (`logger.warning/error`) | e.g. `:316, 339` (dialog found), `:571` (sheet not found), `:616` (extraction error), `:377` ("❌ Error during refresh"), `:246` (Excel busy), `:273-274` (Excel not in front) |
| No run history | (nothing to point at: no history file, table or list exists — READ, grep for "history") |

#### Input → Output

Worker messages → coloured lines in the log box; progress value 0-100; the Summary sheet's three counts.

#### User interaction

Watching; scrolling the log with the mouse wheel or the scrollbar. The box also accepts typing (it is never set to read-only).

#### Concrete trace (RAN, me: `scripts/t19_20_25_panel.py`, last part → `out/t19_20_25_panel.txt`; processor lines RAN in `t21_23_26_com.py`)

**Input:** two clients listed: `ClientA Traders-MH` (ITC ticked) and `ClientB Industries Pvt Ltd-GJ` (Sales ticked). The processor is replaced by a stand-in that sends the same kind of lines the real one sends: for <Client A> success; for <Client B> "❌ Cannot access file: …" through the screen callback, a dialog warning through `logger.warning` only, then failure with values read. During the run the user scans another target folder.

**Log as shown (RAN; the colour tag in brackets):**

```
[info]    [13:37:47] ============================================================
[info]    [13:37:47] PROCESSING STARTED
[info]    [13:37:47] Wait Time: 10 seconds
[info]    [13:37:47] File Suffix: _Refreshed_{timestamp}
[info]    [13:37:47] ============================================================
[info]    [13:37:47]                                    <- the "\n" at the start of the next message
[info]    📋 Processing ClientA Traders-MH - Reports: ITC   <- so this line has no time stamp
[info]    [13:37:47]   Opening Excel file for ClientA Traders-MH ITC
[success] [13:37:47]   ✅ ITC Report - Refreshed successfully
[info]    [13:37:47]   ⏭️ Sales Report - Skipped
[info]    📋 Processing ClientB Industries Pvt Ltd-GJ - Reports: Sales
[info]    [13:37:47]   ❌ Cannot access file: [simulated] network path not found   <- an error, in black
[info]    [13:37:48] Scanning folder for clients...      <- the re-scan writes into the same log
[info]    [13:37:48] Using folder: Annual Statement-221026 0800
[success] [13:37:48] Found 1 clients
[info]    [13:37:48]   ⏭️ ITC Report - Skipped
[warning] [13:37:48]   ⚠️ Sales Report - simulated failure
[info]    [13:37:48]      (Data extracted despite error)
[info]    📊 Creating consolidated report...
[success] ✅ Processing complete!
[success] [13:37:48] 📄 Report saved: PQ_Extraction_Report_20261021_1337.xlsx
[info]    [13:37:48] 📂 Location: …/other_target/Annual Statement-221026 0800
file-log-only:  "  ⚠️ Error detected: [simulated dialog text] (file log only)"   <- never on screen
log box state: normal (editable) | Start button during the run after the re-scan: normal
```

With the real processor and fake COM layer (RAN `t21_23_26_com.py`): one ITC report produces 163-187 screen lines from the processor (169, 170, 187, 163, 163, 169 in Parts 1-6; 187 in Part 3, where every open was rejected), **all** tagged "info", also in the parts where the refresh failed, while "❌ Error during refresh: …", "⚠️ Sheet 'Reco' not found", "Excel busy, retrying…", "⚠️ Error detected: …" and "Could not activate Excel window, continuing anyway…" went to the log file only (RAN, verifier's re-run on 21-10-2026 gave the same counts). [Corrected during verification — originally "163-169 screen lines … including '❌ Error during refresh' paths"; Parts 2 and 3 give 170 and 187, and "❌ Error during refresh" is a log-file-only line.]

**Wheel (RAN):** log at view 0.5001 → one small touch-pad step (`delta -60`) → 0.5001 (no movement, `int(-60/120) = 0`); one wheel notch (`delta -120`) → 0.5032.

**Summary sheet of that run (RAN):** "Total Clients: 2 | ITC Reports Processed: 1/2 | Sales Reports Processed: 0/2" — <Client B>'s ITC (not ticked) and <Client A>'s Sales (not ticked) count as "not processed".

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G67 and A-G65): errors are not shown as errors, and some are not shown at all.** Every line from the processor is "info" black (`report_processor.py:56-60`), so "❌ Cannot access file" looks like progress; the red "error" tag is used only when the whole run fails (`extractor_window.py:958`) or a scan fails (`:666`); warnings written with `logger.warning/error` (sheet not found, Excel busy, dialog text, Excel not in front) reach only the log file (RAN). MEDIUM.

- **A re-scan during a run writes into the same log and re-enables Start** (RAN; A-G65): "Start" became usable again while the worker was still running; pressing it would start a second worker (no "already running" check, `:845-883`).
- **The log box can be typed into** (RAN: state "normal"); typed text is mixed with the worker's lines (A-G65).
- **The log is cleared at every Start** (`:875`) and never trimmed during a run; one report adds about 165 lines, so a 100-client run with both reports would add about 33,000 lines to the box (READ, arithmetic MEDIUM; A-G52).
- **Messages starting with a line break lose their time stamp** (RAN) — cosmetic.
- **The wheel handler drops small movements** (RAN, `int(-1*(event.delta/120))` at `:561`), unlike the Step 1 and Step 4 list handlers which accumulate them (A-F31; A-G68). It is bound with `<MouseWheel>` only (`:565`).
- **The amber warning colour `#ffc107` on `#f8f9fa`** is pale (READ `:570`; LOW; how to check: look at a "⚠️" line on the office screen).
- **No Stop button and no run history** (READ; A-R123): once started, a run can only be ended by closing the program (A-G27); after it ends only the workbook (A-F24), the refreshed copies and the log files (A-F30) remain.
- Figures are written on screen and to the log file at INFO level (A-F23, A-F30; A-G56).

Rules involved: A-R91 (Summary counts), A-R122 (what is logged and at which level), A-R123 (no run history), A-R143 (Start re-enabled by any scan).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:847-850`: "Log text widget (Consolas 9pt, light background); Color tags: info=black, success=green, warning=yellow, error=red" — true; it does not say that processor errors use the info colour.
- `README.md:161` "Monitor progress bar and detailed log" (Step 3 section) and `README.md:279` "Processing seems frozen … check log for activity": during Step 4's fixed waits the log does show "Waiting… n/10 seconds" lines every second (READ `report_processor.py:301`), so the hint works for Step 4.
- The registry wording "the log box … is never trimmed" (A-F25 row) is right for a run, but the box IS emptied at every Start (`:875`) — correction listed for the notes.

#### History

The processing sub-tab is from `edac6a3` (shared - before the split on 30-10-2025; 69 of 78 lines of `:494-571`) plus `4c42758` (29-10-2025, shared - before the split on 30-10-2025: the log wheel handler). The worker's per-report lines came with `bd176bf` (30-10-2025, shared - before the split on 30-10-2025). A only (after the split): `c1f425e` (25-07-2026) moved every log and progress update onto the window thread through `_ui` — the time stamp is now taken when the event happens; `2acffed` (25-07-2026) re-reads the "Last Refresh Status" after a run (`:977-986`).

### A-F26: Automatic closing of Excel error dialogs

#### What it does

While a refresh waits (A-F21), the program looks over **every window on the desktop** for something that looks like an error dialog: a visible, enabled top-level window whose title contains one of ten words and whose window class (the internal Windows "kind" of window) is one of three dialog kinds. If it finds one, it keeps the dialog's text as the report's error, then clicks a button whose label contains "ok", "close", "yes" or "continue" — or, if there is none, asks the window to close. One such dialog makes the report "not successful", even if the refresh check passed.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Five checks one second apart after the wait, one more after the extended wait | `power_query_extractor/core/report_processor.py:304-322`, `:332-342` |
| Which windows count: visible + enabled; title (lower case) contains `'microsoft excel'`, `'error'`, `'problem'`, `'cannot'`, `'failed'`, `'unable'`, `'could not'`, `'exception'`, `'invalid'`, `'missing'`; class `'#32770'`, `'NUIDialog'` or `'bosa_sdm_msword'`; FIRST match in desktop order | `:417-453` (`_find_error_window`; `EnumWindows` over all top-level windows `:450`) |
| Dialog text: title + child "static"/"text" controls longer than 5 characters, spaces squeezed, cut at 200 characters; otherwise "Power Query error occurred" | `:455-490` (`_get_window_text`) |
| Dismissal: first child button whose label CONTAINS ok / close / yes / continue gets a click (`BM_CLICK`); otherwise `WM_CLOSE`; then 0.5 s | `:492-525` (`_close_error_window`) |
| Effect on the result: `success = not error_detected` (`:370-374`); a found dialog sets `error_detected` with its text as the error; a failed BB2 ⚠ check also sets it and REPLACES the dialog text with "Refresh validation failed: …" (`:353-356`) | `:353-374` |

#### Input → Output

The windows open on the desktop during the wait → the first "error-looking" dialog's text as the report's error, and that dialog clicked or closed.

#### User interaction

None; it happens by itself. A user working on the same PC during the run may have their own dialogs clicked.

#### Concrete trace — Windows only (READ; control flow RAN with the fake COM layer, which also fakes the window list: `scripts/t21_23_26_com.py`, Parts 4 and 5 → `out/t21_23_26_com.txt`)

**Input, Part 4:** the refresh of `<Client A>`'s ITC copy as in A-F21, while **another program** shows a dialog titled "Confirm File Delete - could not move to Recycle Bin" (class `#32770`, text "Permanently delete this file?", buttons "&Yes" and "&No"). The Excel main window "Book1 - Excel" (class `XLMAIN`) is also open.

1. `:305-312` first error check → `_find_error_window` → `EnumWindows` (`:450`) goes through both windows.
2. "Book1 - Excel": title contains none of the words → skipped (and its class `XLMAIN` is not a dialog class).
3. "Confirm File Delete - could not move to Recycle Bin": visible, enabled; contains `'could not'`; class `#32770` → listed; first match returned.
4. `:314-315` `_get_window_text` → "Confirm File Delete - could not move to Recycle Bin Permanently delete this file?" (title + static text).
5. `:316` `logger.warning("⚠️ Error detected: …")` → **file log only**.
6. `:317` `_close_error_window` → `EnumChildWindows` → "&Yes" contains "yes" → `SendMessage(BM_CLICK)` (RAN call recorded: `('win32gui.SendMessage', 22, '&Yes', '0xf5')`) → stop → 0.5 s.
7. Remaining waits, BB2 check passes (`'<Client A>'`), `wb.Save()`, close, quit (RAN).
8. Result (RAN): `{'success': False, 'error': 'Confirm File Delete - could not move to Recycle Bin Permanently delete this file?', 'validation_message': 'Validated - Client Name: <Client A>'}`; the values are still extracted (RAN); the screen log would show "⚠️ ITC Report - Confirm File Delete…" and "(Data extracted despite error)" (READ `extractor_window.py:923-925`: this trace called the processor directly, without the panel's worker); in the consolidated workbook <Client A>'s ITC columns would all read "Missing" (A-F24 rule `:128-129`). Fixed waits for this report: 34.5 s.

**Input, Part 5:** an Excel prompt titled "Microsoft Excel" (class `NUIDialog`) with buttons "Save", "Don't Save", "Cancel" and no static text.

- Found (title contains `'microsoft excel'`); text = just the title "Microsoft Excel" (buttons are not static text); no label contains ok/close/yes/continue → `PostMessage(WM_CLOSE)` (RAN: `('win32gui.PostMessage', 3, 'Microsoft Excel', '0x10')`); report marked `{'success': False, 'error': 'Microsoft Excel'}`; the copy was still saved.

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G55): the dialog search and the clicks reach other programs.** Every top-level window of the desktop is examined (`:450`), with no check that it belongs to this Excel; a dialog of another program whose title contains a word such as "could not", "error" or "missing" is clicked — "Yes" included (RAN Part 4: a made-up "Permanently delete this file?" dialog was answered "Yes"). HIGH (code); how often such dialogs appear during a run: NOT SURE.

> **BUG (for Section 7; matches A-G17 and A-G19): any matching dialog marks the report as failed, which blanks its figures in the consolidated workbook,** even when the refresh check passed (RAN Part 4) — and the saved copy is kept (A-F21). MEDIUM.

- **Label matching is by "contains"** (`:506`): a button labelled e.g. "Look in…" contains "ok", "Yes to All" contains "yes" (READ; LOW).
- **Which kinds of Power Query error produce a dialog at all** is NOT SURE (LOW): some refresh errors appear only inside the workbook or in the Queries pane; how to check: break one query's source path in a copy of a report and refresh it once on the office PC, then call `win32gui.GetClassName` on any dialog that appears. One listed class, `bosa_sdm_msword`, is by its name a Word dialog class (LOW).
- With `DisplayAlerts = False` (`:226`) Excel itself suppresses most of its own prompts, so the dialogs most likely to be found are the ones Excel cannot suppress — and those of other programs (MEDIUM, general Excel behaviour).
- The dialog text is kept only in the status (screen: one "⚠️" line in amber; file log: the full line) and cut at 200 characters (`:484`).

Rules involved: A-R105 (error-dialog handling), A-R72 (when the checks happen), A-R73 (success = check passed and no dialog; `Info!BB2` ⚠ template cell position) [GST RULE].

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:334` (F11.5) and `:428` "Check for error dialogs (enumerate windows)" — true; neither says that ALL windows of the desktop are examined or that "Yes" is clicked.
- `README.md:14` and the design note say nothing about dialogs.

#### History

`_find_error_window`, `_get_window_text` and `_close_error_window` are unchanged since `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) (RAN `git blame`: all of `:417-525`). The checks' place in the wait (`:304-342`) is also `edac6a3`; `bd176bf` (30-10-2025, shared - before the split on 30-10-2025) made the extended wait follow the configured wait time.

### A-F27: Settings cache

#### What it does

The main program remembers nine settings between launches in one small text file in the user's home folder, `~/.gst_organizer_cache.json` (on Windows `C:\Users\<user>\.gst_organizer_cache.json`), written in **JSON** (a plain-text format of names and values). The settings are the four paths, processing mode, workflow mode, the global "include client name" tick, dark mode and the maximum client-name length, plus a list of up to five recent source folders that nothing reads. At start-up they are put back. The standalone extractor (A-F29) reads the same file to find the target folder.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| File location; load; save (read the old file for `recent_folders`, then rewrite the file in place) | `gui/utils/cache_manager.py:15`, `:18-27`, `:29-48`; unused helpers `:50-56` |
| What is saved | `gui/main_window.py:358-371` (`save_cache`) |
| What is restored, with defaults | `gui/main_window.py:120-121` (load at start), `:373-389` (`apply_cached_values`; dark mode applied at `:388-389`) |
| When a save happens | Browse of each path `gui/handlers/file_handler.py:25, 36, 47, 55`; workflow radio `gui/tabs/setup_tab.py:347-349`; length arrows `:436-438`; global include-name tick `gui/handlers/client_handler.py:223`; dark-mode switch `gui/main_window.py:196` |
| Readers in the extractor | standalone: `power_query_extractor/extractor_main.py:36-63` (four candidate files, four candidate keys); panel: `power_query_extractor/gui/extractor_window.py:118, 587-599` (only when no folder was handed over) |

#### Input → Output

The current settings (on certain clicks) → one line of JSON in the home folder; at start-up that file → the settings and the Step 4 folder.

#### User interaction

None directly: saves happen as a side effect of the clicks listed above.

#### Concrete trace (RAN, me: `scripts/t27_cache.py` → `out/t27_cache.txt`, real main window on a virtual display, HOME in the scratch folder)

| Step | What the user does | Settings file afterwards (RAN) |
|---|---|---|
| 0 | first start | no file |
| 1 | Browse → source folder `…/src1` | `{"source_folder": "…/src1", "itc_template": "", "sales_template": "", "target_folder": "", "processing_mode": "fresh", "workflow_mode": "organize", "include_client_name": false, "dark_mode": false, "client_name_max_length": 35, "recent_folders": ["…/src1"]}` |
| 2 | TYPES another source path `…/typed_src` | unchanged (`…/src1`) |
| 3 | clicks the "Re-Run" radio | in memory `rerun`, file still `fresh` |
| 4 | clicks the "Full pipeline" radio | now saves everything: `workflow_mode: full`, `processing_mode: rerun`, `source_folder: …/typed_src` |
| 5 | closes the window (no run) | unchanged — closing does not save (`gui/main_window.py:344-355`) |
| 6 | starts again | source `…/typed_src`, mode `rerun`, workflow `full` restored |
| 7 | types `abc` into the length box, then Browse | `TclError` (Tk's own error type): `expected floating-point number but got "abc"` from the save; Browse itself then raises the same error; file not updated |
| 8a | file holds broken JSON `{"source_folder": "x",` | starts with defaults; only a log-file line "Error loading cache: …" |
| 8b | file holds a JSON list `[1, 2]` | start-up fails: `AttributeError: 'list' object has no attribute 'get'` |
| 8c | file holds `{"client_name_max_length": "35"}` (text) | start-up fails: `TypeError: '>' not supported between instances of 'str' and 'int'` (`gui/main_window.py:387`) |
| 8d | file holds `{"processing_mode": "Fresh"}` (capital F) | starts with mode `Fresh` (RAN); no radio has that value, so none shows as selected, and a Step 3 run would stop with `ValueError("Invalid processing mode: …")` (`core/file_organizer.py:62-63`, case-sensitive; Section 6) (READ) |

The standalone extractor's readers (RAN `scripts/t29_launcher.py`, A-F29): home file present → its `target_folder` is used; no home file but a file `gst_organizer_cache.json` with key `output_folder` in the CURRENT folder → that one is used.

Side observation (RAN, steps 7 and 8, belongs to A-F4 in part 1): after a window is destroyed, a pending 350 ms checklist timer prints `invalid command name "…refresh_status"` on the console — harmless noise (LOW).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G24 and A-G26): some settings-file contents and box contents stop the program or every later save.** A JSON list or a length written as text makes start-up fail (RAN 8b, 8c); text in the length box makes every save — and every Browse — fail until it is corrected (RAN 7). MEDIUM.

- **Not every change is saved** (RAN 2, 3, 5; A-R116): typed paths, the processing-mode radios and closing the window do not save; the next save of any kind (step 4) stores whatever happens to be in memory at that moment.
- **A broken file loads silently as defaults** (RAN 8a) — the user is not told their settings were lost; the file is rewritten in place (`cache_manager.py:45-46`, no temporary file), so a save interrupted half-way would cause exactly that (A-G12; registry EXT t8).
- **An unknown mode value is accepted at start-up** (RAN 8d) and only fails later, at Step 3 (READ).
- **`recent_folders` is written but never read** (READ grep; A-R117).
- **Step 4's wait time, suffix and Skip tick are never stored** (READ `save_cache`; A-R119), so a custom suffix used yesterday makes yesterday's refreshed copies read "Never refreshed" today (A-F22).
- **Two program copies share the one file** and the last save wins (READ; A-G71).

Facts against Standard 13 (`rebuild_pipeline_FINAL.md:1039-1040`, "Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`), never in the shared app folder"; facts only): the settings are per user (home folder), not in the program folder — but in `%USERPROFILE%`, not `%LOCALAPPDATA%\<App Name>\`; the extractor still also looks for `gst_organizer_cache.json`, `.cache/app_state.json` and `config/last_session.json` in the CURRENT folder (`extractor_main.py:39-45`).

Rules involved: A-R114 (location and format), A-R115 (keys and defaults), A-R116 (when saved), A-R117 (`recent_folders`), A-R118 (how the extractor finds its folder), A-R119 (Step 4 options not stored).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:305-317` (F10) and `:450, 635, 713`: "Mismatch found … the PQ Extractor will NOT find the main app's cache … This is a confirmed bug" — **no longer true**: since `c1583d6` (23-07-2026, A only (after the split)) the home file is checked first (`extractor_main.py:39-41`; RAN t29 case A). The candidate list in the current folder is still there.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:314` lists eight saved keys; the code now saves nine (`workflow_mode` was added after that document) (READ `main_window.py:360-370`).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:106` "Dead method: `get_cached_value()` and `set_cached_value()`" — still true (`cache_manager.py:50-56`, READ grep).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:686-687` (D4) "For a single-user desktop app saving 8 settings … appropriate" — design opinion of that document; not assessed here.
- `README.md:171` "The extractor auto-loads target folder from cache" — true (RAN t29).

#### History

`cache_manager.py` is unchanged since `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) (RAN `git blame`). A only (after the split): `c1583d6` (23-07-2026) made the extractor read the home file first; `workflow_mode` was added to the saved keys when the workflow radios were added, in `510f7b4` (26-07-2026, A only (after the split); RAN `git log -S`; see A-F5 in part 1). [Corrected during verification — originally pointed to A-F3; the workflow radios are A-F5 in the registry.]

### A-F28: Dark mode / theme

#### What it does

A "🌙 Dark Mode" tick in the blue title strip turns the main window dark: every existing widget gets dark colours (surfaces `#2b2b2b`, input fields `#3b3b3b`, buttons `#404040`, text white) and eleven **ttk styles** (shared appearance settings of Tk's newer widgets) are changed. Each widget's original colours are recorded the first time it is darkened and put back when the tick is removed. The choice is saved (A-F27) and applied at start-up.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Colours, options, ttk styles | `gui/utils/dark_mode_manager.py:25-49` (`THEMED_OPTIONS`, `TTK_STYLES`), `:72-78` (`dark_colors`) |
| Switch on / off | `:105-116` (`apply_dark_mode`), `:117-…` (`restore_original_colors`) |
| ttk styles and the two fixed "selected" maps put back on restore | `:153-172` (`_restore_ttk`; fixed values `:164-172`) |
| Capture once: a widget already recorded is never recorded again | `:199-215` (`_capture`, check `:202-203`) |
| Step 2 row colours | `:271-294` (`update_tree_tags`); reset by every scan: `gui/handlers/client_handler.py:58-59` |
| The tick and what happens after it | `gui/widgets/title_bar.py:29-37`; `gui/main_window.py:157-207` (title strip restored `:165-173`; row tags `:177-178`; Setup checklist recomputed `:180-194`; save `:196`) |
| At start-up | `gui/main_window.py:388-389` |

#### Input → Output

Tick on/off → widget colours of the whole main window; `dark_mode` saved in the settings file.

#### User interaction

Click "🌙 Dark Mode" in the title strip.

#### Concrete trace (RAN, me: `scripts/t28_dark.py` → `out/t28_dark.txt`, real main window on a virtual display)

**Input:** start in light mode; tick Dark Mode; press Scan on Step 1 (three clients) and Scan on Step 4; untick; then (stand-in) let the program change one label's text colour to `#107C10` in light mode, and tick/untick again.

| Moment | What the window showed (RAN) |
|---|---|
| 0 start, light | root `#F8F9FA`; Step 3 progress label white/`#323130`; Step 3 log `#1E1E1E` (dark by design); selected Step 2 row: default `#4a6984` with white text; selected tab: `#dcdad5` |
| 1 ticked | root `#2b2b2b`; label `#2b2b2b`/white; Step 3 log `#3b3b3b`; row tags complete `#1a3d1a`, incomplete `#3d2f1a`, white text; Treeview (Tk's table widget, the Step 2 client list) `#3b3b3b`; selected tab `#404040`/white; Step 4 log `#3b3b3b`; settings file `dark_mode: true` |
| 2 after both scans, still dark | **Step 2 row tags `#E8F5E8` (pale green) and `#FFF4E5` (pale orange) with WHITE text**; **Step 4 client rows `white`** (built after the switch, never darkened) |
| 3 unticked | root, label, Step 3 log back to their start values; row tags pale with black text; but **selected Step 2 row now `#0078D4` with BLACK text** (was `#4a6984`/white) and **selected tab `#f0f0f0`** (was `#dcdad5`) |
| 4 stand-in label changed in light mode, then dark → light | label colour `#323130` — the first recorded value came back, the program's later `#107C10` was lost |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G64): widgets built while dark stay light, and a scan in dark mode makes the Step 2 list unreadable** — white text on pale green/orange rows (RAN 2), because each scan resets only the row backgrounds (`client_handler.py:58-59`) and leaves the dark-mode white text. LOW (cosmetic, but the Step 2 list is where clients are chosen).

- **The first switch changes two "selected" looks for good** (RAN 3): the restore applies fixed values (`dark_mode_manager.py:164-172`) instead of the ones recorded, so after one dark/light round the selected Step 2 row is blue `#0078D4` with black text and the selected tab is `#f0f0f0`.
- **Colours changed by the program in light mode after the first switch are lost on the next restore** (RAN 4 with a stand-in label). Correction to the registry's wording: in A today the only colours the program changes at run time are the Setup checklist's (`gui/tabs/setup_tab.py:553-566, 606-608`, READ grep), and those are recomputed after every switch (`main_window.py:180-194`), so no visible effect was found (MEDIUM — based on a search for colour changes).
- The Step 4 panel's own client rows, tooltips (always pale yellow, `gui/widgets/tooltip.py:54`) and message boxes are not themed (RAN/READ).
- The standalone extractor ignores the saved `dark_mode` (RAN t29 case A: settings file `dark_mode: true`, theme `clam`; READ grep: nothing in `power_query_extractor/` or `launch_extractor.py` reads `dark_mode`, so its window stays light — the trace itself did not print colours).

Rules involved: A-R148 (dark-mode colours and capture).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:293-303` (F9) "stores original colors on first run, applies dark palette (`#2b2b2b` bg, `#ffffff` fg, `#3b3b3b` widget bg) or restores originals" — the colours are true; "restores originals" is not true for the two "selected" maps (RAN 3) and widgets created later (RAN 2).
- `README.md:26` "Dark Mode: Toggle between light and dark themes" — true for the widgets that exist at the time of the switch.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:1028` (history of an earlier commit): "This redundant code was likely causing visual glitches" — the reset of the row backgrounds at `client_handler.py:58-59` is still there and is what produces the white-on-pale rows (RAN 2).

#### History

The manager was rebuilt in A only (after the split): `aeb83cf` (26-07-2026, "Rebuild dark mode so it restores the exact original colours": 255 of its 294 lines) and `725fbaf` (26-07-2026, tree background and tab padding: 10 lines of the manager, plus `main_window.py`); 29 lines remain from `edac6a3` (shared - before the split on 30-10-2025) (RAN `git blame`: 255 + 10 + 29 = 294). Two more A-only commits on the same day changed the dark mode around the manager, not the manager file itself (RAN `git show --stat`): `c89c1d4` ("Fix stale state colours after a dark-mode toggle": the Setup checklist recompute in `gui/main_window.py`, plus `tests/test_dark_mode.py`) and `169c7b5` (Setup checklist dark colours, `gui/tabs/setup_tab.py` only). [Corrected during verification — originally listed all four commits as rebuilding the manager and left out 725fbaf's 10 lines.]

### A-F29: Standalone extractor launcher

#### What it does

`launch_extractor.py` opens Step 4 on its own, without the main program: a separate 1000×700 window titled "Power Query Extractor" with its own header and status bar. It looks for the target folder in the main program's settings file (A-F27) and, if it finds one, scans it at once (A-F19).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Launcher: adds its own folder to the END of Python's search path, starts the extractor, shows an error box if that fails | `launch_extractor.py:1-30` (search path `:10`; error box and re-raise `:20-27`) |
| Log set-up at import (A-F30), finding the folder, opening the window | `power_query_extractor/extractor_main.py:12-19`, `:24-63`, `:65-72` |
| Window: title, size, minimum size, the global `clam` theme, the panel; forwarding of unknown names to the panel | `power_query_extractor/gui/extractor_window.py:85-95` (`configure_extractor_styles`), `:1032-1063` (`PowerQueryExtractorApp`; `__getattr__` `:1051-1063`) |
| Panel start-up: folder handed over → scan now; none → read the settings file itself | `:126-139`, `:587-599` |

#### Input → Output

`python launch_extractor.py` (from any current folder) + the settings file → the standalone window, already listing the clients of the newest run folder; `pq_extractor.log` in the current folder.

#### User interaction

Run `python launch_extractor.py` (README `:166-169`); then the same controls as Step 4 (A-F19 to A-F25), plus a status bar at the bottom.

#### Concrete trace (RAN, me: `scripts/t29_launcher.py` → `out/t29_launcher.txt`; each case is a fresh Python started in its own scratch "current folder"; a small wrapper puts the program folder first on Python's search path — as Python itself does when a script is started — and runs the real `launch_extractor.py` with `runpy.run_path(…, run_name='__main__')`; the only change: the window is read and closed 1.5 s after it opens)

| Case | Input | What happened (RAN) |
|---|---|---|
| A | home settings file with `target_folder` = a folder holding one run `Annual Statement-211026 1055` with `ClientA Traders-MH`, and `dark_mode: true` | log: "Found cache file: …/.gst_organizer_cache.json" → "Found target folder in cache: …" → window "Power Query Extractor" 1000x700, minimum (900, 600), theme `clam`; the saved `dark_mode: true` has no effect — the trace did not print colours, but nothing in `power_query_extractor/` or `launch_extractor.py` reads `dark_mode` (READ grep), so the window is light (HIGH); folder filled; status "Found 1 clients"; row ticked ITC + Sales; Start enabled; `pq_extractor.log` created in the CURRENT folder, none in the program folder |
| B | no settings file anywhere | "No cached folder found"; empty folder box; status "Ready"; Start disabled |
| C | no home file, but `gst_organizer_cache.json` with key `output_folder` in the current folder | that file is used ("Found cache file: gst_organizer_cache.json") and the folder scanned |
| D | `python power_query_extractor/extractor_main.py` | exit 1: `ImportError: attempted relative import with no known parent package` (the module starts with `from .gui…`, `:9`) |
| E (RAN `scripts/t29b_getattr.py`) | asking the standalone window for a Tk method that plain Tk windows forward to the Tcl interpreter (`eval`) | `AttributeError: 'ExtractorPanel' object has no attribute 'eval'` (a plain Tk window answers `9.0.3`); `app.log_message` is forwarded to the panel as intended |

Close handler (RAN, cases A-C): only Tk's default (`destroy`). Closing the window during a run therefore ends the program at once; the worker is a **daemon thread** (a background thread Python stops without waiting, `extractor_window.py:878-883`), so a refresh in progress is cut off and the visible Excel it was driving is left open (READ; effect MEDIUM — not testable here).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G27): closing the standalone window during a run kills the run without asking** (READ; Tk default close, daemon worker). The same holds for Step 4 inside the main window: the main window's close question looks only at Step 3's flag (`gui/main_window.py:344-355`), not at the panel's `is_processing` (READ). HIGH (code).

- **Alternative ways of starting fail** (RAN D; A-G32): running `extractor_main.py` directly fails; `python -m power_query_extractor.extractor_main` from the program folder works (RAN by the verifier: started from a scratch folder with the program folder on `PYTHONPATH`, the window opened with "No cached folder found" and stayed open until it was stopped after 8 s; no import error; `pq_extractor.log` appeared in the folder it was started from). [Corrected during verification — originally "would work (READ, MEDIUM — not run)".]
- **The program folder is added at the END of the search path** (`launch_extractor.py:10`); Python already puts the script's own folder first when the file is run as a script, so the line changes nothing in that case; when the launcher is imported from elsewhere, a different installed package called `utils` (imported by `report_processor.py:39`) would win (READ; LOW).
- **The window's `__getattr__` replaces Tk's own forwarding** (RAN E; nothing in A uses it — READ grep; LOW).
- **The standalone window keeps its own log file** `pq_extractor.log` in the current folder (RAN; A-F30) and ignores the dark-mode setting (RAN A with `dark_mode: true`; READ grep: never read).
- **Nothing stops a second copy** (standalone and the main program's Step 4 at the same time) from driving Excel together (READ; A-G71).
- Case C shows that a stray `gst_organizer_cache.json` in whatever folder the launcher is started from can decide which folder is scanned (RAN; A-R118).

Facts against Standard 13 (`rebuild_pipeline_FINAL.md:900-1046`; facts only): started with `python launch_extractor.py` from a console (README `:166-169`), not through a `.bat` with `pythonw.exe` (no `.bat` exists, READ `ls`); the app's files are found through `os.path.dirname(os.path.abspath(__file__))` appended to the path, not `sys.path.insert(0, APP_DIR)`; the log goes to the current folder, not `%LOCALAPPDATA%\<App Name>\`; no `sys.excepthook` / `threading.excepthook` (READ grep) — an error in the worker thread that escapes its `try` is printed to the console only; the launcher does show a message box when start-up fails (`:20-25`).

Rules involved: A-R118 (how the folder is found), A-R121 (extractor log file), A-R143 (panel behaviour).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:409-411`: "User launches launch_extractor.py → Loads target_folder from cache (BUG: wrong path - see F10)" — **no longer true** since `c1583d6` (A only (after the split)) (RAN case A).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:824-825`: "Window size: 1000×700, minimum 900×600; Theme: Clam ttk theme" — true (RAN).
- `README.md:166-176` (RM23-RM25): "python launch_extractor.py … auto-loads target folder from cache … Select clients … Click 'Refresh & Extract'" — the first two are true (RAN); the button is "🚀 Start Processing" (`extractor_window.py:330`).
- The design note `docs/superpowers/specs/2026-07-25-merge-extractor-design.md` keeps the standalone window as a supported way in (the panel "serves both this tab and the standalone launch_extractor.py window", `gui/tabs/extract_tab.py:4-6`).

#### History

`launch_extractor.py` and `extractor_main.py` are from `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) (RAN `git blame`), with 3 lines of `extractor_main.py` from `c1583d6` (23-07-2026, A only (after the split): home settings file first). `a677670` (25-07-2026, A only (after the split)) split the old window class into `ExtractorPanel` and the thin `PowerQueryExtractorApp` shell with `__getattr__` forwarding.

### A-F30: Application log file

#### What it does

Every module writes progress and error lines through Python's **logging** system (a standard way of recording what a program does). When the main program is loaded, `main.py` sends those lines at level INFO and above to a file `gst_organizer.log` **in the current folder** (the folder the program was started from) and to the console. The standalone extractor does the same with `pq_extractor.log` (A-F29). The files are appended to and never rotated or trimmed.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Main program: console re-wrapped as UTF-8, then `basicConfig(level=INFO, FileHandler('gst_organizer.log'), StreamHandler(sys.stdout))` — at IMPORT time | `main.py:16-17`, `:19-27` |
| Standalone extractor: `basicConfig(level=INFO, FileHandler('pq_extractor.log'), StreamHandler())` — at import time | `power_query_extractor/extractor_main.py:11-19` |
| Step 4 tab deliberately does not import `extractor_main`, so inside the main program everything goes to `gst_organizer.log` | `gui/tabs/extract_tab.py:8-10` |
| Typical INFO lines with client data | `core/excel_handler.py:195` (each Links value, first 50 characters); `power_query_extractor/core/report_processor.py:56` (every Step 4 screen line), `:590`, `:603-611` (every extracted figure) |

#### Input → Output

Log records from all modules → lines `<date time,ms> - <module> - <LEVEL> - <message>` appended to the log file and printed to the console.

#### User interaction

None. The user can open the file in a text editor (README `:289-291, 333`).

#### Concrete trace (RAN, me: `scripts/t30_logfile.py` → `out/t30_logfile.txt`; `scripts/t30b_basicconfig.py` → `out/t30b_basicconfig.txt`)

**Input, case 1:** a Python started in a scratch folder `started_from_here` imports the real `main.py` (which sets up logging at import, `main.py:17-27`; `main()` is not called), then runs one Step 4 ITC report of `<Client A>` through the real `ReportProcessor` with the fake COM layer (invented figure `Diff GSTR-3B!J15 = 1500.5`).

1. `main.py:17` re-wraps the console; `:20-27` creates `started_from_here/gst_organizer.log` (relative name → current folder).
2. The run writes 154 lines (RAN), levels INFO and WARNING, for example (RAN, as written):

```
… - power_query_extractor.core.report_processor - WARNING - pywin32 not available (No module named 'pythoncom'). …
… - power_query_extractor.core.report_processor - INFO - Opening Excel file: ITC_Report_ClientA_Traders_MH_211026_1030_Refreshed_211026_1343.xlsx
… - power_query_extractor.core.report_processor - INFO - ✅ Cell BB2 populated successfully: <Client A>
… - power_query_extractor.core.report_processor - INFO -     ✓ GSTR-3B_Difference_Diff_IGST (J15) = 1500.5
… - power_query_extractor.core.report_processor - INFO -   ✓ GSTR-3B_Difference_Diff_IGST: 1500.5
```
3. No log file in the program folder (RAN).

**Case 2 (RAN, SIMULATED Windows code page — the older character set Windows uses for plain-text files; cp1252 is the Western European one and has no ✓, ❌ or emoji):** the same format through a file handler opened with `encoding='cp1252'` — what Python uses on a Windows PC with that code page when, as here, no encoding is given (`main.py:24`, `extractor_main.py:16`) (MEDIUM: depends on the PC's language settings and Python's UTF-8 mode). Of three lines, only the one without symbols was written; the lines with "✓" and "❌" were lost, each with "--- Logging error ---" on the console (RAN). Applied to case 1's real log, **51 of its 154 lines** contain characters that cp1252 cannot write (RAN count).

**Case 3 (RAN `t30b_basicconfig.py`): README's "Debug Mode" advice** (`README.md:281-287`, "Enable detailed logging: `logging.basicConfig(level=logging.DEBUG)` in main.py or launch_extractor.py"):
- placed after `main.py:20-27`: root level stays 20 (INFO), the debug line does not reach the file — a second `basicConfig` does nothing once logging is set up;
- placed before the extractor's set-up: the extractor's own `basicConfig` then does nothing, so every line — INFO as well as DEBUG — goes to the console only (RAN: the console's last line was `INFO:x:info line`), and `pq_extractor.log` is created (its handler is built before `basicConfig` ignores it) but stays **empty** (size 0).

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G56): client data is logged at INFO into a file in whatever folder the program was started from** (RAN: placeholder client name from BB2, file names, extracted figures). Started from a shared program folder, every user appends to one file there. MEDIUM.

> **BUG (for Section 7; matches A-G56, which already describes the loss — this trace adds the size of it): on a Windows PC whose code page is cp1252, about a third of the log lines are silently lost** (RAN, simulated: 51 of 154 — all 154 lines of that log came from one Step 4 report, so about a third of the Step 4 lines; recounted by the verifier: 51 of 154), because the file is opened without an encoding and the messages contain symbols (✓ ❌ ⚠️ 📋). MEDIUM (depends on the PC's settings).

> **BUG (for Section 7; new — the README instruction does not work): README's "Debug Mode" line changes nothing where the README says to put it** (RAN case 3). LOW.

- **Start-up depends on the current folder being writable** (READ `main.py:20-27`: the file handler is created at import, before any window; a read-only start folder raises `PermissionError` and the program never opens) (A-G23; not runnable here as the root user).
- **Without a console** (`pythonw.exe`), `sys.stdout` is `None` and `main.py:17` fails before logging is even set up (READ; A-G22).
- **Never rotated** (READ: plain `FileHandler`, append mode): grows by about 150 lines per Step 4 report (RAN) plus the Step 3 lines (A-G52).
- **The history keeps an earlier log**: a 5,689-line `gst_organizer.log` with records of real processing runs was part of `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) and was deleted in `cbfa4a5` the same day (RAN `git show --stat`); because `6237597` (29-10-2025, shared) was built on `edac6a3`, not on `cbfa4a5`, the file is still in `6237597`'s snapshot and left the main line only with the merge `9c56a1c` (29-10-2025, shared) (RAN `git ls-tree`); its contents are not repeated here (described by kind only).

Facts against Standard 13 (`rebuild_pipeline_FINAL.md:1034-1040`; facts only): logs go to the current folder, not `%LOCALAPPDATA%\<App Name>\`; the app's files are not located from an `APP_DIR`; `print()` and console handlers are used for messages that, without a console, nobody sees.

Rules involved: A-R120 (main log file), A-R121 (extractor log file), A-R122 (what is logged and at which level).

#### Older documents against the code

- `README.md:289-291` "Logs are saved to: `gst_organizer.log` … `pq_extractor.log`" — true, but they are saved in the CURRENT folder, which README does not say (RAN).
- `README.md:281-287` "Debug Mode" — does not work as written (RAN case 3).
- `README.md:333` "Check the log file: `gst_organizer.log`" — the file may be in a different folder each time, depending on how the program was started (RAN).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:655` (item 9) "`main.py:36` creates `logs/` and `temp/` directories relative to `__file__`, which is correct" — those folders are created inside the program folder (`main.py:35-37`, READ), but nothing writes the log into `logs/`: the log goes to the current folder (READ `main.py:20-27`; RAN case 1).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:972-981` describes the committed log and its deletion — consistent with the history (RAN), except the date: it puts `cbfa4a5` (shared - before the split on 30-10-2025) under "Phase 2: Cleanup (cbfa4a5, 6237597, 9c56a1c — 2025-10-29)" (`:976`), while `cbfa4a5` is dated 31-08-2025 22:14, the same day as `edac6a3` (RAN `git log --format=%ad/%cd`: author and committer date agree); only the merge `9c56a1c` (shared - before the split on 30-10-2025) that brought the deletion into the main line is from 29-10-2025.

#### History

`main.py:1-30` and the extractor's logging set-up are from `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) (RAN `git blame`). The comment in `gui/tabs/extract_tab.py:8-10` came with the Step 4 tab's own commit `bfcb24a` (26-07-2026, A only (after the split); RAN `git blame`: all three lines `bfcb24a`; see A-F18). [Corrected during verification — originally named `a677670` as well, which did not touch this file.]

### A-F31: Help texts, tooltips, scrolling and collapsible sections

#### What it does

Small aids around the screens: **tooltips** (small pale-yellow boxes that appear when the pointer rests on a control) on Step 1; hover help texts for each checklist row, workflow choice, processing mode, the client-name options and the Scan buttons; an "EXPECTED FILE NAMES" card listing the six file-name formats; **mouse-wheel scrolling** of the Step 1 page and of Step 4's client list; and on Step 2 a **collapsible** "📋 VALIDATION INSTRUCTIONS" box that hides and shows with a ▼/▶ button.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Tooltip widget: 450 ms delay, wrap at 520 px, pale yellow `#FFFFE0`, 12 px right of and 6 px below the control | `gui/widgets/tooltip.py:5-69` |
| Hover help texts | `gui/tabs/setup_tab.py:71-133` (checklist `:73-86`; Browse `:88-93`; workflow `:95-101`; processing modes `:103-110`; include name `:112-117`; length `:119-123`; Scan `:125-128`; Re-scan `:130-133`) |
| Where tooltips are attached; path tooltips show the full path and follow it | `:314, 336, 381, 386, 409, 423, 444, 462, 467`; `:570-576` |
| Step 1 wheel: one application-wide binding, 22 px per unit, 3 units per notch, small movements added up | `:137-138`, `:193-201`, `:220-248` |
| "EXPECTED FILE NAMES" card, built from the `description` strings of the patterns | `:477-503`; `utils/constants.py:13-62` [GST RULE] ⚠ |
| Step 4 client-list wheel (same method as Step 1) and log wheel (A-F25) | `power_query_extractor/gui/extractor_window.py:433-478`; `:558-565` |
| Collapsible box and its use on Step 2 | `gui/widgets/collapsible_frame.py:7-41`; `gui/tabs/validation_tab.py:44-66` |

#### Input → Output

Pointer rests / wheel turns / ▼ click → popup shown / page scrolled / box hidden or shown.

#### User interaction

Hovering, the mouse wheel or touch-pad, clicking ▼/▶.

#### Concrete trace (RAN, me: `scripts/t31_help.py` → `out/t31_help.txt`, real main window on a virtual display)

**Input:** the user rests the pointer on "Scan Files"; rolls the wheel over the Step 1 page (one notch, then three small touch-pad steps, then a Linux-style wheel click); opens Step 4 with 40 listed clients and rolls the wheel once over the list; clicks the ▼ of "VALIDATION INSTRUCTIONS" twice.

| Action | What happened (RAN) |
|---|---|
| Start-up | 33 tooltips attached on Step 1 (READ count: 4 checklist rows × 4 widgets, 2 workflow, 4 path boxes, 4 Browse, 3 processing modes, include-name, length, Scan, Re-scan); 18 different texts, one of which is the empty text shared by the four path tooltips, which stay empty — and show nothing — until a path is chosen (`tooltip.py:32`) |
| 0.3 s after the pointer enters "Scan Files" | no popup |
| 0.7 s after | popup "Read the source folder and work out which client each file belongs to.\n\nRead-only - nothing is copied or changed. Results appear on Step 2.", `#FFFFE0`/`#333333`, at (76, 1341) for a button at (64, 1303) 32 px high — i.e. +12, +32+6 |
| Pointer leaves | popup gone |
| Tooltip for a control 60 px from the right screen edge (screen 1920×1080) | popup from x 1872 to **2402** and y 1076 to **1112** — runs off the screen (the verifier's re-run had a 1280×1024 virtual screen: popup from x 1232 to **1762**, y 1020 to **1056** — the same result) |
| Step 1 page, one wheel notch (`delta -120`) | view 0.000 → 0.043 |
| three touch-pad steps (`delta -40` each) | 0.043 → 0.058 → 0.072 → 0.087 (each step = one unit, thanks to the adding-up) |
| Linux wheel click (`<Button-5>`) | no movement |
| Step 4 list, 40 rows, one notch | 0.000 → 0.104 |
| Step 2 box | outer frame `#FFF3E0` (orange), but title strip, ▼ button and inner area `#d9d9d9` (Tk's default grey); texts "• Review the scan summary below", "• Check client list - ✅ = complete, ⚠️ = missing files", "• Select clients using click or keyboard (↑↓ + SPACE)", "• Try 'Dry Run' to preview without processing", "• Click 'Start Processing' when ready" |
| ▼ clicked | button "▶", inner area hidden; clicked again → "▼", shown |
| Buttons that exist (whole window) | "🧪 DRY RUN", "🚀 START" (Step 2), "🚀 START PROCESSING" (Step 3), "🚀 Start Processing" (Step 4) |

#### What the trace shows against what the feature claims

> **BUG (for Section 7; matches A-G66): several help texts promise behaviour the code does not have** (READ, each against the code):

| Help text | Where | What the code does |
|---|---|---|
| "Continue a run that was interrupted, skipping files that are already in place." | `gui/tabs/setup_tab.py:108-109` | re-copies everything unless in the same minute (A-F13 in part 1; A-G10) |
| "Fast - no Excel automation involved." (organise only) | `:96-97` | Step 3 drives Excel through COM on Windows (`core/excel_handler.py:90-95`; A-F14/A-F15 in part 1) |
| "…and then open each report in Excel to refresh its Power Query and pull the figures out." ('full') | `:98-101` | 'full' only switches to Step 4 and scans; the refresh still needs "🚀 Start Processing" (`gui/handlers/processing_handler.py:139-142`; A-F18) |
| "Each run adds a new timestamped folder, so previous runs are left untouched." | `:83-85` | untrue for Re-run, Resume and runs in the same minute (A-G9) |
| "Longest client name allowed in a folder name before it is shortened." | `:119-123` | shortens only one folder level (A-F6 in part 1; A-G21) |
| Example "'ITC (ABC Pvt Ltd)'" (a generic example name) | `:112-117` | real folders are written with underscores, `ITC (ABC_Pvt_Ltd)` (A-G66) |
| "Click 'Start Processing' when ready" (Step 2) | `gui/tabs/validation_tab.py:57` | Step 2's own button is labelled "🚀 START" (`:133`) (RAN button list) — LOW |

> **BUG (for Section 7; matches A-G68): the wheel is bound to `<MouseWheel>` only** — a Linux-style wheel click does nothing (RAN); and the Step 4 log drops small movements (A-F25, RAN). On Windows the Step 1 page and Step 4 list work, including small touch-pad steps (RAN). Whether a real wheel under Tk 9 on Linux also sends `<MouseWheel>`: NOT SURE (LOW; irrelevant for Windows users).

- **Tooltips are not kept on the screen** (RAN edge test; `tooltip.py:47-48` adds +12/+6 without a check) — near the right or bottom edge part of the text is off-screen. LOW.
- **The collapsible box colours only its outer frame** (RAN): grey strips around the orange instruction lines. LOW.
- **The "EXPECTED FILE NAMES" card shows hand-written descriptions, not the patterns themselves** [GST RULE] ⚠: the comment says "read from FILE_PATTERNS so this card can never drift" (`setup_tab.py:478-479`), but it shows the `description` strings (`utils/constants.py:19, 27, 35, 43, 51, 59`), which say `.xlsx` while every pattern also accepts `.xls` and any letter case (`re.IGNORECASE`), and do not mention the optional "(N)" copy markers (after the IMS date, before `GSTR3B`) or that the IMS date must be exactly eight digits (`:23, :31`) (READ). Which file names are acceptable is the author's rule; only the difference between card and pattern is stated here.
- Tooltips, the help card and the collapsible box are not themed by dark mode (A-F28).

Rules involved: A-R146 (mouse-wheel scrolling), A-R147 (tooltips), A-R1 to A-R7 (file-name patterns behind the card) [GST RULE] ⚠.

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:794-800`: "File patterns card (purple): 6 expected file name formats … Collapsible instructions (orange)" — true (RAN).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:1032-1035` (history of `4c42758`, shared - before the split on 30-10-2025): "Focus-based binding prevents scroll conflicts" — that Enter/Leave method was replaced (it switched off other tabs' scrolling; comment `extractor_window.py:433-446`) by the add-up method in `bf2f1dd` (A only (after the split)).
- `tests/README.md` and the design note do not describe the tooltips.

#### History

The collapsible box and the Step 2 instructions are from `edac6a3` (31-08-2025, shared - before the split on 30-10-2025); the first wheel handlers for Step 4 came with `4c42758` (29-10-2025, shared - before the split on 30-10-2025). A only (after the split): `322717f` (25-07-2026: tooltip widget, Setup restructure, "EXPECTED FILE NAMES" card), `d64d3a9` (25-07-2026: Step 1 scrolling), `169c7b5` (26-07-2026: the 7-line `_is_dark` helper at `gui/tabs/setup_tab.py:204-210`, inside the scrolling code), `bf2f1dd` (27-07-2026, "Fix scrolling dying after visiting other tabs; add hover help": 64 lines of the help texts and 42 lines of the Step 4 list wheel) (RAN `git blame`).

### Findings from the traces in this part, for Section 7

Every difference between what a feature claims (its labels, help texts, messages, comments or the older documents) and what the trace showed. "Evidence" says how it was found; "Section 7" gives the matching gap (A-G#) or says **new** where Section 7 has no gap for it yet. Severity words are this part's (HIGH / MEDIUM / LOW). [GST RULE] marks findings that touch how GST data is chosen or shown; nothing here says what the right GST treatment is.

| # | Feature | Finding | Evidence | Section 7 |
|---|---|---|---|---|
| 1 | A-F16 | The per-client organisation report shows run-wide "Successfully Copied" / "Folders Created" counts (Client C: 20 / 16 for 6 files) | RAN t16 R1 | A-G19 |
| 2 | A-F16 | After Stop, the summary workbook says "Total Clients 3, Successful 1, Failed 0"; the unprocessed clients appear nowhere and nothing says "stopped" | RAN t16 R4 | **new** |
| 3 | A-F16 | When every client fails there is no workbook, but the run "completes" and the box says "Successfully processed 0/3 clients" | RAN t16 R3 | A-G69 (the box wording); the missing workbook is **new** |
| 4 | A-F16 | Client Status lists successful clients only; "Success" means "no exception", not "both reports made" | RAN t16 R2; READ | A-G19 |
| 5 | A-F16 | "Complete %" is written as text (`'83.3%'`) and "Processing Date" as the run-stamp text | RAN t16 | **new** (LOW) |
| 6 | A-F16 | The organisation report's file list is every file of the client, including files whose copy failed | READ | **new** (LOW) |
| 7 | A-F17 | "Successfully processed p/t clients" after a Stop or failures | RAN t16, t17_gui | A-G69 |
| 8 | A-F17 | Greyed STOP / START still react to clicks; neither works from the keyboard | RAN t17_gui | A-G62 |
| 9 | A-F17 | Closing the window mid-run kills the worker about 1 s later | RAN t17_close | A-G27 |
| 10 | A-F17 | At the summary step the ETA still shows one client's time (the "one client behind" effect is only there) | RAN t16 | A-G72 (clarified) |
| 11 | A-F17 | "📂 Opened output folder" logged when opening failed (macOS/Linux) | RAN t16 | A-G69 |
| 12 | A-F18 | After the 'full' hand-over Step 4 no longer follows Step 1's target | RAN t18_tab | A-G65 |
| 13 | A-F18 | Step 4's status messages are invisible inside the main window | RAN t18_tab | A-G65 |
| 14 | A-F18 | The Step 4 fallback notice blames pywin32 for any import failure and can no longer appear for pywin32 itself | RAN t18_notice | A-G66 |
| 15 | A-F18 | "✅ Step 4 ready" is logged even when the Step 4 scan found nothing | RAN t18_tab | A-G66 (related; the "found nothing" case is **new**) |
| 16 | A-F19 | "Newest run folder" = most recently changed: opening an older run's summary workbook in Excel switches Step 4 to that older run | RAN t19b | A-G45 |
| 17 | A-F19 | Pointing the box at a run folder itself finds nothing and keeps the old list | RAN t19 | A-G45 |
| 18 | A-F19 | Only the newest Version folder of each client is looked at | RAN t19 | **new** (MEDIUM) |
| 19 | A-F19 | A file named like a run folder breaks the scan and leaves the run folder pointing at the file | RAN t19 | A-G29, A-G45 |
| 20 | A-F20 | A failed refresh is shown as "last refresh", dated with the original's time | RAN t19b | A-G2 |
| 21 | A-F20 | With suffix `_R{timestamp}` every original counts as refreshed | RAN t19 | A-G16 |
| 22 | A-F20 | A client with no Sales-type files gets a pre-ticked Sales row | RAN t19b | A-G5 |
| 23 | A-F20, A-F25 | Any scan re-enables "🚀 Start Processing" during a run; a second worker can be started | RAN t19 | A-G65 |
| 24 | A-F21 | "Refreshed successfully" rests on `Info!BB2` alone; it passes even when the key presses never reached Excel ⚠ [GST RULE] | RAN (fake COM) Part 6; READ | A-G3 |
| 25 | A-F21 | A failed refresh is saved and kept under the "Refreshed" name | RAN (fake COM) Parts 2, 3; READ | A-G2 |
| 26 | A-F21, A-F23 | `Dispatch` may attach to the user's own Excel, which is then keyed into, hidden and quit; keys go to whatever window is in front | READ | A-G54, A-G55 |
| 27 | A-F21 | After three "rejected by callee" failures the code presses keys with no workbook open; error "'NoneType' object has no attribute 'Save'" | RAN (fake COM) Part 3 | A-G73 |
| 28 | A-F21 | Fixed waits: 39 s per report at the default 10 s (136 s + 3 s at 60 s), two Excel sessions per report, no Stop | RAN (count); READ | A-G50 |
| 29 | A-F21 | The wait box accepts 0 and text (text makes Start silently do nothing) | RAN (PQ gui_probe); READ | A-G26 |
| 30 | A-F21 | An `except` that can never run (`report_processor.py:169-181`) and an unreachable `return` (`:383-386`) | READ | A-G74 |
| 31 | A-F22 | Skip Refresh reuses a copy whose refresh failed and reports success | RAN t22 1a-1b | A-G2 |
| 32 | A-F22 | Suffix problems: `_R{timestamp}` makes originals "refreshed"; `{TIMESTAMP}` is written literally; `{timestamp}` alone makes copies invisible; an empty suffix fails every report | RAN t22 2-5 | A-G16 |
| 33 | A-F22 | The "original" can be an earlier refreshed copy (first file listed) | RAN t22 1b, 6 | A-G15 |
| 34 | A-F22 | `auto_refresh_if_missing` and `show_timestamp` are never read; the Skip tick rewrites a shared option in memory | READ | A-G74 (related) |
| 35 | A-F23 | A sheet named with other letter case (`reco`) gives `None` for all its fields; the warning goes to the log file only [GST RULE] | RAN (fake COM) Part 1 | A-G17, A-G67 |
| 36 | A-F23 | The log says "Opening file in read-only mode" / "File opened in read-only mode"; it is opened read-write | RAN (fake COM); READ | A-G66 |
| 37 | A-F23 | Excel error values may arrive as large negative numbers (not testable here) [GST RULE] | READ (MEDIUM) | A-G18 |
| 38 | A-F24 | 9 of 22 configured columns ("IMS Reco", "Extra In Purchase", "Extra In GSTR2B": IGST/CGST/SGST) are filled from `Diff GSTR-3B!J15/K15/L15` ⚠ [GST RULE] | RAN t24 | A-G1 |
| 39 | A-F24 | ✓ follows the refresh status only; any non-success status shows "Missing" in every column, even for values read; a ✓ row can hold only "Missing" [GST RULE] | RAN t24 | A-G17 |
| 40 | A-F24 | One save error loses the whole run's figures; a second run in the same minute replaces the workbook | RAN t24 | A-G29, A-G9 |
| 41 | A-F24 | A re-scan during a run sends the workbook to the folder scanned last | RAN t19 | A-G65, A-G71 |
| 42 | A-F24 | Text numbers stay text, zero is not formatted, `True` gets the number format (`False`, equal to 0, does not — READ) [GST RULE] | RAN t24 | A-G18 |
| 43 | A-F24 | Summary counts are refresh successes over ALL clients; "not ticked" counts as "not processed" | RAN t24, t19 | A-G19 |
| 44 | A-F25 | Every processor line is "info" black, errors included; warnings written with `logger.warning/error` never reach the screen | RAN t19, t21 | A-G67, A-G65 |
| 45 | A-F25 | The log box can be typed into; it is emptied at every Start | RAN t19; READ | A-G65 |
| 46 | A-F25 | The log wheel ignores small movements (`delta -60` → no scroll) | RAN t19 | A-G68 |
| 47 | A-F26 | Every desktop window is examined; another program's dialog is answered "Yes" | RAN (fake COM) Part 4; READ | A-G55 |
| 48 | A-F26 | Any matching dialog marks the report failed, so its figures show "Missing" although the refresh check passed | RAN (fake COM) Part 4 | A-G17, A-G19 |
| 49 | A-F26 | When the BB2 check also fails, its message replaces the dialog text, so the dialog's text is lost from the status | READ (`report_processor.py:353-356`) | **new** (LOW) |
| 50 | A-F27 | A JSON list or a length stored as text stops start-up; text in the length box makes every save and Browse fail | RAN t27 7, 8b, 8c | A-G24, A-G26 |
| 51 | A-F27 | Typed paths, the processing-mode radios and closing the window do not save; a broken file silently loads as defaults | RAN t27 2, 3, 5, 8a | A-G24 (related); A-R116 — the "not saved" part is **new** as a gap |
| 52 | A-F27 | An unknown processing-mode value (`Fresh`) is accepted at start-up, leaves no radio selected and makes a later Step 3 run fail with "Invalid processing mode" | RAN t27 8d; READ | **new** (LOW) |
| 53 | A-F28 | Widgets built while dark stay light; a scan in dark mode gives white text on pale rows | RAN t28 2 | A-G64 |
| 54 | A-F28 | The first dark/light round changes the selected-row and selected-tab colours for good | RAN t28 3 | A-G64 |
| 55 | A-F28 | Colours changed in light mode after the first switch are lost on restore — latent: no such widget in A today besides the Setup checklist, which is recomputed | RAN t28 4 (stand-in); READ | A-G64 (latent) |
| 56 | A-F28, A-F29 | The standalone extractor ignores the saved dark-mode setting | RAN t29 A | **new** (LOW) |
| 57 | A-F29 | Closing the standalone window — or the main window during a Step 4 run — ends the run without a question | READ | A-G27 |
| 58 | A-F29 | `python power_query_extractor/extractor_main.py` fails (relative import) | RAN t29 D | A-G32 |
| 59 | A-F29 | The window's `__getattr__` replaces Tk's own forwarding (`app.eval` fails) | RAN t29b | **new** (LOW; nothing in A uses it) |
| 60 | A-F29 | A stray `gst_organizer_cache.json` in the current folder decides which folder is scanned when the home file is missing | RAN t29 C | A-G32 (related; **new** as such) |
| 61 | A-F30 | Client data is logged at INFO into a file in the current folder | RAN t30 1 | A-G56 |
| 62 | A-F30 | With a cp1252 code page about a third of the lines are lost (51 of 154 for one Step 4 report) | RAN t30 2 (simulated) | A-G56 |
| 63 | A-F30 | README's "Debug Mode" line changes nothing where README says to put it; before the extractor's set-up it leaves `pq_extractor.log` empty | RAN t30b | **new** (LOW) |
| 64 | A-F30 | An unwritable current folder stops start-up; without a console `main.py:17` fails | READ | A-G23, A-G22 |
| 65 | A-F31 | Help texts that promise behaviour the code does not have (Resume, "no Excel automation", 'full', "previous runs untouched", length, name example, "Start Processing" on Step 2) | READ; RAN (button list) | A-G66 |
| 66 | A-F31 | Wheel bound to `<MouseWheel>` only | RAN t31 | A-G68 |
| 67 | A-F31 | Tooltips are not kept on the screen near its edges | RAN t31 | **new** (LOW) |
| 68 | A-F31 | The collapsible box colours only its outer frame (grey strips around orange lines) | RAN t31 | **new** (LOW) |
| 69 | A-F31 | The "EXPECTED FILE NAMES" card shows hand-written descriptions that omit `.xls`, letter-case freedom, "(N)" markers and the exactly-8-digit IMS date, although its comment says it "can never drift" ⚠ [GST RULE] | READ | **new** (LOW; A-G38 is related) |

### Trace scripts and outputs used in this part

All under `/root/merge-scratch/run/A-S3b/` (scripts in `scripts/`, outputs in `out/` with the same name and `.txt`); "fake COM" = `scripts/fakecom.py`, which records calls and reads workbooks with openpyxl and is **not** Excel. Verification: every script was copied to `/root/merge-scratch/run/VERIFY-A-S3b/scripts/` (only the scratch paths changed) and re-run there on 21-10-2026, 14:43-14:51, against a fresh copy of the snapshot (`diff -rq` against `/root/merge-scratch/A-copy` empty; checksums of the copy identical before and after the runs). After masking clock times, run stamps and scratch paths, the outputs agree with the ones quoted here except for the byte sizes of made-up files and the virtual screen size in `t31_help.py` (see A-F16, A-F31).

| Script | Features | What it drives |
|---|---|---|
| `common.py` | all | shared set-up: scratch paths (asserts that `HOME`/`TMPDIR` point into the scratch folder), workbook helpers, log capture |
| `fakecom.py` | A-F21 to A-F26, A-F30 | the fake COM layer and fake window list |
| `t16_17_step3.py` | A-F16, A-F17 | the real Step 3 worker, runs R1-R4 |
| `t17_gui.py`, `t17_close.py` | A-F17 | buttons, ETA, end box; closing mid-run (virtual display) |
| `t18_tab.py`, `t18_notice.py` | A-F18 | Step 4 tab hand-over and folder sync; the fallback notice |
| `t19_20_25_panel.py`, `t19b_scan_realistic.py` | A-F19, A-F20, A-F24, A-F25 | the real Step 4 panel: scans, rows, Start, worker log, re-scan during a run |
| `t21_23_26_com.py` | A-F21, A-F23, A-F26 | the real `ReportProcessor` through the fake COM layer, Parts 0-6 |
| `t22_naming.py` | A-F22 | the real `_process_report` naming / Skip logic |
| `t24_consolidator.py` | A-F24 | the real `DataConsolidator.create_report` with marker values |
| `t27_cache.py` | A-F27 | the real main window's settings saves and restores |
| `t28_dark.py` | A-F28 | the real main window's dark mode |
| `t29_launcher.py`, `t29b_getattr.py` | A-F29, A-F30 | the real launcher as a script in separate current folders |
| `t30_logfile.py`, `t30b_basicconfig.py` | A-F30 | `main.py`'s log set-up and one Step 4 report; cp1252 simulation; README's Debug Mode line |
| `t31_help.py` | A-F31 | tooltips, wheel, collapsible box in the real main window |

