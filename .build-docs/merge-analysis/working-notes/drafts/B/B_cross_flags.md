# Cross-section flags from the section writers (project B)

Each writer listed facts that OTHER sections must reflect, items the registry lacks (UNREGISTERED), and corrections to the working notes. The checkers have since corrected some of these in place; confirm each in the code before acting.

## From writer S3a

### Facts other sections must reflect
- B-F1 (RAN): the window cannot start without pywin32. requirements.txt:14 leaves it commented out, and main.py:51-53 recommends 'pip install PyQt5 openpyxl', which does not fix it.
- B-F1/B-F17 (RAN): B has no sys.excepthook, so an uncaught error in any window handler aborts the program (exit 134) without a message or a settings save. Examples: a text recent_source_folders in session_state.json; {"runs": 5} in run_history.json, which aborts at the end of processing.
- B-F1 (RAN): one wrongly typed value in settings.json (pq_wait_time "10") stops the window from being built.
- B-F2 (RAN): the migration reads 'recent_source_folders' but v3 wrote 'recent_folders'. A migrated max length of 0 becomes an effective 15-character client-folder limit (spin-box minimum), where v3 used 35.
- B-F3 (RAN): a path pasted with quotes gives 'Folder not found'. A template path pointing to a folder passes the check, then every ITC report fails while the run still says 'Successfully processed 1/1'. Three of the four Browse buttons are invisible (white on white).
- B-F4 (RAN): the greyed-out 'Max length' box still cuts the client folder name when the include-name tick is off.
- B-F6 (RAN): a hyphen inside a client name splits it silently (client 'Client', state 'D Exports', key 'Client-DE') or makes the name fail; state spelled two ways ('Tamil Nadu'/'Tamilnadu' gives codes TN/TAM) makes two clients. [GST RULE] state-code logic.
- B-F6 (RAN/READ): .xlsm sources are collected but never recognised; files in sub-folders are ignored silently; a file whose size cannot be read is dropped from every list; RE-SCAN during a scan clears the log and is silently refused.
- B-F7 (RAN): the on-screen counts do not add up. The log lists [1/28]..[28/28], but 'Files found: 27' and the status bar says '27 files'; the rejected file is only in 'Variations: 3'. The HTML scan log slows the window down as it grows (1,000 lines 5.6 s, 5,000 lines 136.6 s, VW).
- B-F8 (RAN): every client is ticked after a scan, complete or not. Number columns sort as text (1, 1, 15, 3, 5) and processing follows that order. 'Folder Name' shows the scan key, not the real (abbreviated or cut) folder.
- B-F9 (RAN): broken. Problem files are never named to the user, and the processing summary's Variations sheet says 'No file variations found' although the scan found 3 (processing_controller.py:73 passes []). This is a regression from 47a1492.
- B-F10 (RAN): the EXPORT LIST button has 0 connected receivers; clicking it does nothing.
- B-F11 (RAN): the dry run checks no templates, target or mode, and promises both reports for every client. After a processing run without a re-scan it shows cleaned names, because core/file_organizer.py:216-229 rewrites the window's scan data.
- B-F12 (RAN): Resume does not continue an interrupted run. It makes a new minute-stamped Version folder and copies everything again; it skips only within the same clock minute. Re-Run overwrites only within the same minute. The three on-screen mode descriptions disagree with each other and with the code.
- B-F12 (RAN): two source files whose cleaned names coincide ('… Apr 1' and '…_Apr_1') keep one copy while both report Success; the first survives only as a _backup_ (Fresh) or is lost (Re-Run).
- B-F12/B-F18 (RAN): Re-Run/Resume choose the batch folder by newest modification time, so touching an older batch folder sends the next run there. Step 4 chooses by reverse name instead: two different 'latest' rules.
- B-F12 (READ/LT RAN): partly made folders stay when a client fails (cleanup_on_error is never called); same-minute backups overwrite each other silently.
- B-F13 (fake-COM RAN): Step 3 attaches to an Excel the user already has open (GetActiveObject), hides it and always calls Quit() on it. The comment at core/excel_handler.py:206 says the opposite. It does this twice per client (ITC, then Sales).
- B-F13 (fake-COM RAN): if Excel dies during open, the finally block raises at core/excel_handler.py:204, so Quit() and CoUninitialize() are skipped. The unfilled template copy stays under the real report name and Step 4 later treats it as a report.
- B-F13/B-F14 (RAN): when a client has no file of a type, the template's own Links cell content stays [GST RULE] ⚠. A missing Links sheet is created empty (COM) or left unwritten (openpyxl), and success is still reported. A report is made even when the client has none of its inputs.
- B-F13/B-F14 (RAN): only the first file of each type is linked; e.g. two half-year Sales files, only 'Apr-Sep' named in Links!B3 [GST RULE]. After a skipped or failed copy, the Links cell names the uncleaned original file, which is not in the folder.
- B-F15 (RAN): Client Status lists only successful clients and hard-codes 'Success'. Report-creation failures never appear in the Errors sheet. Each organisation report's 'Successfully Copied' and 'Folders Created' are running totals of the whole batch (6/12 for a 5-file client). A failed summary workbook is silent on screen.
- B-F16 (RAN): Stop takes effect only between clients, and a stopped run ends as a normal completion ('Successfully processed 1/5 clients', 100%, move to Step 4; run history failed_clients 0). The HTML Step 3 log cuts text after '<' ('❌ ITC Report Error: '') and drops indentation and blank lines. The log is never cleared.
- B-F17 (RAN): 'Last Run' appears only after the next scan. Per-client history is never trimmed (55 records kept, runs capped at 50). itc_report_path, sales_report_path and errors are always empty.
- Older documents (README.md) disagree with the code: Resume 'Continue interrupted processing' (142-145); ✅/⚠️ icons (150-152); 'variations tab' (156, 299); Dry Run on the Processing tab (158-159); output name formats DDMMYYYY_HHMMSS (182, 188-189, 195, 202-203; code uses DDMMYY_HHMM); Variations sheet (200); 'temporary files' (42, 260); 'no external links preserved' (44, 262); 'Template not found → Links sheet' (278).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- MW notes (B_code_mainwindow_controllers_theme.md, _on_processing_completed) say the calls in that handler are 'mostly self-protecting', except for a non-list recent_source_folders. RAN in this task (scripts/t_f17_badhistory.py): a run_history.json holding {"runs": 5} makes RunHistory.add_run raise AttributeError at models/run_history.py:58, called from gui/pyqt_main_window.py:379. PyQt5 then aborts the whole program (exit code 134) after all the processing output has been written. So that list of exceptions is incomplete.
- Registry B-F17 and B-R119 (and LT §5.2 #32-35) say that wrong inner types in run_history.json 'crash the next save'. The crash actually happens when the run is recorded (add_run, models/run_history.py:58, or add_extraction_result), before any save. In the window it is not just an error: it aborts the program at the end of processing (RAN).
- Registry B-F6 lists GSTR-2B-Reco, SalesReco and AnnualReport as the patterns that mis-split a hyphenated client name, and says others 'fail'. RAN in this task (scripts/t_f6_hyphen.py) confirms this and adds the exact split: those three accept client 'Client', state 'D Exports' and push the real state into the period or year. ImsReco, GSTR3B and Sales give 'No matching pattern found', so one client's files are split between a wrong client and the problem list. This is an addition to the note, not a contradiction.

## From writer S67

### Facts other sections must reflect
- Gap IDs B-G1 to B-G82 are defined in S67.md section 7.10 (register). Sections 8, 9 and 10 and the comparison should cite them.
- try/except inventory for B: 90 try statements, 91 handlers, 3 finally blocks, 0 bare except. 15 handlers only 'pass'; 45 of the 91 never reach the screen; 15 are in dead code.
- Sections 3 (B-F23) and 4 (data flow): a consolidated-report save failure leaves an unopenable half-written PQ_Extraction_Report_*.xlsx under the normal name, and a same-minute run would be replaced by it (RAN).
- Section 3 (B-F12) and Section 5 (B-R43): with a cleaned-name collision, Fresh and Re-Run keep the SECOND file's content, while Resume keeps the FIRST and marks the second 'Skipped'. The skipped file's name stays uncleaned, so the Links cells name a file that does not exist (RAN).
- Section 3 (B-F27, B-F28) and Section 10: a damaged ~/.gst_organizer/session_state.json (a JSON list) makes switching dark mode, scanning or closing the window abort the program (exit 134), and the settings are not saved (RAN).
- Section 10: the window's minimum size is 1200x800 (gui/pyqt_main_window.py:72), taller than a 1366x768 laptop screen.
- Section 9 (dependencies): cleaned client names depend on the Python version (3.11-3.13 against 3.14, trailing dot). requirements.txt gives only minimum versions.
- Section 3 (B-F19, B-F26), with recording stand-ins: keystrokes are sent when activation fails and when no workbook opened; a dialog in another program had a button containing 'ok' ('Book a demo') clicked; an Excel question was answered 'Yes' (RAN me r5).
- Section 3 (B-F15): with a folder given as both templates, the real ProcessingWorker reported 'Processing complete: 1/1', Client Status 'Success' and an Errors sheet saying 'No errors encountered', while both reports had failed (RAN me r8).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- LT §2.3 item 1 left open whether older Python versions keep a trailing dot in sanitize_filename. RAN here: Python 3.11, 3.12 and 3.13 drop it ('Alpha Test Traders.' becomes 'Alpha_Test_Traders'); 3.14 keeps it ('Alpha_Test_Traders.'). Now gap B-G80.
- Registry B-F12 and LT §6.5 say two source names that clean to the same name 'keep only one copy while both report Success'. That holds for Fresh and Re-Run. In Resume mode the second file is 'Skipped / File already exists' and the FIRST file's content is kept (RAN me r2 run4).
- PQ and LT §4.10 say a time-zone-aware date makes the consolidated report fail. Additionally, a half-written PQ_Extraction_Report_*.xlsx (5,125 bytes, no [Content_Types].xml, openpyxl cannot open it) is left on disk, and within the same minute it would replace a good earlier report (RAN me r3).
- XL §5.2 step 6 counts 12 open attempts when Excel rejects every open. That is right only on Windows when a separate 8.3 short path exists. Without one it is 9 (RAN me r5; agrees with the Section 3 t19b run).
- Section 3 (B-F2) said, as READ/LOW, that the migration marker is still created when settings.json cannot be written. RAN me r4 confirms it: the marker is set and 'Migrated settings from old cache file' is logged at INFO although settings.json was not written.
- The earlier interrupted draft of this section (run/B-S67/parts/p1_s6.md) had two errors, both fixed in S67.md. (1) It said the consolidated-report error is 'not in the program log', but data_consolidator.py:172-174 logs it with a traceback before re-raising. (2) It said the second wb.Close in report_processor.py's finally 'always fails'; that is an expectation for real Excel, now marked MEDIUM.

## From writer S12BF

### Facts other sections must reflect
- Start-up history: from c258047 the window cannot open without pywin32. The chain is main.py:38 -> gui/__init__.py:7 -> gui/pyqt_main_window.py:27 -> gui/controllers/extraction_controller.py:10 -> power_query_extractor/core/report_processor.py:11-15, an unguarded import. I ran it on extracted trees: fails without pywin32 at c258047, c21f788 and b0e7d0e; the 47a1492 organiser imports fine without it. At a62d3cc the program could not start at all, even with stand-in pywin32 modules ('No module named gui.main_window'). Sections 1, 9 and 12A should match this.
- VERIFICATION RULE result: no file imports a name first added after the file's own last change. The only reverse case is power_query_extractor/extractor_main.py:9, which imports a file deleted in a62d3cc and has been broken since. gui/widgets/__init__.py imported deleted files from a62d3cc until 7e09e5c. Any 'never changed' claim in Section 2 or 12A should agree.
- No GST rule value (file patterns, expected types, state table, template Links cells, Info!BB2 check, extraction cells) changed in any B-only commit; utils/constants.py and power_query_extractor/config/cell_mappings.py were never touched after 47a1492. Sections 5 and 12A should not attribute any GST rule change to B's own commits.
- Scan acceptance changed in 7e09e5c, which I confirmed by running a scan with made-up files on four trees. At 47a1492 and c21f788, files under 1 KB or with a wrong signature vanish silently and a valid ~$ lock-file name is listed as unrecognised. At 7e09e5c and b0e7d0e, the bad files are listed as 'File too small to be valid Excel' / 'Invalid or corrupted Excel file' and ~$ names are skipped. A still drops such files silently (per H-B2), which matters for the comparison.
- The Step 4 run-folder choice was replaced three times in three days (c258047 whole-target name search, 7e09e5c run history plus reverse name order, b0e7d0e greatest name). The original Tk extractor chose by modification time (47a1492:power_query_extractor/gui/extractor_window.py:517). Section 8 design decisions should not present the snapshot rule as a long-standing choice.
- The consolidated report location moved from the Level 1 folder (Tk extractor) to the target folder (c258047) and back to the Level 1 folder (b0e7d0e). Its row label went from client folder name to '<client> - <state>' to '<client folder> / <version folder>'.
- Rules lost in the PyQt5 rewrite (c258047/a62d3cc), for Sections 3, 7 and 11: long-client-name confirmation, per-client include-name toggle, client-list export (button left unconnected), Variations sheet content, Dry Run input checks, the 'No valid GST files found' warning, the close-while-processing question, opening the output folder on macOS/Linux, settings saved on every change, log timestamps, the .xltx template picker, the extractor's 'Open the consolidated report?' prompt, and per-report refresh status.
- B has never had tests, pytest.ini, lint configuration or CI at any commit; c258047 removed even the commented dev-tool lines from requirements.txt.
- README.md was last changed in d674386 (shared) and describes v3.0 Tkinter. Its OS line 'Windows 10/11 (best), macOS, Linux' (README.md:51) has been false since c258047, because the window needs Windows-only pywin32 to start.

### UNREGISTERED items
- (none)

### Corrections to the working notes
- shared_history_12_commits.md (6237597 entry): says core/file_parser.py adds self.progress_callback at line 30 and _update_progress at 49-55@6237597; in the code at 6237597 they are line 29 and lines 48-54 (READ git show 6237597:core/file_parser.py).
- B_history_part2.md (7e09e5c, setup_view): says the help-card mode notes 'Fresh: New timestamped structure / Re-Run: Overwrite existing / Resume: Skip already processed' are 'unchanged from the old help'. They are shortened rewordings of c258047's help panel ('Creates a new timestamped folder structure', 'Overwrites files in existing structure', 'Skips clients already processed', c258047:gui/views/setup_view.py:176-179). The Tk v3 Setup tab had no such hand-typed lines; it showed the PROCESSING_MODES descriptions (47a1492:gui/tabs/setup_tab.py:212-223).
- B_history_part1.md (launch_extractor section): says B HEAD gui/pyqt_main_window.py:190 calls _populate_extraction_from_disk() at start-up and that this was 'added in 7e09e5c'. A pickaxe search shows _populate_extraction_from_disk was added in b0e7d0e; 7e09e5c added _populate_extraction_from_history at start-up, which b0e7d0e then deleted.
- B_history_part1.md / shared notes: the Tk folder creation in main.py at the split is 47a1492:main.py:29-43 (setup_environment; loop 35-43), not 'split lines 33-42'.
- B_registry.md B-R45, B-R130, B-R131: 'never used' is correct at the snapshot, but in the history PROCESSING_MODES 'name'/'description' (47a1492:gui/tabs/setup_tab.py:216,221), GUI_CONFIG (47a1492:gui/main_window.py:58-63, client_handler.py:58-59, validation_tab.py:24,28) and four MESSAGES entries (47a1492:gui/handlers/file_handler.py:123-139, processing_handler.py:30,61) were used by the Tk layer. They became dead only when the Tk files were deleted in a62d3cc. This is a nuance for history sections, not an error at the snapshot.
- Line-count note: B_files.txt and the registry use wc -l (7,805 Python lines at the snapshot); git's own count is 7,821, because several files lack a final newline. S12BF uses git counts and says so.

## From writer X345

### Facts other sections must reflect
- B has no tests at all. This was checked over the snapshot tree, all 17 commits and every ref (RAN). pytest appears only as the commented line requirements.txt:11@edac6a3, removed by c258047. Section 2 should tag no file [TEST]; Sections 9 and 12F should state this.
- HANDOVER.md and FILE_INDEX.md never existed in any commit of A or B (git log --all over combined; find over /home/user/AnnualStatement_oldv2). B also never had PROJECT_AUDIT_ANNUALSTATEMENT.md, docs/ or tests/.
- The original B checkout and the combined repo have a second branch, claude/pensive-feynman-0mz1aa, pointing at b0e7d0e. It adds no commits (git log --all = 17). Any statement that B has only one branch should mention it.
- gst_organizer.log exists only in edac6a3 and 6237597. It is the same git object in both: 803,193 bytes, 5,689 lines. cbfa4a5 (via GitHub) deleted it on one line of history and the merge 9c56a1c dropped it again. History only; never opened.
- 70 .pyc files are in the 7 commits edac6a3, cbfa4a5, 6237597, 9c56a1c, d674386, 5ea77c2, 1cb878b (the identical set in each) and were removed in 01862a0. Every one embeds an absolute Windows drive-letter path, none containing \Users\. 35 are cpython-312 and 35 cpython-313 (RAN, count only).
- The 9-column consolidator mismatch (B-G1) could not occur when the README was written. At d674386 there were 8 mappings, all with two-word labels (READ git show d674386:power_query_extractor/config/cell_mappings.py). 47a1492 introduced it by adding the one-word labels IGST/CGST/SGST. Section 12 should attribute it there. Re-confirmed by RAN consol_check.py: 22 columns, 9 mismatches.
- a62d3cc deleted gui/main_window.py while gui/__init__.py:7@a62d3cc still imported it, so the program could not start until c21f788, 28 minutes later.
- The on-screen note 'Names >10 chars will show warning…' (gui/views/setup_view.py:385) was copied back into the PyQt screen by 7e09e5c, after a62d3cc had deleted the features it describes. It was never true in B's PyQt code.
- README verdict counts and the 14 ranked discrepancies (D1-D14) are in X4.1.13 and X4.6. Except RM8 (47a1492) and RM108, every NO LONGER TRUE claim was caused by a B-only commit (mostly c258047, also a62d3cc and 7e09e5c).
- Commit-message claims that the diffs contradict (X4.5): c258047 says '.claude already in .gitignore' (it was added by that commit), '12-field AppSettings' (11), and 'New Files (20)/Modified (15)' (18/14). a62d3cc says 'All functionality has been reimplemented' (it was not). 6237597 says 'Corrected app.py reference in main.py' (main.py was not touched; main.py:71@47a1492 still lists app.py).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- B_docs.md §0.1 and §10 (runs 10, 11) speak of 40 .py files ('Module and class docstrings of all 40 .py files', '0 of 40 files fail ast.parse'). B has 38 Python files (RAN: git ls-tree and ast over a copy). All 38 parse with the 3.7 grammar and all 38 have a module docstring.
- B_docs.md §6 #10 quotes 'Manages light/dark theme toggling via QSS stylesheets' as gui/theme/theme_manager.py:2-3. That text is the class docstring at :409. The module docstring at :1-4 says 'Theme manager for PyQt5 - replaces the 212-line dark_mode_manager.py. Uses QSS stylesheets instead of recursive widget walking.'
- B_docs.md §3 (requirements.txt line 7-8) cites the openpyxl imports as core/excel_handler.py:31-33. They are at :30-32.
- shared_history_12_commits.md §9 Q13 says all commits other than 47a1492 are 'goyal-sidhant or later Sidhant Goyal'. In B's history, 16 commits have author name goyal-sidhant (cbfa4a5 committed through GitHub) and 47a1492 has 'Sidhant'. No B commit is authored 'Sidhant Goyal'; that name appears only in A's later commits (32 there) (RAN git log --format=%an).
- B_docs.md R93 (debug logging 'in main.py or launch_extractor.py') can be made exact: launch_extractor.py@47a1492 had no basicConfig either. The extractor's logging set-up was in power_query_extractor/extractor_main.py:12-19@47a1492, which that launcher imported. So the README claim was never true for launch_extractor.py, not just untrue now (READ git show 47a1492:launch_extractor.py).
- B_docs.md §7 calls the b0e7d0e message 'vague but not contradicted'. Reading the b0e7d0e diff confirms H2's more precise reading: the 'statistics tracking' is scan statistics added to gui/pyqt_main_window.py:244-316, and the 'logging' is Extraction_Log_*.txt (gui/controllers/extraction_controller.py:101, 109-121).

## From writer S3b

### Facts other sections must reflect
- Section 7: the tab-dot defects of B-F29 have no gap number. RAN: green after an empty scan and after an extraction that failed or was stopped; dots never reset; Step 2's dot can never turn green (no caller sets it to 'complete'); after an extraction error the status bar still says 'Extraction complete'. Only 7.8 row S4 mentions that dots are never reset (under B-G56). Suggest adding these to B-G71 or giving them a new gap.
- Section 7: the Step 1 and Step 4 options are not saved when Start (Step 3) or Start Extraction is pressed, only at a later settings save. A crash in a handler skips even the save on close. No gap covers this (READ gui/pyqt_main_window.py:345-371, 506-547, 601-608).
- Section 7, B-G40: RAN t34c. Closing during an extraction gave exit code 0 with no message, left an unrefreshed copy, and wrote no workbook, Extraction_Log or run history. This differs from the XL expectation of an abort.
- Section 7, B-G43: RAN t23. A time-zone-aware date in one figure made create_report raise TypeError and left a 4,579-byte file that is not a valid Excel file (openpyxl BadZipFile) under PQ_Extraction_Report_<minute>.xlsx.
- Section 7, B-G74: RAN t34b. A scan and a full processing run both completed while an extraction thread was running. The end of processing rebuilt the Step 4 tree onto the new run folder mid-extraction, while the extraction's report and log went into the old run folder.
- Section 7, B-G6: RAN t24b. The stand-in performs no refresh at all, yet B logged '✅ Refresh completed successfully' and 'Validated - Client Name: …' because Info!BB2 already held a value. The check proves only that BB2 was filled when the file was last saved. [GST RULE] ⚠ (cell BB2).
- Section 7, B-G61: RAN t26. With another program's 'Backup failed' dialog listed first, B clicked its OK button and left the Excel error dialog open. An Excel Yes/No question was answered '&Yes'. A 'Look up…' button was clicked because its caption contains 'ok'. The copy was saved in every case.
- Sections 7 and 9, B-G44: RAN t30. launch_extractor.py writes no log file and the INFO lines appear nowhere. main.py's gst_organizer.log lands in the start folder, not the program folder, and is appended at every start (6 lines after two starts). With a simulated cp1252 encoding the line containing ✓ was missing and stderr showed '--- Logging error ---'.
- Section 7, B-G71: RAN t24b. 536 of 541 Step 4 log lines were 'info', including '❌ Cell BB2 is empty' and '⚠️ Refresh completed with errors'. 'Extraction complete: 3/3 clients processed' was shown although one of the three failed. About 180 log lines per report.
- B-F20 / B-F25: a refresh that failed because of an error dialog keeps validation_message 'Validated - Client Name: …' in its stored status (RAN t26). It sits in run_history.json next to the failure.
- Section 2 dead code, confirmed by RAN grep: session_state.json is write-only (active_tab, window_size and recent_source_folders are never read); theme COLORS, the '#info' rules and ThemeManager.toggle; CollapsibleGroup (0 instances); CollapsibleCard.addLayout and its _bg/_border/_text_color attributes; power_query_extractor/extractor_main.py (cannot be imported); the 'client_key' of each Step 4 job is never read anywhere in power_query_extractor/ (grep); EXTRACTION_OPTIONS auto_refresh_if_missing and show_timestamp are never read.
- Section 12 (history): Skip refresh and its fixed '_Refreshed_' search date from edac6a3 (shared). The configurable suffix that broke the agreement dates from bd176bf (shared). The tree status that also hard-codes '_Refreshed' dates from b0e7d0e (B only).
- Section 10 (UI): at the default 1400x900 window, the Step 1 left column is 815 px high and SCAN FILES starts at y=989, so the main button is hidden until the user scrolls (RAN t33). 0 of 205 widgets have a tooltip (RAN).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- Registry B-F21 (Origin column) says 'suffix and skip bd176bf'. The code history shows that Skip refresh existed in edac6a3, shared - before the split on 30-10-2025: the Tk tick box (edac6a3:power_query_extractor/gui/extractor_window.py:188-198, writing EXTRACTION_OPTIONS at :234-238) and the fixed '<stem>_Refreshed_*.xlsx' search together with a fixed copy name, in edac6a3:power_query_extractor/core/report_processor.py. Only the configurable suffix (and the wait time) came in bd176bf, which broke the agreement between the copy name and the skip search (bd176bf:report_processor.py lines 28, 93, 103). READ via git show.
- Registry B-F26 and B-R103 say the dialog text 'goes to the program log only' / 'is written only to the program log'. The immediate line '⚠️ Error detected: <text>' does go to the program log only (report_processor.py:272, :295). But the same text also reaches the Step 4 screen log and Extraction_Log_*.txt at the end, as '⚠️ Refresh completed with errors: <text>' (report_processor.py:321-322). RAN scripts/t26_dialogs.py.
- XL notes §5.2 step 8 expect Qt to abort the program when the window is closed during an extraction (MEDIUM). RAN scripts/t34c_close.py (real event loop, same exit path as main.py) shows otherwise: the process ended with exit code 0 and no message while the worker was mid-refresh. Left behind: an unrefreshed '_Refreshed_' copy, and no consolidated workbook, extraction log or run_history.json.
- PQ notes say report_processor.py has '23 logger.warning/error/debug calls'. RAN grep: there are 23 logger.* calls in all, of which 18 are warning/error/debug and 5 are info (one of those is inside _log itself). So 22 calls bypass the screen log.
- My own file, B-F18 item 3: the reference to 'B-F24's second run' for a failed-refresh version turning 'Refreshed' was wrong. It now points to B-F20 Scenario C and B-F25 Run B, which is where that case was RAN (t25_full). B-F24's second run is the Stop run.

## From writer S45

### Facts other sections must reflect
- B-R83 precise condition: IMS Reco / Extra In Purchase / Extra In GSTR2B show Diff GSTR-3B J15/K15/L15 unless that value is None. '' gives 'Missing', 0 gives 0 (RAN t3). Sections 3 (B-F23), 6-7 should use this exact condition.
- B-R60: a client counts as 'Success' (and is listed as Success in Client Status) even when both reports failed. In practice only a folder-creation exception makes it 'Failed'. Relevant to B-F15 and the Section 7 gaps.
- B-R66: if all three Step 4 open attempts get 'Call was rejected by callee', the code still sends Alt+A,R,A with no workbook open. The refresh status then ends as a 'NoneType' save error (report_processor.py:195-205, 300-337).
- B-R99: Step 4 open attempt 2 closes every workbook with the same file name in that Excel without saving (report_processor.py:691-699). Because Dispatch may attach to the user's Excel, this could close one of the user's workbooks (MEDIUM). Relevant to the Section 8 Standard 12 comparison.
- Resource counts: Step 3 attaches to or starts one Excel per report (2 per client) and always quits it. Step 4 uses one visible refresh Excel plus one hidden reading Excel per report (up to 4 per client). At most one QThread per job kind (3 at once). The Step 4 disk scan runs on the window thread.
- Split-era values for the A/B comparison: the Tkinter extractor's wait box was 5-60, default 10 (47a1492:power_query_extractor/gui/extractor_window.py:181-186). B uses 5-120 (c258047). The template picker was '*.xltx *.xlsx' with the label '(.xlsx or .xltx)' (47a1492:gui/handlers/file_handler.py:31, 42; 47a1492:gui/tabs/setup_tab.py:121). B uses '*.xlsx *.xlsm' (c258047; label added in 7e09e5c). Report naming expressions were identical at the split (47a1492:gui/handlers/processing_handler.py:229-235, 292-293).
- The consolidated workbook's Summary sheet gets no column widths (data_consolidator.py:208-235). Only the output sheets are width-adjusted (cap 50).
- The processing summary is saved in, and Explorer opens, the run folder of the first client whose folders were made (gui/controllers/processing_controller.py:100-101, 201; gui/pyqt_main_window.py:404-410).
- The table header colour is #366092 in the light theme but #2d5280 in the dark theme (gui/theme/theme_manager.py:155, 340). The processing summary uses 366092 and the consolidated workbook 0078D4.

### UNREGISTERED items
- UNREGISTERED: Extraction output folder falls back to the target folder when no run folder is given (gui/controllers/extraction_controller.py:28-29; DEAD in practice because gui/pyqt_main_window.py:512-515 refuses to start without one; B only, b0e7d0e)
- UNREGISTERED: Every organiser path is built by safe_path_join, which silently drops empty name parts and raises ValueError('No valid path components provided') when none are left (utils/helpers.py:49-64; used at core/file_organizer.py:94, 102, 132, 153, 176, 255, 314; shared, edac6a3)
- UNREGISTERED: FileOrganizer.get_summary (dead) would count folders_created as distinct folders, unlike the organisation report's running total (core/file_organizer.py:392-408 vs 345-351; shared, edac6a3)

### Corrections to the working notes
- PQ notes (B_code_pq_models_startup.md mapping table rows 14-22), LT §4.6 and OF §1.4C say the three one-word output sheets show their own cells when Diff GSTR-3B J15/K15/L15 are 'empty'. Precisely, the fallback happens only when the value read is None (an empty cell or a missing sheet). An empty text '' gives 'Missing', not the configured cell, and 0 gives 0 (power_query_extractor/core/data_consolidator.py:130-142; RAN t3_lookup.py).
- Registry B-R60 says only an exception 'during folder creation or copying' makes a client Failed. Copy errors are caught per file in core/file_organizer.py:232-283 and never fail the client. In practice only folder creation (create_client_structure re-raises, core/file_organizer.py:140-144) or an unexpected bookkeeping error does. Organisation-report errors are only logged (core/file_organizer.py:300-362).
- Registry B-R137 says every refresh/extraction message arrives as 'info'. That holds for the refresh/reading code's messages (power_query_extractor/core/report_processor.py:30-37). The extraction worker's own lines carry success/error/warning levels (gui/controllers/extraction_controller.py:62, 84, 93, 95, 98).
- OF §5.2 approximates the refresh wait as 'about 2×wait + 20 s per report'. The exact sum of the fixed sleeps is 19 + w + max(10, w) seconds (34 s at w=5, 39 s at 10, 259 s at 120), plus Excel start-up (report_processor.py:139-141, 207-319; arithmetic RAN in t3).

## From writer S89

### Facts other sections must reflect
- Python version (Sections 1, 11, 12F): the 'Python 3.7 or higher' in requirements.txt:2 and README.md:51 cannot be met. PyQt5 5.15.11 needs >=3.8, and the newest pyqt5_sip 12.20.0 needs >=3.10. Windows wheels for Python 3.14 exist for every package B needs (RAN pip download, 23-10-2026). B's syntax compiles on 3.11-3.14.
- Qt on Windows (Sections 10, 12F): the newest PyQt5-Qt5 build for Windows is Qt 5.15.2, while all screenshots and look-checks on this Linux machine ran Qt 5.15.19, so Windows rendering may differ slightly.
- Warnings switch (Sections 2, 7 B-G31): core/__init__.py imports core.excel_handler, so importing ANY core module (even just core.file_parser for a scan) runs warnings.filterwarnings('ignore') and loads openpyxl (RAN).
- Section 11, Item 4 constraint: the keystroke refresh predates git. The 25-06-2025 compiled report_processor already had SendKeys and no RefreshAll (RAN, byte counts), and no B commit ever contained RefreshAll, DispatchEx or CalculateUntilAsyncQueriesDone (RAN, git log -S). Any change must guarantee the refresh has completed before saving (Standard 12 lines 860-867), so Section 11 must not recommend a plain RefreshAll followed by Save.
- Sections 4 and 11: B builds every report and refreshed copy directly at its final path (no temporary-file stage; gst_excel_temp and temp_gst_ are pre-git remnants), and B has no folder-listing cache. The Step 4 list source changed three times (session results in c258047, run history in 7e09e5c, disk walk in b0e7d0e), and the disk walk was kept deliberately after the history-based list gave phantom and stale rows.
- Standard 13 runtime (Sections 9 and 11, Foundation): pywin32 312 relies on its pywin32.pth file (adds win32, win32\lib, pythonwin and runs import pywin32_bootstrap). Whether an embeddable _runtime honours it is unverified; check with `_runtime\python\python.exe -s -c "import pythoncom, win32com.client, win32api, win32con, win32gui"`.
- Standard 13 launcher effect (Sections 7 and 11): under start /D "%~dp0" the start folder is the shared app folder, so gst_organizer.log (main.py:20) would be shared by all users. A simulated pythonw start (stdout and stderr None) gives exit code 1, nothing visible and one log line (RAN).
- Section 12 (API history): the window's direct access to Setup and Extract widgets (gui/pyqt_main_window.py:185-207, 544-546) already existed at c258047 (RAN git show). The scan-button switching came in 7e09e5c. setup_view.py had 0 setStyleSheet calls at c258047 and 36 at 7e09e5c; there are 76 inline style sheets in views and widgets at the snapshot.
- Every gui.* module (views, widgets, theme, controllers) fails to import without pywin32 in a fresh process, because gui/__init__.py:7 eagerly loads the whole window (RAN). The business core, models (except signals.py), cell_mappings and data_consolidator import without PyQt5 and without pywin32 (RAN).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- Qt version: B_code_pq_models_startup.md says PyQt5-Qt5 / Qt 5.15.19, while B_history_part1.md and Section 12F.2 say 'PyQt 5.15.11 / Qt 5.15.14'. Both are partly right. RAN (me): PyQt5 5.15.11 was built against Qt 5.15.14 (QT_VERSION_STR) and runs Qt 5.15.19 (qVersion()) in venv-B. On Windows, pip's newest PyQt5-Qt5 build is 5.15.2.
- B_docs.md R27 says 'Only Python 3.14 was available'. /usr/bin/python3.11, python3.12 and python3.13 also exist on this machine. All 38 B files compile with each of them and with 3.14.6: 0 syntax errors, 0 syntax warnings (RAN (me) compile_all.py, in memory).
- Section 12F.3 says the 'Python 3.7 or higher' line matches the dataclasses requirement. That is true only for the syntax. The wheel metadata (RAN (me)) shows PyQt5 5.15.11 needs Python >=3.8, pywin32 312 >=3.9 and pyqt5_sip 12.20.0 >=3.10, so 3.7 cannot be met by the current packages.
- B_dead_code_patterns.md P-05 counts '8 named constants' never read. Precisely, 7 constants (GUI_CONFIG, MESSAGES, EXCEL_EXTENSIONS, TEMPLATE_EXTENSIONS, MAX_FILE_SIZE_MB, EXCEL_SIGNATURES, EXCEL_SAFETY) plus the FOLDER_STRUCTURE['level2'] key are never read. This is a counting nuance only.

## From writer S2

### Facts other sections must reflect
- The openpyxl report path (core/excel_handler.py:88-93 and :219-262) is unreachable through the app on any real PC, not just 'in practice'. core/excel_handler.py sets USE_WIN32COM from the same two imports that report_processor.py:11-12 needs unconditionally, and pywin32 does not exist off Windows. Sections 3 (B-F13, B-F14) and 9 should say so.
- Deleting ~/.gst_organizer/.migrated while the old v3 file ~/.gst_organizer_cache.json still exists re-imports the old v3 settings over the current ones at the next start (RAN: source folder and pq_wait_time were reverted). Relevant to Sections 3 (B-F2), 7 and 9.
- main.py creates gst_organizer.log in the current folder even when start-up fails, because logging is configured before the failing import (RAN, 83 bytes). launch_extractor.py creates no log file, and its application name is the file name 'launch_extractor.py' (RAN). Relevant to Sections 3 (B-F30, B-F32) and 9.
- AppSignals has 18 signals, not 19. B has 34 live signal declarations; each is emitted and connected. No QTimer, threading, importlib or eval/exec exists in B (RAN grep).
- ReportProcessor._log(level='info'): none of its 96 calls passes a level, so every refresh and extraction message, errors included, reaches the Step 4 log as 'info' (RAN count). This is consistent with B-R137.
- Commit-message claims that disagree with the code, for Section 12. c258047 says 'New Files (20)' and 'Modified Files (15)'; it actually added 18 files, modified 14 and deleted 1 (RAN git show --name-status). c258047 says '.claude ... already in .gitignore'; the same commit added .claude to .gitignore:81-82. 7e09e5c says ValidationView got 'collapsible sections for instructions and scan summary'; validation_view.py has no collapse function (READ grep).
- Real line counts: the 38 .py files have 7,821 lines; 18 of the 41 files lack a final line-break. Sections that quote 7,805 are quoting wc -l.
- B has no tests and no file carries the [TEST] tag (consistent with flag item 6).
- ScanResult.scanned_files and ScanResult.variations are write-only after a scan; only statistics and client_data are read. This is the in-memory reason the problem-file list never reaches the user (B-F9).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- B_dead_code_patterns.md D-01 and §2.5 say AppSignals holds 19 signals; models/signals.py declares 18 (:13, :16-19, :22-27, :30-34, :37-38). Checked by READ and a RAN count of the pyqtSignal( lines.
- B_dead_code_patterns.md §2.2 says ScanResult is 'used for its four fields'. Only statistics (gui/pyqt_main_window.py:247) and client_data (:274, :335, :368) are ever read. scanned_files and variations are written at gui/controllers/scan_controller.py:33, 35 and read only by the dead properties and clear().
- B_dead_code_patterns.md D-03 gives power_query_extractor/gui/__init__.py as '0 lines', and B_files.txt lists the config/ and core/ __init__.py files of power_query_extractor as 0 lines too. Each holds 1 line (a docstring with no final line-break). wc -l counts line-break characters, not lines.
- B_files.txt, B_dead_code_patterns.md §0.1 and the registry scope notes give 7,805 lines for the 38 .py files. That is the wc -l count; the real number is 7,821, because 16 of the 38 .py files (18 of all 41 files) have no final line-break (RAN).
- shared_history_12_commits.md §1.14 lists ReportProcessor.cleanup() as dead at edac6a3. It was called, although it does nothing, at power_query_extractor/gui/extractor_window.py:585@edac6a3 and :832@47a1492. It lost its caller only when a62d3cc deleted that file (RAN git grep).
- B_code_core_utils.md says the openpyxl report fallback is 'reachable only in the odd case where pythoncom imports but win32com.client does not'. It is not reachable even then: power_query_extractor/core/report_processor.py:11-12 imports both pythoncom and win32com.client unconditionally, so in that case the window cannot start either. Through the app the fallback is unreachable on every real PC (READ).

## From writer S10

### Facts other sections must reflect
- Section 7 / registry B-F29, B-R140: tab dots mean 'this step ran', not 'it succeeded'. Step 2's dot can never turn green: the only _update_tab_state calls are at gui/pyqt_main_window.py:305, 306, 391, 392, 551 (READ + grep). There is no gap of its own; the closest is B-G71.
- B-R139 / B-G71: after an 'Extraction Error' the status bar keeps its previous text, e.g. 'Extraction complete' from an earlier run, and the Step 4 dot stays green (RAN e7 by calling the window's own handlers; consistent with S3b t29). If the error happens before extraction_controller.py:86, that run gets no consolidated workbook, no Extraction_Log and no rebuilt list (READ).
- B-G35 / B-G36 (RAN e7): with ~/.gst_organizer/session_state.json holding a JSON list, the window opens normally. View > Dark Mode, File > Save Settings, SCAN FILES and closing the window each abort the program (SIGABRT, exit 134 in a console) with AttributeError at models/cache_manager.py:47, and settings.json is not written.
- B-G62 / B-F19: when AppActivate fails three times, B still sends Alt+A, R, A. This is logged only to the program log ('Could not activate Excel window, continuing anyway...', power_query_extractor/core/report_processor.py:226, 230), while the screen log shows 'Starting refresh sequence...' and '✓ A key sent - Refresh All command executed' as usual (READ :234-249).
- B has no QComboBox (dropdown), no context menus and no double-click handlers anywhere; the only lists are the Step 2 QTreeView (branch arrows off, gui/views/validation_view.py:321-323) and the Step 4 QTreeView. This matters for the A/B comparison of screens.
- Processing-controller line 'Summary error: <e>' (gui/controllers/processing_controller.py:193) is unreachable in practice, because ExcelHandler.create_summary_report catches every exception and returns False (core/excel_handler.py:344-346). A failed summary is therefore silent on screen (B-R94, B-G71).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- Qt version: B_history_part1.md:70, B_history_part2.md:75, B_code_mainwindow_controllers_theme.md:12 and B_code_views.md:13 give the run environment as 'PyQt 5.15.11 / Qt 5.15.14'. 5.15.14 is the Qt version PyQt5 was built against (QT_VERSION_STR); the libraries actually running are Qt 5.15.19 (qVersion(), RAN 23-10-2026). B_code_pq_models_startup.md has it right. This section's own earlier text repeated the mistake; now corrected.
- B_history_part2.md:316 calls the Setup help card's mode lines ('Fresh: New timestamped structure', 'Re-Run: Overwrite existing', 'Resume: Skip already processed') 'unchanged from the old help'. They are shortened rewordings of c258047's radio texts ('Re-Run - Overwrite existing files', 'Resume - Skip already processed', c258047:gui/views/setup_view.py:102-104) and help-panel texts ('Overwrites files in existing structure', 'Skips clients already processed', :176-179). READ via git show.
- B_excel_pq_safety.md:279 (step 8) expects Qt to abort the program when the window is closed during an extraction (MEDIUM). RAN by the Section 3b writer (t34c): the process ended normally with exit code 0 and no message.
- Own earlier S10 text: 'The reason a report failed is never shown on screen' was too absolute. A failed BB2 check ('❌ Cell BB2 is empty (None)', '⚠️ Refresh validation failed: …', report_processor.py:310-313, 611-629) and an Excel pop-up's text ('⚠️ Refresh completed with errors: <text>', :321-322) do reach the Step 4 screen log, as info-level blue lines. Every other cause (Excel start/open failure, missing sheet, unreadable cell, report not found, not ticked) goes only to the program log (:155, 333, 340, 504, 522, 543, 548, 567; :50-68, 79-87).
- Own earlier S10 text: the cross-reference 'gui/pyqt_main_window.py:228 → :590-594' pointed at the wrong file. The second part is gui/views/setup_view.py:590-594 (collapse_info_cards).

## From writer S12own

### Facts other sections must reflect
- Repository origin: GitHub's creation dates show B (AnnualStatement_oldv2) was created on 31-08-2025 at 22:08:53 IST, 3 minutes before edac6a3. A (AnnualStatement_oldv1) was created on 07-02-2026 at 11:21:52 IST, 6 minutes before B's PyQt5 rewrite c258047. Yet the old name goyal-sidhant/AnnualStatement now resolves to A (bf2f1dd). This contradicts project A's audit inference that 'A is probably the original repository, renamed (MEDIUM)'; the likelier reading is that B is the original (MEDIUM). Sections 1 and 9 of both audits and the comparison should reflect this.
- B was last pushed to GitHub on 09-02-2026 at 08:42:37 IST, 5 seconds after b0e7d0e. GitHub holds only main and HEAD at b0e7d0e (ls-remote, 23-10-2026). B's own development lasted about 45 hours, from 07-02-2026 11:27 to 09-02-2026 08:42.
- The second branch name claude/pensive-feynman-0mz1aa in B's checkout was created by 'Claude' on 21-10-2026 at 03:04:18 UTC (.git/logs/HEAD). It is not the author's branch and holds no extra commits.
- c258047 changed the openpyxl report path (core/excel_handler.py:251-252). A template with a merged Links cell used to be skipped silently and the report counted as made. Now the report is reported as failed and the unfilled template copy stays under the report's name (RAN me). This path can only run without pywin32, which B cannot start without, so it is dead in practice (B-R97). Sections 6 and 7 may cite it.
- At c258047 the tkinter package is still loaded at start-up; from a62d3cc on it is not (RAN me). This matters for Standard 13's embeddable Python, which lacks tkinter: B at the snapshot does not need it.
- README.md in B, unchanged since the shared d674386, disagrees with the snapshot on: the version (v3.0), the folder tree (gui/handlers, tabs, utils), the requirements block, the template types (.xltx), the stand-alone extractor and its 'Refresh & Extract' button, the log files (pq_extractor.log), and the mode texts against the Step 1 help card. The table is in S12own.md and Sections 2, 9 and 10 should agree with it.
- Project A, by contrast, still drops tiny or invalid Excel files silently at discovery (A utils/helpers.py:310-361), whereas B since 7e09e5c lists them as problem files (but shows only their count). This is for the comparison.

### UNREGISTERED items
- (none)

### Corrections to the working notes
- B_history_part1.md says models/signals.py (AppSignals) has 20 signals; it has 18 (models/signals.py:13-38; RAN me: 18 pyqtSignal( lines). S12BF.md already carries the same correction.
- B_history_part2.md says the 7e09e5c help-card mode notes ('Fresh: New timestamped structure', 'Re-Run: Overwrite existing', 'Resume: Skip already processed', gui/views/setup_view.py:510-512) were 'unchanged from the old help'. They are shortened from the c258047 wording ('Re-Run: Overwrites files in existing structure', 'Resume: Skips clients already processed', gui/views/setup_view.py:177-179@c258047). They also do not come from the v3 Tkinter card, which had no mode notes (gui/tabs/setup_tab.py:334-377@47a1492).
- B_history_part2.md says output_folder is set at gui/controllers/processing_controller.py:158. It is line 157 at c258047 and line 159 at the snapshot.
- B_history_part1.md's example 'max 8 gave ClientAl for both' depends on its test names. With my made-up names (ClientA Private Limited / ClientA Traders), maximum 8 gave 'ClientA ' with a trailing space for both (RAN me). The finding, two clients folded into one folder, stands.
- In the c258047 diff, report_processor.py had 12 bare except: clauses converted, not 11 (RAN me: 12 at 47a1492). The 26 total in the commit message is correct (12 + 7 + 2 + 2 + 2 + 1).
- B_history_part1.md took the GitHub creation dates (B 31-08-2025, A 07-02-2026) 'from the task brief'. I confirmed them with the GitHub API (RAN me, 23-10-2026): B created 2025-08-31 16:38:53 UTC, A created 2026-02-07 05:51:52 UTC, which is 6 minutes before c258047.

## From writer X12

### Facts other sections must reflect
- The openpyxl fallback (core/excel_handler.py:219-262) strips Power Query and returns True (RAN F1). It is unreachable through the window only because report_processor.py:11-15 imports pywin32 without a guard and the window loads it at start-up (c258047, B only). Any merge that makes the extractor import optional re-opens a silent query-dropping path with '✅ Success' on screen.
- At the split point (47a1492) the Tkinter organiser did not import the extractor, so a Windows PC without pywin32 took the openpyxl fallback and lost Power Query. Sections on history (12C/12D) and dependencies (9) should say so.
- Standard 12 verdicts per session (X1.3.6): S1 Step 3 meets 4, misses 8 (incl. R2a DispatchEx, R2c DisplayAlerts True at Quit, R2d takes over user's Excel, R3 Quit skipped when clean-up raises, R4a/R4b, R7b/R7c). S2 Step 4 refresh meets 6, misses 11 (incl. R2b Visible True, R5a-R5d no BackgroundQuery/RefreshAll/CalculateUntilAsyncQueriesDone/Refreshing). S3 meets 6, misses 6. Fallback misses R1a-R1c.
- Step 4 sends the Alt+A, R, A keystrokes even when every open attempt failed and wb is None (RAN fake COM, open_rejected) - B-G62 should cite this run.
- The error-dialog closer picks the first matching dialog of any program. RAN fake COM: it recorded an unrelated program's dialog text as the refresh error, clicked that program's 'Close' button, and left Excel's own error box open (B-G61).
- When Excel is missing or the refresh fails before Excel starts, an unrefreshed '<stem>_Refreshed_<DDMMYY_HHMM>.xlsx' copy is still left (report_processor.py:120 runs before Dispatch). The Step 4 tree then shows 'Refreshed', and Skip refresh re-uses it with success True and ✓ figures (RAN S4 no_excel and skip_after_failed; B-G7).
- Any failed Step 3 fill (open error, long path, Excel dying) leaves the unfilled template copy under the final report name (RAN S3 long_path and excel_dies, 8,406-byte file). Excel dying also skips Quit and CoUninitialize (RAN). Affects B-F13/B-F14 status and B-G5/B-G38.
- c258047 (B only) narrowed the fallback's bare except to (KeyError, ValueError): a merged Links cell now makes the fallback report fail (AttributeError, RAN F2), where at the split it was silently skipped.
- RAN E2E confirms with the real window: IMS Reco / Extra In Purchase / Extra In GSTR2B columns show Diff GSTR-3B J15/K15/L15 values (B-G1). A failed BB2 check turns all read figures into 'Missing' (B-G10). Extraction history is stored under 'client folder / version folder' keys that nothing reads, with pq_refreshed True even for the failed client (B-G25). The summary 'Variations' sheet is empty although the scan had 3 problem files (B-G20). 'Extraction complete: 2/2 clients processed' is shown although one client failed both reports.
- The on-screen Step 4 log dropped the BB2 value '<Client A>' because the log inserts HTML (B-R137); the Extraction_Log file kept it. Any client name or value containing '<...>' disappears on screen.
- A corrupt settings.json gives defaults, and the window close then overwrites the user's file with defaults (RAN C4). 'pq_wait_time': '10' (text) or a JSON list stops the window being built (RAN C2, C3) - relevant to Sections 6/7 (B-G36) and Standard 13 (no error box).
- gst_organizer.log follows the current folder (RAN M1/E2E). Under the Standard 13 launcher this means one shared log in the shared app folder, written by several users and holding client names, Links paths and extracted figures. launch_extractor.py writes no log at all (RAN M2).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- B_excel_pq_safety.md §5.2 step 6 says Step 4's open routine makes 'up to 4 opens per try, 3 tries = 12 opens'. RAN (fake COM, S4 open_rejected): with a path of 218 characters or less and no differing 8.3 name, each try makes 3 opens (attempt 1, attempt 2, the keyword open at report_processor.py:725-731), so 9 for the refresh plus 3 for the extraction. 4 per try happens only when the 8.3 path differs.
- B_excel_pq_safety.md finding 13 and B_outside_files_state.md §4.2 present 'on this Linux disk the refreshed copy was picked' as the Linux result. RAN here: it depends on folder-listing order even on the same Linux disk - in scenario copy_of_copy the original came first; in custom_suffix_skip the earlier _R… copy came first and a copy of the copy was refreshed.
- B_excel_pq_safety.md §5.2 estimates 'about 40 s' per report. RAN with recorded sleeps: exactly 39 s of fixed waits per report at the default wait of 10 s (19 + w + max(10, w)), plus two Excel start-ups. With an error dialog found at the first check: 34.5 s.
- B_excel_pq_safety.md §3 table marks S1 rule R4 (Quit in finally) 'PARTLY MET'. Under the task's three-value scale it is NOT MET: Quit() is skipped whenever core/excel_handler.py:204-205 raise - the after-an-error case the rule names (RAN fake COM, S3 excel_dies).
- B_outside_files_state.md §2.5 says STATE_CODE_MAPPING has 38 entries. utils/constants.py:234-275 has 39 names and 36 distinct codes (already corrected in the registry; X2.2.4 uses 39/36).
- B_outside_files_state.md §0 says '41 .py files'. B has 38 Python files and 41 files in all (already corrected in the registry).
- B_outside_files_state.md S6 and B_dead_code_patterns.md T7 describe the cp1252 log-line loss. Re-run here (scripts/log_encoding.py): a FileHandler with cp1252 prints '--- Logging error ---' and drops the line with ✓, keeping the plain lines around it - confirmed.
- Registry / OF note on migration: a saved client_name_max_length of 0 is copied as 0 into settings.json. RAN C1 additionally shows that the spin box (range 15-100) then displays 15, and the window close saves 15 - so the stored value becomes 15 after the first session.
