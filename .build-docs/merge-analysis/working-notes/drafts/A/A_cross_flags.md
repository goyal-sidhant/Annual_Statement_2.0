# Cross-section flags from the section writers (project A)

Each writer listed facts that OTHER sections must reflect, items the registry lacks (UNREGISTERED), and corrections to the working notes. The checkers have since corrected some of these in place; confirm each in the code before acting.

## From writer S3a

### Facts other sections must reflect
- Section 7 must include the 25 findings listed at the end of S3a.md (e.g. silent death without a console; install offer installs all of requirements.txt incl. pytest; Dark Mode tick off-window and status bar hidden at 1000x700; close during Step 3 kills the worker, close during Step 4 gives no warning; typed paths/mode not saved and an empty length box blocks every save and ends every run; checklist remembered-existence answers go stale and a file passes as a folder; 'Full pipeline' does not refresh and 'organise ... no Excel automation' is untrue on Windows; length setting shortens only Level 2; silent scan drops, hyphen mis-filing, .xlsm never matches; Re-scan mid-run aborts the run; number-looking client/state text breaks keys; Resume does not resume; same-minute overwrite; clients sharing one folder; Links cells for missing types keep template text; missing Links sheet reported as success; user's Excel taken over, hidden and quit; openpyxl fallback drops Power Query; 14 [DEBUG] lines).
- A-F15 -> A-F16/A-F20/A-F21: a client without Sales, Sales Reco or Annual Report files gets a Sales report that is an unchanged template copy (no Links cell written), logged '✅ Success', counted in 'Sales Reports Created', and shown in Step 4 with has_sales=True and the Sales tick pre-ticked, so Step 4 would refresh it (RAN).
- A-F16: the summary workbook's 'Include Client Name in Folders' shows only the global tick (processing_handler.py:225) - it said 'No' while a per-client flag put the name into one client's folders (RAN). The organisation text report's 'Successfully Copied'/'Folders Created' are run-wide running totals (RAN again: 2 files -> 'Successfully Copied: 9, Folders Created: 11').
- A-F17: after a fatal error at the very start of a run (e.g. empty length box), the progress area keeps the previous run's '📊 Progress: 100.0%' and 'Processing complete' text; the 'Processing Complete' and 'Processing Error' message boxes are raised from the worker thread (RAN: thread names recorded).
- A-F2/A-F7/A-F17 (UI, Section 10): status-bar messages are overwritten at once by the tab-change text - 'Welcome to GST File Organizer v3.0' (run()) and 'Found N clients' (scan) are never visible (RAN).
- A-F27: typed paths and the processing-mode radio are not saved; a non-numeric 'Max client folder name length' box makes every save_cache raise TclError (RAN).
- A-F10/A-F27: per-client folder-name flags keep keys of clients from earlier scans (client_handler.py:49 sets but never clears); turning the global tick OFF leaves all per-client flags ON (RAN).
- A-F14/Section 9 (Standard 12): fake-COM run confirms the COM call order GetActiveObject -> Visible/DisplayAlerts/ScreenUpdating False -> Open staging copy -> write non-empty cells -> Save -> Close -> ScreenUpdating True -> DisplayAlerts True -> Quit (always) -> CoUninitialize; with no Links sheet it calls Sheets.Add and returns True; an exception in the finally block skips Quit/CoUninitialize and leaves the staging file while the report already exists at the destination.
- Section 8/9 (threading): Step 3's worker can only talk to Tk while the main loop runs - when the harness pumped root.update() instead of mainloop(), every Tk call from the worker failed with 'RuntimeError: main thread is not in main loop' (the same error GUI exp4 saw after closing the window).
- A-F11: after a Step 3 run the client records hold the cleaned copy names (in-place rename, core/file_organizer.py:233-246), so a later Dry Run lists names not present in the source folder (inference from RAN data, HIGH).
- Section 12 (history): the 14 user-visible '[DEBUG]' lines already existed at 47a1492 (14 at 47a1492 and 14 at bf2f1dd, RAN grep); 064c634's message says they 'are intentionally left'; the whole-file install is deliberate per a5ad487's message; 7347383's message says 'a path-length guard remains the durable fix'.
- Section 9 (Standard 13 facts): A requires a console (main.py:17 raises with sys.stdout None, no log created) and offers to pip-install into the running Python; the log file lands in the current folder; logs/ and temp/ are created in the program folder and never used (RAN).

### UNREGISTERED items
- UNREGISTERED: Step 3 copy log line names the document type as the destination folder ('📁 <original file> → <type> folder', e.g. '→ GSTR-2B Reco folder' for a file that went into ITC, '→ Annual Report folder' for one in the Version folder) - gui/handlers/processing_handler.py:284 (UI text rule not covered by A-R140)

### Corrections to the working notes
- Do the real ITC and Sales templates hold any text (sample paths or file names) in their Links cells B2-B11? Today, cells for file types a client does not have keep whatever the template contains (RAN).
- Should a client that has no Sales, Sales Reco or Annual Report file get a Sales report at all? Today it gets an unfilled copy of the template, reported as success and pre-ticked in Step 4. [GST RULE area: which returns feed which report - your decision]
- When a client has several files of one type, which one should feed the report? Today the first in alphabetical path order is linked and nothing tells the user. [GST RULE - your decision]
- Should 'Resume' continue inside the interrupted run's own Version folder and skip files already copied? Today it makes a new Version folder and copies everything again unless started in the same minute.
- Should 'Full pipeline' start the Power Query refresh automatically, or should its description and help text be changed? (Today it only switches to Step 4 and scans.)
- Is a console window wanted at start-up? Should the install offer install the whole requirements.txt (including pytest), or only the missing package?
- Should the 'Max client folder name length' setting also shorten the Level-4 folders, copied file names and report names? Should typed values below 15 be refused?
- When the global 'Include client name' tick is turned OFF, should the per-client ticks be cleared too?
- Should Step 2 list the unrecognised files, and the files the scan drops silently (small files, sub-folders, .xlsb/.xltx/.csv/.xlsm)?
- Do source file names ever have an all-digit client name, a state written as digits, or a hyphen inside the client name? Each behaves badly today.
- Do you or your staff keep your own Excel open while Step 3 runs? The code attaches to that Excel, hides it and then quits it.
- Should Scan, Re-scan, Browse and the Step 2 buttons be locked while Step 3 runs?
- Keep or remove the 14 '[DEBUG]' lines in the Step 3 log?

## From writer S67

### Facts other sections must reflect
- A-F24 / consolidated workbook: 9 columns (IGST/CGST/SGST on 'IMS Reco', 'Extra In Purchase', 'Extra In GSTR2B') hold Diff GSTR-3B!J15/K15/L15 values instead of their configured cells (data_consolidator.py:141-150). RAN me r2: 13 columns OK, 9 mismatched. The lookup is in shared commits, so project B probably has the same defect. [GST RULE] for which cell is right.
- A-F13 'Resume' re-copies everything unless the run happens in the same minute, so it never resumes across minutes (RAN me r3). The README:145 description is not true.
- A-F14/A-F15: a template without a Links sheet produces useless reports marked as success (COM path adds the sheet; openpyxl path writes nothing).
- A failed openpyxl save leaves a broken file under the final name. RAN me r2b: a 3010-byte PQ_Extraction_Report that cannot be opened. There are no temporary-name saves anywhere.
- A-F21/A-F22: a failed refresh leaves a _Refreshed_ copy carrying the original's date. The next Skip-Refresh run reuses it with status success (RAN me r4).
- A rescan during Step 3 stops the WHOLE run ('Fatal error'). It is not counted as one failed client (processing_handler.py:232 vs :243, RAN me r5).
- A client or state written only in digits ('0042', '27') comes back from the Treeview as an int. This breaks Dry Run, Start and Export, and one such client stops the whole run (RAN me r5).
- Step 4 run-folder scan: a newer FILE named 'Annual Statement-…' is chosen as level1_folder, and the scan then fails with 'Not a directory'. Pointing at a run folder itself finds nothing (RAN me r7; extractor_window.py:622-668).
- sanitize_filename treats a trailing dot differently on Python 3.14 and on 3.11/3.13. The same client gets different file names on PCs with different Python versions (RAN me).
- main.py:17 crashes before logging or any window when there is no console (pythonw, double-click, shortcut), so the program dies silently. This matters for the author's Standard 13 launch style.
- core/excel_handler.py:56/242: excel_app is never set, so Step 3 always calls Quit() on whatever Excel it attached to via GetActiveObject, including the user's own open Excel.
- Step 4 Start can be pressed again inside the main window and start a second worker. There is no Stop button in the tab (A-G65).
- After a 'full' Step 3 run the hand-over overwrites Step 4's folder unconditionally (processing_handler.py:166-168), bypassing extract_tab.py:121-131 (A-G65).
- Defensive-coding rating: MODERATE overall. Step 3 Excel reports, Step 4 refresh and the consolidated workbook are FRAGILE. Handler counts: 147 handlers, 68 leave no trace, 124 never show anything to the user, 23 bare excepts.
- Logs are written in the START folder (main.py:24) with the PC's default code page. With cp1252 the emoji success/failure lines are lost (RAN me r6). Logs contain client names and figures (confidentiality).
- There is no GSTIN/PAN logic anywhere, no period/financial-year check, and no free-space check (READ, grep).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- Q1 (A-G1): On the 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' sheets of the consolidated workbook, are the IGST/CGST/SGST columns meant to show the configured cells (IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1, ExtraInGSTR2B!S1/T1/U1) or the Diff GSTR-3B figures the code actually writes? [GST RULE]
- Q2 (A-G5): What do the real ITC and Sales templates hold in the Links cells for file types a client does not have? The templates are not in the repository.
- Q3 (A-G6): When a client has two or more files of one type, which one should feed the report, or should the run stop and ask? [GST RULE]
- Q4 (A-G10): What was 'Resume' meant to do, and has anyone relied on it?
- Q5 (A-G22, A-G23, A-G56): How and from which folder do staff start the program: a console, double-click, shortcut or pythonw; a local or a shared network folder?
- Q6 (A-G54, A-G55): Do staff usually have other Excel workbooks open, or use the PC for other work, while Step 3 or Step 4 runs?
- Q7 (A-G33, A-G34): Do client or state names in real file names ever contain a hyphen? How are such files named today?
- Q8 (A-G36, A-G77): Which state spellings actually appear in file names (full names, short forms, codes, '&' versus 'and')?
- Q9 (A-G39): Is the tool expected to notice files from different periods or financial years in one source folder? [GST RULE]
- Q10 (A-G3, A-G44, A-G79): Which Excel version and display language are on office PCs? Is the Data ribbon in English everywhere (the refresh presses Alt+A, R, A)? Are there add-ins or Protected View dialogs?
- Q11 (A-G46): Do the templates' Power Query sources use drive letters or \\server\share paths, and do all PCs map the same letters?
- Q12 (A-G15, A-G16): Which 'Refreshed File Suffix' values are used in practice?
- Q13 (A-G2, A-G3): What value in Info!BB2 means 'refresh worked', and can it be present when part of the refresh failed? [GST RULE]
- Q14 (A-G7, A-G31, A-G57): Is the program ever run on macOS or Linux, or only on Windows?
- Q15 (A-G32, A-G65): Is the standalone extractor (power_query_extractor/extractor_main.py) still used?
- Q16 (A-G71, A-G9): Can two people run Step 3 or Step 4 at the same time against the same target or run folder?
- Q17 (A-G48, A-G49, A-G50): How many files does a typical source folder hold, and is it on a network share?
- Q18 (A-G21, A-G41): Is it intended that 'Max client folder name length' shortens only the client folder and not the ITC (<client>) folders, file names and report names?

## From writer S12BF

### Facts other sections must reflect
- The VERIFICATION RULE passes: no file that was never changed imports a name added later (RAN, verify_imports.py). Other sections still should not equate 'file never changed' with 'behaviour never changed': data_consolidator.py has never changed, yet it began putting another cell's value into 9 columns when 47a1492 changed its configuration (A-R88, gap A-G1). Effect: edac6a3 config 0 of 8 columns; 47a1492 and bf2f1dd 9 of 22 (RAN).
- The consolidator cross-wiring comes from a shared commit (47a1492, shared - before the split on 30-10-2025), so the comparison with project B must treat it as inherited by both projects.
- 7347383 did not change the positions of the Links cells, but it did change the folder text written into ITC B2/B7/B10 and Sales B2/B8 (…/ITC instead of …/Other ITC related files, etc.). Section 5's description of A-R54/A-R55 and the author's template questions should reflect this.
- pywin32 was used but not declared from edac6a3 (31-08-2025) to 064c634 (23-07-2026). README.md:104-108 has never matched requirements.txt (openpyxl>=3.0.0 never existed).
- The texts `RefreshAll`, `CalculateUntilAsyncQueriesDone` and `DispatchEx` occur in no .py file at any of A's 46 commits (RAN, pickaxe). In A's git history the key-press refresh is the only refresh method ever used (relevant to Section 8 design decisions).
- Stale texts still call pywin32 optional (report_processor.py:11-13, extract_tab.py:21-23 and 101-105), although be6418c made it essential on Windows. The Step 4 notice cannot appear for a missing pywin32.
- Hover texts added in bf2f1dd disagree with the code: organise-only says 'no Excel automation' (setup_tab.py:96-97); the full pipeline 'then open each report in Excel to refresh' (:98-100); Resume 'skipping files already in place' (:108-109). This affects Section 3 (A-F5, A-F13, A-F31) and Section 7 (A-G66, A-G10).
- Every dead item listed in the registry (EXCEL_SAFETY, validate_results, cleanup_on_error, get_cached_value, _write_section, _create_details_sheet, auto_refresh_if_missing, show_timestamp, recent_folders, etc.) has had an unchanged number of mentions in every commit since edac6a3 (RAN, pickaxe): dead from day one, never wired in.
- The only changes ever made to [GST RULE]-tagged rules are: 47a1492 (A-R77–A-R82 added; A-R88 effect), bd176bf (A-R89: unticked reports show Missing), 5c8ce95 (A-R2), 33d0baf (A-R17/A-R23 key alignment; A-R60 start guard), 510f7b4 (A-R60 hand-over appended), 275d9a4 (A-R59 recognition), 7347383 (A-R54/A-R55 values) and 322717f (A-R1–A-R6 display only). All other GST rules are textually identical to edac6a3.
- Since 12d505f, Step 3 reports are saved by Excel in %TEMP%\gst_excel_temp and then copied. Excel may record that folder inside the workbook (x15ac:absPath; MEDIUM, not verifiable on Linux).

### UNREGISTERED items
- UNREGISTERED: "Invalid or corrupted Excel file" variation (edac6a3, removed a603c02; it could never fire)
- UNREGISTERED: old-format extraction mapping `power_query_extractor/config/cell_mappings copy.py` [GST RULE] (edac6a3, deleted 47a1492; could not be imported; four commented-out candidate cells)
- UNREGISTERED: check of typed values in the max-length box (10–100) and the hint "Names >10 chars will show warning" (edac6a3, removed 322717f)
- UNREGISTERED: dark-mode switch permanently enlarging the tabs via TNotebook.Tab padding=[20, 10] (edac6a3, removed 725fbaf)
- UNREGISTERED: pywin32 treated as optional on Windows (c463256/f363946, reversed 10 minutes later by be6418c)

### Corrections to the working notes
- Do the Power Queries in the ITC/Sales templates read the folder paths from the Links cells, or do any of them rely on the old folder names 'Other ITC related files', 'Sales related files' or 'GSTR-3B Exports' (renamed in 7347383)?
- [GST RULE] Are the 14 extraction cells added on 30-10-2025 (A-R77–A-R82) still at those sheets and cell positions in today's ITC template ⚠? And which cells are meant to feed the IGST/CGST/SGST columns of 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' (the code shows Diff GSTR-3B J15/K15/L15 there)?
- When both an IMS Reco file and its '(1)' duplicate download exist for a client, which should feed the ITC report? Since 5c8ce95, the '(1)' copy sorts first and is the one linked.
- [GST RULE] Is the two-letter state list (A-R21) your reference list ⚠? The tests also pin the quirk 'Unknown Place' → 'UP'.
- Before git (June–August 2025), did you try Excel's own refresh command through COM (RefreshAll / CalculateUntilAsyncQueriesDone) before settling on the Alt+A, R, A key presses? It never appears in A's git history.
- Did the move to the shorter share root assumed by 7347383 happen? Its own model leaves a thin margin (215 vs ~218 characters ⚠). Do you want the path-length guard the commit calls 'the durable fix'?
- What should 'Resume' do? The hover text added in bf2f1dd says it skips files already in place, but each run creates a new Version folder, so it only skips within the same minute.
- Which 'Refreshed File Suffix' values are used in practice, and should the suffix be remembered between sessions? Short suffixes such as _R{timestamp} make the original report count as refreshed.
- Should requirements.txt pin exact versions and leave out pytest (Standard 13)? Should the program ever run pip itself (a5ad487)?
- Which Python will users run (normal 3.14 or the free-threaded python3.14t mentioned in c463256)? Will the shared runtime include tkinter?
- Is the README's pywin32 post-install step still part of your set-up?
- gst_organizer.log and the compiled files with your folder paths are still readable at edac6a3 in both repositories. Do you want the history rewritten?

## From writer X345

### Facts other sections must reflect
- RAN (me) re-confirmed A-G1: the real DataConsolidator, given made-up values per configured cell, put Diff GSTR-3B J15/K15/L15 into the IGST/CGST/SGST columns of the IMS Reco, Extra In Purchase and Extra In GSTR2B sheets. 13 columns OK, 9 MISMATCH (data_consolidator.py:139-150). No test covers this (Sections 3 A-F24, 5 A-R88, 7 A-G1, 12).
- Mutations M08, M16 and M21 survive the whole suite (RAN, me). So the fixes of c89c1d4 (refresh_status after a theme toggle), 33d0baf fix 3 (per-client folder key by state code) and 2aaa2b1's Skip Refresh wiring in report_processor.py:140-147 have no test protecting them. Section 12 should not describe them as tested.
- tests/test_main_window_tabs.py rewrites the user's real ~/.gst_organizer_cache.json when it holds dark_mode true (RAN). It also creates %TEMP%/gst_excel_temp. Supports A-G60; relevant to Sections 9 and 13.
- The subprocess probe tests write .pyc files into the project's own folders even when pytest runs with -B. This matters for Standard 13 (rebuild_pipeline_FINAL.md:969-970) if tests ever run from a shared folder.
- HANDOVER.md and FILE_INDEX.md never existed in any commit of A or B, nor in the checkout (RAN). Sections 1 and 2 should state this.
- The Step 3 log says 'press Process to refresh' (gui/handlers/processing_handler.py:170-172), but the Step 4 button is labelled '🚀 Start Processing' (extractor_window.py:328-340). Relevant to Section 10 and A-G69.
- Docstring and comment errors for Sections 2, 8 and 10:
- gui/utils/setup_validation.py:68 `_short` says 'Trim the middle'; it trims the start and returns 45 characters.
- gui/tabs/validation_tab.py:27 says '40%/60%'; the left column is a fixed 450 px (:30).
- report_processor.py:541/543 logs 'read-only mode'; the file is opened read-write (:737).
- gui/widgets/tooltip.py:1 says tooltips show truncated paths only; they are also used for help texts since bf2f1dd.
- These texts still call pywin32 optional, although be6418c made it required on Windows: report_processor.py:11-13, gui/tabs/extract_tab.py:20-22, and the user-visible notice at gui/tabs/extract_tab.py:101-105.
- The version label 'v3.0 | Production Ready' and __version__ '3.0.0' are unchanged since edac6a3 through all 34 of A's own commits (A-R156).
- Commit-message test counts for Section 12A (RAN, collected per commit): 322717f gives a wrong breakdown (19+8, not 21+9); d64d3a9 says 140 but has 132; 2aaa2b1 says 20 new but added 21. From a677670 to aeb83cf, 1-4 test files could not load on a Python without pywin32 (ModuleNotFoundError: pythoncom), until c463256.
- Confidentiality for Section 2 and other writers: the company-like name occurs only in PROJECT_AUDIT_ANNUALSTATEMENT.md lines 130,133,134,158,160-162,177,179,180,182,227,228,231,232. It was introduced by f7eb8f1, is in 33 A commits and is not in B. Lines 5 and 663 of the same file hold the author's local path. Never quote either.

### UNREGISTERED items
- (none)

### Corrections to the working notes
- X6-Q1 [GST RULE] ⚠ The IGST/CGST/SGST columns of the IMS Reco, Extra In Purchase and Extra In GSTR2B sheets show Diff GSTR-3B!J15/K15/L15 instead of their configured cells (IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1, ExtraInGSTR2B!S1/T1/U1). Which cells are intended? Was this ever noticed in a past PQ_Extraction_Report? (data_consolidator.py:137-150)
- X6-Q3 [GST RULE] ⚠ What does Info!BB2 hold in the empty templates and after a good refresh? Is 'not empty / not 0 / not an error' the right success test? (report_processor.py:638-685)
- X6-Q11 [GST RULE] When a client has two files of one type (e.g. an original and a '(1)' copy), which one should feed the report? Or should the run stop? (utils/helpers.py:361; core/excel_handler.py:628-700)
- X6-Q12 [GST RULE] ⚠ Is 'all six types, no duplicates, any number of GSTR-3B' the set every client must have? (utils/constants.py:68-75; core/file_parser.py:252-286)
- X6-Q16 [GST RULE] ⚠ Period/year parts of file names are never checked. Should files of different periods ever be mixed? (core/file_parser.py:97-106)
- X6-Q20/Q21 [GST RULE] ⚠ Do real file names contain hyphens in client names? How is the state written: full name, letter code, or numeric code? (utils/constants.py:15-55; utils/helpers.py:592-618)
- X6-Q26 [GST RULE] What do the templates' Links cells hold for file types a client lacks? (core/excel_handler.py:192-193, 300-301)
- X6-Q33 What should 'Resume' do? Today it makes a new Version folder each run and skips nothing unless resumed in the same minute. (core/file_organizer.py:102-108, 283-287)
- X6-Q41 Do staff keep their own Excel open during Step 3 or Step 4? Has Excel ever closed by itself? (core/excel_handler.py:123-133, 238-246)
- X6-Q42/Q43 Do staff work on the PC during a Step 4 refresh? Which Excel version and UI language are used (Alt+A, R, A key tips)? (report_processor.py:258-293)
- X6-Q55 Should 'full pipeline' start the refresh itself, as its label promises? Is the 'Continue to extraction' button still wanted? (processing_handler.py:136-175)
- X6-Q57 How do you refresh an older Annual Statement batch? Pointing Step 4 at it finds nothing. (extractor_window.py:622-632)
- X6-Q62 How do staff start the program (console, pythonw, shortcut) and from which folder? (main.py:17, 20-27)
- X6-Q63/Q64 Where should settings and logs live? Should the program ever pip-install packages, and should that be the whole requirements.txt including pytest? (cache_manager.py:15; dependency_installer.py:88-158)
- X6-Q66 Is pywin32 meant to be required on Windows (be6418c)? If yes, should the Step 4 notice and comments stop calling it optional? (extract_tab.py:84-105; report_processor.py:11-13)
- X6-Q77 Did running the tests ever change your ~/.gst_organizer_cache.json? (tests/test_main_window_tabs.py:36-38)
- X6-Q78 d64d3a9 says 140 tests, but it has 132. Do the uncommitted harnesses named in seven commit messages still exist?
- X6-Q81 Is the company-like name in PROJECT_AUDIT_ANNUALSTATEMENT.md lines 130-232 a real client?
- X6-Q82 Are the repositories private? Should history be rewritten to remove gst_organizer.log, the .pyc paths and the share path in commit messages?

## From writer A-S45: Sections 4-5 data flow + rules

### Facts other sections must reflect
- Start-up failure: a settings file ~/.gst_organizer_cache.json holding valid non-object JSON (e.g. [1] or 5) stops the main window being built (AttributeError at gui/main_window.py:376). RAN me s4. Sections 3 (A-F27), 7 (gap) and 9 should reflect this.
- A missing client key (number-looking client text, or a Re-scan during Step 3) ends the WHOLE Step 3 run with 'Fatal error', because gui/handlers/processing_handler.py:232 sits outside the per-client try. It is not a per-client failure.
- Section 5 gives ⚠ to A-R104 (the 218-character figure) and A-R105 (error-dialog title words, button words and window classes), although the registry marks them not time-varying.
- Tag counts for summaries: 157 rules; 44 [GST RULE]; 29 ⚠ (27 registry + 2 added); 27 DEAD or dormant.
- Three UNREGISTERED decisions are documented in Section 5.18 (U1 Step 3 per-file log wording, U2 Step 4 log/progress format, U3 client order in each step). Sections 3 and 10 should use them.
- Step 3 and Step 4 logs use two different colour palettes (Step 3: #4CAF50/#FF9800/#F44336/#2196F3 on #1E1E1E; Step 4: #28a745/#ffc107/#dc3545/#000000 on #f8f9fa). Relevant to Section 10.
- In a Step 4 refresh status, a failed Info!BB2 check REPLACES an error dialog's text in the stored error (power_query_extractor/core/report_processor.py:354-356). Relevant to A-F21 / A-F26 in Section 3.
- For a missing sheet, _extract_data stores None under the field names only; the bare cell-reference keys are not added (power_query_extractor/core/report_processor.py:570-576). Relevant to the dormant calculations (A-R85).
- The Step 2 instructions text '⚠️ = missing files' (gui/tabs/validation_tab.py:54) is not exact: ⚠️ is also shown for 'Has duplicates' (core/file_parser.py:284-286). Relevant to Section 10.

### UNREGISTERED items
- UNREGISTERED: Step 3 per-file log line names the file TYPE, not the real folder ('📁 <original name> → <type> folder', e.g. 'Annual Statement' files logged as '→ Annual Report folder' although they go into the Version folder) — gui/handlers/processing_handler.py:281-284 (Section 5.18 U1)
- UNREGISTERED: Step 4 run log and progress format (log cleared at Start; header lines; per-report result lines; [HH:MM:SS] lines in colours info #000000 / success #28a745 / warning #ffc107 / error #dc3545; progress i/total×100, 95, 100, then 0 'Ready') — power_query_extractor/gui/extractor_window.py:875, 888-950, 985, 1012-1029, 568-571 (Section 5.18 U2)
- UNREGISTERED: Order of clients in each step (Step 2 = order of each client's first file in the path-sorted scan; Step 3 = table order; Step 4 list = file-system listing order, not sorted; consolidated rows = Step 4 order) — core/file_parser.py:233, utils/helpers.py:361, gui/handlers/client_handler.py:30, 46, 165, power_query_extractor/gui/extractor_window.py:638, 688, 849 (Section 5.18 U3)

### Corrections to the working notes
- A_older_docs.md §1.7 G6 says a Re-scan during Step 3 makes later clients fail with a KeyError that is 'caught and counted as a failed client (processing_handler.py:442-453)'. Wrong: the lookup self.app.client_data[client_key] is at gui/handlers/processing_handler.py:232, before the per-client try at :243, so the KeyError goes to the outer handler at :522-525 and ends the WHOLE run with '💥 Fatal error' and no summary workbook (TRACE T8 §2 shows the same for a missing key).
- Registry A-R114 (and A-F27) say a settings-file load error 'silently gives empty settings'. Nothing shows on screen, but gui/utils/cache_manager.py:25-26 does log ERROR 'Error loading cache: …'.
- Registry A-F27 / START note: 'valid JSON that is not an object stops all later saves'. RAN (me, s4_cache_shapes.py) adds more: a non-empty non-object value ([1] or 5) makes GSTOrganizerApp() raise AttributeError ("'list'/'int' object has no attribute 'get'") in apply_cached_values (gui/main_window.py:376), so the program cannot start. An empty list [] is treated as 'no settings'.
- Counting tags in A_registry.md with a plain '|' field split gives wrong numbers, because rows A-R18 and A-R59 contain escaped pipes (\|). The true counts are 44 [GST RULE], 27 TIME-VARYING and 27 DEAD (incl. A-R18, A-R25, A-R117). An earlier draft of this section claimed 58 / 33 / 28; that is now corrected.
- Clarification to registry A-R157 / TRACE T8 §7: on Linux, truncate_path treats a Windows-style backslash path as ONE part, so it returns '...' + the last 57 characters. The result on Windows differs (RAN me s3).

## From writer A-S3b: Section 3 features (part 2)

### Facts other sections must reflect
- A-F24 (RAN t24): the IGST/CGST/SGST columns of 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' (9 of 22 configured cells) receive Diff GSTR-3B!J15/K15/L15 values (data_consolidator.py:141-145). It is right only by accident when the Diff GSTR-3B sheet is missing. Began with 47a1492 (shared before the split) [GST RULE] - A-G1.
- A-F24: the ✓/✗ column follows only the refresh status, and any non-success status (failed BB2 check, a dialog seen, not selected) writes 'Missing' in every column even when values were read (data_consolidator.py:108-129). A ✓ row can hold only 'Missing' when a sheet name differs in letter case (RAN) - A-G17.
- A-F21/A-F22/A-F20: a refresh that fails is still saved under the Refreshed name (report_processor.py:349). It is then shown as the last refresh and reused by Skip Refresh as 'Skipped - Using existing refreshed file', success (RAN t22 1a-1b, t19b) - A-G2.
- A-F26 (RAN through the fake COM layer, Part 4): every top-level desktop window is examined (report_processor.py:450), and another program's dialog whose title contains 'could not', 'error', 'missing' etc. is answered 'Yes' - A-G55. When the BB2 check also fails, its message replaces the dialog text (report_processor.py:353-356) - new, LOW.
- A-F24/A-F25 (RAN t19): a re-scan during a Step 4 run re-enables Start and redirects the consolidated workbook to the folder scanned last, because self.level1_folder is read at the end of the run (extractor_window.py:945) - A-G65/A-G71.
- A-F25 (RAN t19, t21): every processor line is shown in the 'info' colour (report_processor.py:56-60), errors included. Warnings written only with logger.warning/error (sheet not found, Excel busy, dialog found, Excel not in front) never reach the screen. The red error tag is used only for a whole-run failure (extractor_window.py:958).
- A-F21 performance: the fixed waits come to 39 s per report at the default 10 s (36 s in the refresh plus 3 s), or 136 s + 3 s at 60 s; there are two Excel sessions per report and no Stop button (RAN count) - A-G50.
- A-F30: on a cp1252 code page, 51 of 154 log lines of one Step 4 report cannot be written (RAN count on a real log + simulation) - adds the size of the loss to A-G56. README's Debug Mode line has no effect where README says to put it, and placed before the extractor's set-up it leaves pq_extractor.log empty (RAN t30b) - new.
- A-F16 (RAN t16 R4): after Stop, the Step 3 summary workbook reports 'Total Clients 3, Successful 1, Failed 0'; the unprocessed clients appear on no sheet and nothing records the Stop - new.
- A-F19 (RAN): Step 4 only looks at the newest Version folder of each client (by date changed); a report present only in an older Version folder is not offered - new.
- A-F27 (RAN t27 8d): an unknown processing-mode value in the settings file is accepted at start-up (no radio selected) and a later Step 3 run fails with ValueError 'Invalid processing mode' (core/file_organizer.py:62-63) - new, LOW.
- A-F29: closing the main window during a Step 4 run gives no question, because on_closing checks only Step 3's is_processing (gui/main_window.py:344-355); the daemon worker and a visible Excel are left mid-refresh - A-G27.
- New LOW items found in A-F28/A-F29/A-F31: the standalone extractor ignores the saved dark mode (RAN t29); PowerQueryExtractorApp.__getattr__ replaces Tk's own forwarding, so app.eval fails (RAN t29b); a stray gst_organizer_cache.json in the current folder decides the scanned folder when the home file is missing (RAN t29 C); tooltips are not kept on the screen (RAN t31); the collapsible box colours only its outer frame (RAN t31); the 'EXPECTED FILE NAMES' card descriptions omit .xls, letter case, (N) markers and the exactly-8-digit IMS date [GST RULE].
- New LOW items found in A-F16/A-F18: 'Complete %' is written as text and 'Processing Date' as the run-stamp text (RAN); the organisation report lists every file of the client, including failed copies (READ); '✅ Step 4 ready' is logged even when the Step 4 scan found nothing (RAN t18_tab).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- Registry A-F25 says the Step 4 log box 'is never trimmed': it is never trimmed during a run, but it is emptied at every Start (power_query_extractor/gui/extractor_window.py:875).
- Notes say the standalone extractor window has no close handler: it has Tk's default WM_DELETE_WINDOW (destroy) (RAN t29) but no handler of its own, so closing during a run still ends it without asking. Also, inside the main window, the close question checks only Step 3's flag (gui/main_window.py:344-355), so closing during a Step 4 run gives no question at all.
- Registry A-F19 says 'a scan that finds nothing keeps the previous list': that is true only when no Annual Statement-* folder exists (early return, extractor_window.py:625-628). A run folder without Version folders clears the list to 'No clients found', while Start keeps its earlier state.
- Registry A-F27 says 'valid JSON that is not an object stops all later saves': a non-empty JSON list stops START-UP (AttributeError, RAN t27 8b), and a length stored as text also stops start-up (TypeError, gui/main_window.py:387). Only an empty list starts normally and then makes every save fail.
- Registry A-F17 / A-G72 say 'the ETA is one client behind': RAN shows this only at the summary step (after the last client it still shows one client's time); during the run the estimate is right for an even workload.
- Registry A-F28 says 'colours the program changes in light mode after the first switch are lost on restore': the mechanism was RAN with a stand-in label, but in A today the only colours the program changes at run time are the Setup checklist's (gui/tabs/setup_tab.py:553-566, 606-608), and those are recomputed after every switch (gui/main_window.py:180-194). So the effect is latent (MEDIUM, based on a search for colour changes).
- Registry A-F31 says the file-name card 'omits the 8-digit IMS date': the card does show 'DDMMYYYY'. What it omits is that the date must be exactly 8 digits (utils/constants.py:23), that .xls and any letter case are accepted, and the optional '(N)' copy markers.
- A-G68 says 'on Linux nothing scrolls by wheel': a synthetic <Button-5> did nothing (RAN t31), but whether a real wheel under Tk 9 on Linux also sends <MouseWheel> is NOT SURE.
- PROJECT_AUDIT_ANNUALSTATEMENT.md F10 (:305-317), :409-411, :450, :635 and :713 say the extractor reads the wrong settings path: fixed in c1583d6 (23-07-2026, A only). The home file is now checked first (RAN t29 case A); the extra candidate files in the current folder remain.
- PROJECT_AUDIT_ANNUALSTATEMENT.md:314 lists 8 saved settings keys: the code saves 9 (workflow_mode added) plus recent_folders (gui/main_window.py:360-370).
- PROJECT_AUDIT_ANNUALSTATEMENT.md:655 suggests logs live under the program folder's logs/ (created by main.py:35-37): the log file goes to the current folder (main.py:20-27; RAN t30).
- README.md:173 'Click Refresh & Extract': no such button exists; it is '🚀 Start Processing' (extractor_window.py:330).
- README.md:281-287 Debug Mode (logging.basicConfig(level=logging.DEBUG) in main.py or launch_extractor.py) does not work as written (RAN t30b). README.md:289-291 does not say that the logs go to the current folder.
- README.md:15, 25 says configured cell values are consolidated: true for 13 of the 22 cells only (RAN t24).
- Code comments that are wrong: cell_mappings.py:259, 263 ('Backward compatibility' / 'Build old format for compatibility') describe the very table the extraction uses (report_processor.py:546). setup_tab.py:478-479 ('can never drift') is wrong, because the card shows hand-written description strings, not the regular expressions. report_processor.py:541/543 says 'read-only mode', but the file is opened read-write.
- The history inference about the 9-column mismatch is now RAN, not MEDIUM: at edac6a3 every output column title contained a space, so the 'contains' look-up could never match. The mismatch began when 47a1492 added the single-word titles IGST/CGST/SGST.

## From writer A-S89: Sections 8-9 design + deps

### Facts other sections must reflect
- Python version changes names (ran on 3.11/3.12/3.13 vs 3.14): sanitize_filename turns a client name ending in '.' (e.g. 'ClientA Co.') into 'ClientA_Co' on 3.11-3.13 but 'ClientA_Co.' on 3.14. This changes report file names (gui/handlers/processing_handler.py:292), Level-4 folder suffixes (core/file_organizer.py:90) and copied file names. Section 3/5 (A-R20) and Section 11 should state that names depend on the Python version.
- Under pythonw (Standard 13's launcher) the start-up dies silently at main.py:17 (ran). Separately, _pause() (main.py:45-54) does not catch RuntimeError: 'lost sys.stdin', which input() raises when there is no stdin. This affects A-F1 / A-G22 and the Section 11 start-up design.
- The start-up install offer fails on any Python without pip ('No module named pip', ran). Standard 13's embeddable runtime normally has no pip, and its embeddable package normally has no tkinter (MEDIUM, general knowledge, not checked here). Relevant to A-G53/A-G59 and to Section 11's runtime plan.
- The standalone extractor (launch_extractor.py) runs no requirements check, and since ba2a753 depends on the main app's utils package (power_query_extractor/core/report_processor.py:39). It finds that package only through the script folder on the search path; launch_extractor.py:10 appends the folder rather than inserting it first. Relevant to A-F29 and A-G32.
- README.md:68-95's folder tree omits utils/excel_com.py, utils/requirements_check.py and utils/dependency_installer.py; following it literally gives 'Could not start'. The README's requirements block (:104-108) was never true. Section 12 docs drift and the Section 11 setup steps should use the real list.
- Step 4 refresh depends on the Windows Script Host COM object WScript.Shell (power_query_extractor/core/report_processor.py:260). Standard 13 notes VBScript is being switched off around 2027 ⚠; whether WScript.Shell survives that is NOT SURE. Add to A-G79 / Section 11 risks.
- Excel itself is never checked at start-up (only pywin32 is). Without Excel every Step 3 report fails with advice blaming the Links sheet (gui/handlers/processing_handler.py:322-323, 396-398). Section 3 A-F14/A-F15 and Section 10 messages should reflect this.
- ExtractorPanel.is_processing is read on tab change (power_query_extractor/gui/extractor_window.py:841) but never checked before a new Start (:845-883), and any scan re-enables Start (:755-757). This matches A-G65; Section 3 A-F20/A-F21 should not describe a run guard for Step 4.
- Python 3.7 is stated as the minimum (requirements.txt:2, README.md:50). The code grammar fits 3.7, but the package versions that ran here need 3.8 (openpyxl 3.1.5) and 3.10 (pytest 9.1.1), and only 3.14 was ever tested ⚠. Section 1/11 should not repeat '3.7+' without this caveat.

### UNREGISTERED items
- (none)

### Corrections to the working notes
- A_older_docs.md RM12 and 4.1 item 1 say README's folder layout omits gui/tabs/extract_tab.py. README.md:83-87 lists the gui sub-folders (handlers/, tabs/, utils/, widgets/) only as folders, so copying tabs/ brings extract_tab.py along. Only the three utils files are actually missing: excel_com.py, requirements_check.py and dependency_installer.py.
- A_code_startup_widgets.md (requirements.txt table) gives openpyxl's lxml switch as LXML = lxml_available() and OPENPYXL_LXML == 'True' without saying the variable defaults to 'True' when unset (openpyxl/xml/__init__.py: os.environ.get('OPENPYXL_LXML', 'True')). The note's conclusion that lxml is used whenever installed is right; I ran it and openpyxl.xml.LXML is True with the variable unset.
- A_dead_code_patterns.md 2.7 and A_code_startup_widgets.md say importing gui.utils.setup_validation loads 222 modules. My run loaded 334. The count depends on what the interpreter had already loaded; the real finding (tkinter, openpyxl, gui.main_window and the extractor all get loaded) is confirmed.
- A_history_part1.md says only the ten July commits in its range carry an AI co-author line. A git log check shows all 32 July 2026 commits (a603c02..bf2f1dd) carry a Co-Authored-By line.
- A_outside_files_state.md 5.1 cites utils/requirements_check.py:66-90 (and 1.2 cites :73-89) for the pywin32-essential logic. The code is REQUIREMENTS at :37-61 with essential=ON_WINDOWS at :57; the registry already records this.
- Older audit PROJECT_AUDIT_ANNUALSTATEMENT.md:692-693 (D6) cites the comment 'Use state code instead of full name' at file_organizer.py:87; it is now at core/file_organizer.py:91. This is a document-versus-code drift, listed in Section 8.9.

## From writer S2

### Facts other sections must reflect
- core/excel_handler.py:241-246: the 'Don't quit if Excel was already open' branch can never run because self.excel_app is only ever None (:56). Excel is always quit, including an Excel the user already had open (attached at :124). Sections 3 (A-F14/A-F15), 7, 8 and 11 must reflect this (A-R97).
- gui/tabs/extract_tab.py:67-69, 84-105: the 'Power Query refresh is unavailable - needs pywin32' card cannot appear for a missing pywin32 (RAN: EXTRACTOR_IMPORT_ERROR is None without pywin32). It would appear only for another import error, and would then blame pywin32 (A-R129). Sections 3 (A-F18), 6/7 and 10 must not describe it as working for the pywin32 case.
- Dead functions and methods (16, Section 2.8.3), e.g. cleanup_on_error, validate_results, _write_section, _create_details_sheet, get_unique_path, calculate_file_hash, the three main_window delegators and UIHelpers.update_status, must not be described as features in Section 3. Section 11 should rebuild them only if the author asks.
- Most dead code dates from edac6a3 (shared - before the split on 30-10-2025), so B probably inherited it; B's Section 2 and the comparison should check (Section 2.8.13).
- main.py:17 makes A die silently with no console (pythonw or a no-console shortcut; RAN). launch_extractor.py skips main.py's package check. power_query_extractor/extractor_main.py runs only via `python -m` (RAN). Relevant to Sections 9, 10 and 11 and to the Standard 13 comparison.
- ReportProcessor._log is never given a level (97 calls), so every Step 4 processor line, errors included, shows in the plain 'info' colour (A-F25; Section 10).
- README.md:71-95 folder list omits utils/excel_com.py, utils/requirements_check.py, utils/dependency_installer.py and gui/tabs/extract_tab.py; a copy made from it cannot start (Sections 9 and 11).
- Package __init__.py re-exports are unused and stale (ExtractTab, Tooltip, setup_validation and the newer utils modules are missing). gui/__init__.py:7 makes any gui sub-module import load the whole main window (Sections 8 and 11).
- main.py:35 creates logs/ and temp/ in the program folder, and nothing ever uses them (RAN). Logs go to the current folder (Sections 9 and 10, Standard 13).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- /root/merge-scratch/audit/A_files.txt gives `wc -l` counts, one lower than the real line count for the 37 files without a final newline (e.g. .gitignore 79 -> 80, core/__init__.py 12 -> 13, power_query_extractor/config/__init__.py 0 -> 1). RAN.
- A_code_gui_tabs_handlers.md section E: 'app.processing_thread is written but never read' - it is read once at gui/handlers/processing_handler.py:134 to start the thread; never read again or joined. READ.
- A_dead_code_patterns.md section 0: 'A-copy already contains __pycache__ folders' - at this task's time A-copy contains none (RAN find) and matches the repository blob-for-blob (RAN hash check). Environment remark only.
- A_dead_code_patterns.md section 1.11 cites the unused import PowerQueryExtractorApp at tests/test_extractor_panel.py:11; the statement starts at :11, the name is on :12. READ.
- A_dead_code_patterns.md section 1.10 misses three never-supplied parameters: _pause(message) (main.py:45), ReportProcessor._log(level) (report_processor.py:53; 97 calls, none passes level, RAN AST), and the process_itc/process_sales 'backward compatibility' defaults (report_processor.py:87, 97; the only caller always sets both at extractor_window.py:855-856). Additions.
- Additions not in the dead-code notes: the 'no stat' branches utils/helpers.py:144-147 and :293-295 (tests only), the .xltx/.xltm endings in utils/helpers.py:299 (unreachable in the program), the safe_copy_file(verify=True) default, styles['title']['alignment'] (core/excel_handler.py:422), the variation field 'path' (core/file_parser.py:214), unused OLE_SIG in tests/conftest.py:15.
- Registry corrections for title_bar.py:32 (command=) and tooltip.py:39/58/67 (except clauses) were re-checked and are right; the dead-code notes' citations there are wrong, as the registry already says.

## From writer S10

### Facts other sections must reflect
- Dark mode (A-F28): the Step 4 log's 'info' lines keep the fixed colour #000000 (power_query_extractor/gui/extractor_window.py:568) on the dark log background #3b3b3b. Almost all Step 4 log lines are black on dark grey and hard to read. RAN, shot 38.
- Dark mode (A-F28): the Step 3 log background changes from #1E1E1E to the lighter #3b3b3b (the Text widget is treated as an input field). RAN.
- Dark mode (A-F28): every coloured classic-Tk button becomes the same #404040 grey (Step 2 DRY RUN / EXPORT LIST / START, Step 3 START PROCESSING / STOP, Step 4 buttons including the blue Start Processing bar). Colour-coded meaning is lost. RAN, dumps 17/19/20.
- Dark mode (A-F28): Step 4 client rows created while dark (a scan or the 'full' hand-over) stay white with white tick boxes. RAN shot 36; the registry has this only as READ for Step 4.
- Standalone extractor (A-F29): the selected sub-tab is drawn smaller than the unselected one. configure_extractor_styles sets padding [20, 10] (extractor_window.py:95), but 'clam' maps a smaller padding for the selected tab. RAN, shots 29/30.
- Step 3 (A-F6/A-F17): with the max-length box emptied, Start ends at once with log '💥 Fatal error: expected floating-point number but got ""' and the box 'Processing Error' showing that raw text, raised from the worker thread. No folder is created. RAN, tour 6.
- Step 4 (A-F24): if the consolidated workbook cannot be saved, the box 'Error' / 'Processing failed: [Errno 2] No such file or directory: …' appears. Start is re-enabled and the progress label goes back to 'Ready'. RAN, tour 6.
- Keyboard (A-R135): the Step 2 / Step 3 buttons have an empty command. Invoking them the way Space would does nothing, so they cannot be used from the keyboard. RAN, tour 5.
- Install window (A-F1): the 'Installing…' window is a plain 760×380 text area with no button, no progress bar and no 'please wait' text. It shows only 'Installing into: …' and 'Running: …' until pip ends. RAN, shot 37, with a stand-in for pip. With an optional package missing, the 'Installation failed' text ends without the 'cannot start' sentence.
- Links-sheet messages (A-F14/A-F15): the start-up log line 'Templates must have Links sheet' and the hint 'Check if template has Links sheet' are never triggered by a missing Links sheet. The Excel (COM) path creates a new Links sheet and reports success (core/excel_handler.py:166-183); the openpyxl path writes nothing and reports success (:290-310). READ, HIGH.

### UNREGISTERED items
- UNREGISTERED: client list order - the Step 2 table lists clients in the order the scan first met them in a name-sorted file list (utils/helpers.py:361; dictionary order), not alphabetically (made-up files gave ClientA, ClientC, ClientB). The Step 4 list uses the folder-listing order, unsorted (power_query_extractor/gui/extractor_window.py:638). No registry rule states either; A-R60 only says the run follows 'table order'. RAN.

### Corrections to the working notes
- No reader note was found to be wrong on a point used in this section. The notes said Step 4 rows created while dark are not themed, confirmed only by reading; this is now confirmed on screen (shot 36).
- Fixes to the earlier, interrupted draft of S10.md: (a) the 10.1 table said 'RAN (texts of 8 of them)' for the message boxes; now all 25 have been seen in runs (23 in this section's runs, M17/M18 in GUI exp8). (b) Shots 30-33 cited by the draft do exist (an earlier listing of mine was cut short). (c) The two Step 2 row-order labels were made one: 'UNREGISTERED: client list order'.

## From writer S12own

### Facts other sections must reflect
- Several July commit messages state test counts that are wrong. 322717f says '21 unit tests; 9 smoke tests'; the real numbers are 19 and 8. d64d3a9 says '140 tests pass'; the suite has 132. 2aaa2b1 says '20 new'; it adds 21. Any section that quotes these numbers should use the counted ones (RAN me --collect-only).
- From a677670 (25-07-2026 23:01) to c463256 (26-07-2026 23:10), the test suite cannot even be collected on a Python without pywin32. From bfcb24a (26-07-2026 20:40) to c463256, the whole program cannot start without pywin32. At HEAD the suite runs without pywin32: 288 collected, 287 passed, 1 Linux-only failure, 3.28 s (RAN me).
- 33d0baf stopped abbreviating underscore-joined client names. They were not shortened between 23-07-2026 08:48 and 25-07-2026 18:45, until 4219ffb. Run-together names have never been abbreviated since 33d0baf. At HEAD the Level-2 folder key (create_client_state_key, utils/helpers.py:620-644) still keeps 'Private'/'Limited' for underscore-joined names (RAN me). This matters for Sections 5 (A-R19), 7 and 11.
- 7347383 changed the folder paths written into the Links cells of every new ITC and Sales report (A-R54/A-R55 values; [GST RULE] ⚠). The design note b276542 put 'any change to how reports are produced' out of scope. Whether any template query names the old folders is a question for the author.
- The code comment at utils/constants.py:87-90 calls the figures 235/27 and 215/0 'Measured'. The 7347383 message says they are modelled and assume a planned shorter share root; the measured figures were 270 characters and 1,094 files over the limit.
- Without pywin32, the Step 4 'Power Query refresh is unavailable' notice (gui/tabs/extract_tab.py:84-105) cannot appear: at HEAD EXTRACTOR_IMPORT_ERROR is None and the panel is built normally (RAN me). The 'optional pywin32' texts (report_processor.py:11-13, extract_tab.py:21-23, 101-105) contradict utils/requirements_check.py:49-57.
- tests/README.md is stale. It lists only 4 test files and says 'No GUI ... (<1s)'. At HEAD there are 18 test files and 288 tests, the suite includes Tk and subprocess tests, and it takes about 3.3 s.
- Registry A-F25 says the Step 4 log 'is never trimmed'. ExtractorPanel.start_processing clears the log at the start of every run (power_query_extractor/gui/extractor_window.py:874-875). Section 3/10 wording should say 'never trimmed within a run'.
- b276542's design note has false claims about tests. It says _classify_version_reports and _latest_refresh_times are unit-tested, and it mentions 'existing _log_async tests'. Neither was true at b276542. _classify_version_reports is still untested at HEAD, and no _log_async test has ever existed (RAN me grep).
- a5ad487's message says pip output is streamed 'so it never looks frozen'. It is not: subprocess.run with capture_output waits until pip ends (utils/dependency_installer.py:47-50). This supports gap A-G53.
- The old GitHub name goyal-sidhant/AnnualStatement resolves to A's HEAD bf2f1dd (RAN me ls-remote). A is probably the original repository, renamed (MEDIUM). This is relevant to Sections 1 and 9 and to the comparison.

### UNREGISTERED items
- (none)

### Corrections to the working notes
- A_history_part3.md (bfcb24a entry) says that 'from 510f7b4, even unrelated test files that import gui.* cannot be collected'. RAN (me) at bfcb24a without pywin32: the suite already fails at collection there with 2 errors, tests/test_extractor_panel.py AND tests/test_processing_handler.py (ModuleNotFoundError: pythoncom). So unrelated gui-importing test files fail from bfcb24a, not from 510f7b4. tests/test_main_window_tabs.py run alone gives 7 failed, which agrees with the notes.
- The task text says ls-remote facts are in the notes. None of the notes contains any ls-remote output. I ran git ls-remote myself on 21-10-2026. goyal-sidhant/AnnualStatement_oldv1 has only HEAD and refs/heads/main, both at bf2f1dd. The old name goyal-sidhant/AnnualStatement gives the same two refs. B's remote has only HEAD and main at b0e7d0e.
- The notes call A/claude/pensive-feynman-0mz1aa 'not a separate line of work', which is true. A detail they leave out: the checkout's reflog shows the branch was created on 21-10-2026 03:04 UTC by a checkout in this audit environment. It is not the author's branch.
- c1f425e: the commit message and the code comment say that catching TclError and RuntimeError stops a leaked Excel instance by keeping the cleanup in finally. ReportProcessor.cleanup() is an empty method (power_query_extractor/core/report_processor.py:801-804; also at c1f425e :759-762). A_history_part2.md calls the change of order 'harmless' but does not say that cleanup does nothing.
- A_history_part2.md and the registry give the shell's __getattr__ problem as RAN (H-A2). I confirmed its premise (RAN me): tkinter.Tk defines its own __getattr__ ('Delegate attribute access to the interpreter object'), and the shell's version replaces it. This is not a correction, only a confirmation.

## From writer X12

### Facts other sections must reflect
- Section 3 A-F14/A-F15 and Section 7: S1 (core/excel_handler.py:123-133, 238-246) attaches to the user's running Excel with GetActiveObject and hides it without ever restoring Visible. It sets DisplayAlerts back to True right before Quit(), and always quits because self.excel_app is never set (RAN fake COM R2 a).
- Section 6/7: S1's finally (core/excel_handler.py:239-240) is unprotected. When Excel is busy or crashes, the exception escapes, Quit/CoUninitialize/staging delete are skipped, the staging file stays in %TEMP%\gst_excel_temp, and the window shows the clean-up error ('Call was rejected by callee') instead of the real one (RAN fake COM R2 c, g).
- Section 3 A-F21/A-F22: when every open of the refreshed copy fails, the unrefreshed copy stays under the _Refreshed_ name. A later Skip-Refresh run reuses it as success='Skipped - Using existing refreshed file', and its old values reach the consolidated workbook with ✓ (RAN fake COM R3 f, f2). _process_report's 'Open method of Workbooks class failed' branch (report_processor.py:169-181) is unreachable, because S2 catches every exception itself (:388-393).
- Section 3 A-F26: the error-dialog scan clicks the 'Yes' button of ANOTHER program's #32770 dialog whose title contains an error word, and records that program's text as the report error (RAN fake desktop R3 c).
- Section 3 A-F21 / Section 9: fixed waiting is 39 s per report and 78 s per client at the default 10 s wait (RAN, counted from A's own sleep calls); 139 s per report at 60 s (READ).
- Section 9 / Section 7: the start-up check (utils/requirements_check.py:64-70) uses find_spec on 'pythoncom' only, so a pywin32 that is present but broken passes it, and Step 3 falls back silently to openpyxl, dropping Power Query (RAN R4, simulated). The standalone launcher (launch_extractor.py) runs no dependency check at all.
- Section 3 A-F18 / Section 10: gui/tabs/extract_tab.py:101-105 says 'creating the reports works without pywin32' — no longer true since be6418c. The 'needs pywin32' card cannot appear for a missing pywin32 (RAN R4: EXTRACTOR_IMPORT_ERROR is None).
- Section 9 / Standard 13: under pythonw.exe main.py:17 raises AttributeError before any log exists (RAN R5). gst_organizer.log and pq_extractor.log land in the CWD (RAN R5), and with the Standard 13 launcher that is the shared program folder. Settings are in %USERPROFILE% rather than %LOCALAPPDATA%\<App>. logs/ and temp/ are created in the program folder at every start (RAN R5).
- Section 12/registry A-R39: the 215-character worst case quoted in utils/constants.py:85-90 depends on a planned shorter share root (commit 7347383 message). The real measured run had a worst path of 270 characters with 1094 files over 218.
- Section 7: there is no DataMashup check before the openpyxl save at core/excel_handler.py:306; RAN R1 lost customXml/item1.xml, itemProps1.xml, its rels and xl/connections.xml while returning True. Nothing prevents Step 3 and Step 4 from driving Excel at the same time, and closing the window during Step 4 asks nothing (gui/main_window.py:344-355).
- Section 3 A-F24: R3 re-confirmed (RAN) that the 'IMS Reco' and 'Extra In Purchase' columns carry the Diff GSTR-3B J15/K15/L15 markers, while 'Reco Diff' carries its own configured cells.
- Section 8 pattern consistency: every ReportProcessor line, including '❌' errors, reaches the window with level 'info' (RAN fake COM R3, R4).

### UNREGISTERED items
- (none)

### Corrections to the working notes
- A_outside_files_state.md §1.2 cites utils/requirements_check.py:73-89 as making pywin32 essential, and §5.1 cites :66-90. The code has REQUIREMENTS at :37-61 (the pywin32 entry at :44-60, essential=ON_WINDOWS at :57), _is_installed at :64-70 and check_requirements at :73-74. The registry also records this.
- A_excel_pq_safety.md §3.2 (S2, rule IgnoreReadOnlyRecommended) says the option is used on 'the 6th and last open attempt; attempts 1-5 (737, 751, 766, 774)'. In report_processor.py:_safe_open_workbook the helper makes at most four Workbooks.Open calls: 774 and 777-783 are alternatives, and 774 runs only if a \\?\ prefix was added. An exception from the short-path Open at :766 is swallowed by the bare except at :767-768 and logged as 'Could not get short path'. RAN (fake COM R3 f) showed three Open calls, the third with IgnoreReadOnlyRecommended=True.
- A_excel_pq_safety.md §9.1 places the first-file index [0] at core/excel_handler.py:628, 639, 650, 666, 682, 698. The [0] index is on lines 629, 640, 651, 667, 683, 699; each statement starts one line earlier. A_live_traces.md T3 has the right numbers.
- A_excel_pq_safety.md E4 / §5.2 says that on this Linux file system the refreshed copy is listed before the original by glob('ITC_Report_*.xlsx'). In my run R6 on the same machine the original was listed first. The order depends on the file system and is not stable, so neither order can be relied on.
- A_outside_files_state.md §6 cites Standard 13's report_fatal_error block as rebuild_pipeline_FINAL.md:980-1023. The main.py code block is at :984-1029, with sys.excepthook at :1024.
- Registry A-R39 repeats the code comment at utils/constants.py:85-90 ('the long forms ... pushed the worst path to 235 chars with 27 files over; short forms bring the worst case to 215 with none over') as figures from the real run of 3030 files. The message of commit 7347383 says the real run's worst path was 270 characters, with 1094 files (36%) over 218. The 235/27 and 215/0 figures were modelled with a planned shorter share root. The comment drops that condition.
