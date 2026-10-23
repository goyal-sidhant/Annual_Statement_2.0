# MERGE_DECISIONS.md - draft of Parts 3, 4 and 5 (task D3)

Draft for the assembler. Local item ids (P3-1, P4-1, P5-1 ...) are renumbered 1, 2, 3 ... in the assembled list. "A" = AnnualStatement_oldv1 at `bf2f1dd` (27-07-2026; its window is built with Tkinter, Python's built-in window toolkit; 288 tests); "B" = AnnualStatement_oldv2 at `b0e7d0e` (09-02-2026; a "v4.0" window built with PyQt5, an add-on window toolkit based on Qt; no tests). The names mislead: A is not the older project and B is not the newer or better one. "MERGE_COMPARISON.md section n" points to the comparison for detail. A "GST rule" is anything that decides a tax figure; which return, table, sheet or cell a figure belongs to; a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. Values that change over time carry ⚠. Evidence marks: READ = established by reading the code; RAN = the code was run on made-up inputs in a scratch copy; HIGH / MEDIUM / LOW = how sure an inference is that was not run.

**Words used in Parts 3 to 5** (each explained once here):

- **The four steps.** Step 1 Setup (choose the source folder of downloaded returns, the ITC and Sales templates, and the target folder), Step 2 Validation (the list of clients found), Step 3 Processing (copy each client's files into folders and build its ITC and Sales reports), Step 4 Extract (refresh the reports and collect their figures).
- **Template, Links sheet, Power Query, refresh.** A template is your ITC or Sales Excel workbook, copied once per client to make that client's report. Its Links sheet is where Step 3 writes the client's folder paths and file names. Power Query is Excel's data-import tool inside the template; its queries read the Links sheet to load the client's files. To refresh is to make Excel re-run those queries.
- **openpyxl, pywin32, COM.** openpyxl is a Python add-on that writes Excel files without Excel, but silently drops Power Query (and drawn shapes such as text boxes) when it saves a workbook. COM is Windows' way for one program to drive another; pywin32 is the Python add-on that lets the tool drive Excel through COM.
- **Folders made by Step 3.** A run folder `Annual Statement-<DDMMYY HHMM>` per Fresh run; inside it a client folder `<client>-<state code>`; inside that a Version folder `Version-<DDMMYY HHMM>` per run; inside that the category folders (GSTR-3B, ITC, Sales). **Fresh / Re-Run / Resume** are Step 3's processing modes: a new run folder, a new Version folder in the newest run folder, or the same while skipping files already copied.
- **Summary workbook** = the Excel file Step 3 writes to summarise a run. **Consolidated workbook** = Step 4's one workbook with every client's extracted figures.
- **Network share** = the folder on the office server. **Console** = the black text window. **Settings file** = the small file in which the tool remembers your paths and choices. **Log file** = a text file in which the program records what it did.
- **The window's own line of work** = the single sequence of steps that also redraws the window (a "thread"); slow work done there freezes the window. A **background worker** is a second line of work beside it.
- **Function / helper** = a named piece of code that does one job. **Dead code** = code that exists but is never used by the running program.

## Part 3 - Found in only one project

Every feature or rule that exists in only one of the two projects: the "only A" / "only B" rows of the cross-project map (MAP.md sections 1 and 2), the "only" rows of MERGE_COMPARISON.md section 5 (Categories 13 to 16), the three features B lost in its rewrite (section 3.3), and features that sit inside a "different" row but exist on one side only. The four screens are the same in both. **No GST rule exists in only one project:** the 13 rules the map marks "only A" or "only B" (A-R110, A-R111, A-R113, A-R125, A-R129, A-R133, A-R136, A-R137, A-R142, A-R146, A-R147, A-R154, B-R118) are all untagged, and every GST-rule difference is a pair of versions, which Part 2 sets out. Where an item below touches a GST rule, that part is pointed to Part 2 or Part 6 and is not part of the recommendation. One-sided pieces that Part 1 or Part 2 already asks about inside a wider item (for example B's Stop button for Step 4, or A's hover help) are not asked twice; the list "Asked in another part" at the end of Part 3 says where each one is.

### P3-1. A's automated tests                                  Part 3 · refs A-G60 (A's `tests/` folder has no feature ID)
What this is about: A has 288 tests - small programs that check the code automatically, for example by feeding it a made-up name such as `ImsReco-<Client A>-Delhi-31032026 (1).xlsx` and checking that it is recognised as an IMS Reco file. B has none. Nothing old is copied into the new tool, so the question is what becomes of A's tests.
A's way: 18 test files in `A:tests/`, added after the split from commit `6c1e1c6` (23-07-2026, 81 tests; a commit is one saved step in a project's history) to `bf2f1dd` (288). They cover file-name recognition [GST RULE] ⚠, name cleaning, state codes and client keys, which files the scan accepts, the Setup checklist, dark mode, Step 4's file names, the start-up checks and scrolling. Run here: 287 passed, 1 failed (a Windows-only path case that passes on Windows).
B's way: no test at any of its 17 commits; the B audit lists 8 B faults that a simple test would have caught on the day they came in (B audit Appendix X3.5), for example the dead EXPORT LIST button and the always-empty "Variations" sheet.
Evidence: copying A's test files into the new project is ruled out: no old code goes into NEW, and your pipeline writes its own tests before the code (Build Loop step .2, "Tests first", `rebuild_pipeline_FINAL.md:344`). Run against B's shared code, 7 of A's tests fail - 4 pin (check for exactly) A's IMS "(N)" pattern [GST RULE] ⚠ (Part 2) and 3 pin A's name cleaning. The tests do not cover the organiser (the part that copies files into folders: 0 of its 12 functions entered), the report builders (1 of 15), the consolidated workbook (0 of 6) or any real Excel session, and one test can rewrite your real settings file (A-G60). How the new tool is tested is Part 8 (Testing). MERGE_COMPARISON.md section 8.1.
Recommendation: Keep them as reference only - their scenarios go into the reference note on A's tests offered in Part 9, and any test that pins a GST rule follows your Part 2 answer - because they record edge cases met in real use (such as the IMS "(1)" downloads) that the new tests should not lose.
Your answer: ____

### P3-3. Offer to install missing add-ons                                  Part 3 · refs A-R125, A-F1 (offer part), A-G53, A-G59
What this is about: when an add-on is missing, A offers (Yes / No) to install it with pip (Python's own package installer) and restart. Example: a new article assistant's PC, set up for the GSTR-9 season, lacks pywin32 and the assistant clicks "Yes".
A's way: "Yes" runs `<this Python> -m pip install -r requirements.txt` - the whole list, pytest (the test tool) included - in a window that waits up to 600 seconds, then restarts the tool (`A:utils/dependency_installer.py:26-158`; `A:main.py:78-95`; `a5ad487`, 26-07-2026; 13 tests).
B's way: nothing comparable.
Evidence: your Standard 13 says "Nothing is installed on users' PCs" (`rebuild_pipeline_FINAL.md:924`). With the shared portable Python in the apps folder (`_runtime`), one user's "Yes" would change it for every app and user (A-G59); that embeddable Python normally has no pip (MEDIUM), and on a Python without pip the offer fails (RAN, A audit Section 9 and 11.1.4). The window is blocked while pip runs (A-G53); versions are not pinned ⚠. The start-up check itself is P1-1, whose recommendation also leaves this offer out. MERGE_COMPARISON.md section 5, Category 15.3.
Recommendation: Drop, because under Standard 13 packages go into the shared runtime once, by you, with exact versions.
Your answer: ____

### P3-4. Live Setup checklist                                  Part 3 · refs A-F4, A-R113, A-R133, A-G20 (part), A-G49, A-G75 (part)
What this is about: while you type or browse the four paths on Step 1 (source folder, ITC template, Sales template, target folder), A shows ✓ / ○ / ✕ for each and a "Ready - all 4 items set" label. Example: the path of the FY 2025-26 GSTR-2B and GSTR-3B downloads folder on the network share, typed with a mistake, shows ✕ "Folder not found" before you press anything.
A's way: re-checks 350 ms (a third of a second) after typing stops, remembers answers per path, enables Scan only when the source folder exists (`A:gui/tabs/setup_tab.py:282-314, 506-608`; `A:gui/utils/setup_validation.py:13-115`; `322717f`, 25-07-2026; 19 logic and 13 window tests).
B's way: a fixed 7-line INSTRUCTIONS card; paths are checked only when SCAN or START is pressed (`B:gui/views/setup_view.py:469-486`; `B:gui/pyqt_main_window.py:218-225, 352-361`).
Evidence: A's own defects (RAN): a file passes as the source folder and a folder as a template; a ✕ stays after the folder is created until restart or a dark-mode switch (A-G20); the checklist does not control Step 3's Start, which keeps its own weaker check (A audit Section 3, A-F4); the checks run on the window's own line of work, so an unreachable share can freeze the window (MEDIUM, A-G49). MERGE_COMPARISON.md section 3, row 4; section 5, Category 16.5.
Recommendation: Keep, with real folder-or-file checks and the checks moved off the window's line of work, because it catches wrong paths before a run instead of in the middle of one.
Your answer: ____

### P3-5. Workflow choice: "organise only" or "full pipeline"                                  Part 3 · refs A-F5, A-R154, A-G66 (part)
What this is about: before Step 3, A asks what the run should do. Example: in October you only sort the clients' files and build the reports; in December you go on to refresh them (re-run each template's Power Query so it pulls in the client's files) and collect the figures in Step 4.
A's way: "organise only" (default) stops after Step 3; "full pipeline" then opens Step 4, fills its folder and lists the clients - it never starts the refresh itself (`A:utils/constants.py:163-172`; `A:gui/handlers/processing_handler.py:136-175`; `510f7b4`, 26-07-2026; 7 tests).
B's way: no choice; after every Step 3 run it rebuilds the Step 4 list and switches to Step 4 - in effect A's "full", and it never starts the refresh either (`B:gui/pyqt_main_window.py:384, 401`).
Evidence: A's texts promise more and less than the code does: "Also refresh Power Query and extract the data" (the refresh is not started) and "Fast - no Excel automation" (Step 3 still drives Excel); the hand-over also runs after Stop and breaks Step 4's link to Step 1's target (RAN A audit). MERGE_COMPARISON.md section 3, row 5; section 5, Category 16.10.
Recommendation: Keep A's choice, with texts that say exactly what each option does, because you added it deliberately in A's last week of work (`510f7b4`, 26-07-2026) and its faults are only its wording and the hand-over after Stop.
Your answer: ____

### P3-7. Warning about long client names                                  Part 3 · refs A-R136, B-R141 (note part)
What this is about: a long name such as "<Client A> Infrastructure Private Limited" makes folder paths long, and Excel cannot open files whose full path exceeds about 218 characters ⚠. A warns before this happens; B only promises a warning in its help note.
A's way: warns at a fixed 10 characters when you switch a client's name toggle on (`A:gui/handlers/client_handler.py:191-199`), and at Start against your length setting, but only when the global tick is on (`A:gui/handlers/processing_handler.py:86-113`).
B's way: no warning; its note "Names >10 chars will show warning" (`B:gui/views/setup_view.py:385-388`) describes code that does not exist in B.
Evidence: the two thresholds disagree (10 against the setting), and names added per client are not covered at Start (A audit 11.4.4, S3a-10). The long-path problem is real: in one real run 1094 of 3030 files were over Excel's ~218-character limit ⚠ (A's commit `7347383` message). Whether names can be added per client at all is P1-9; a check of every whole path against Excel's limit is proposed separately as contract C13 (Part 7). MERGE_COMPARISON.md section 5, Category 16.1.
Recommendation: Keep, as one warning before Step 3 for every folder name longer than your length setting, whichever option added the name (the A audit's plan for the rebuild, 11.4.4 S3a-10), because one warning before the run is cheaper than reports failing file by file.
Your answer: ____

### P3-8. Export the ticked client list                                  Part 3 · refs A-F12, A-R142, B-F10, A-G31, B-G73 (part)
What this is about: a button on Step 2 that saves the names of the ticked clients to a text file - for example to send a colleague the list of clients going into this batch. The rest of the Step 2 list is P1-8.
A's way: writes the file and opens it (`A:gui/handlers/client_handler.py:225-257`; unchanged since the first commit `edac6a3`).
B's way: an "EXPORT LIST" button that is connected to nothing - clicking it does nothing (`B:gui/views/validation_view.py:186-200`; RAN B audit: no action connected); B's rewrite `a62d3cc` deleted the code behind the old button, while its message says "All functionality has been reimplemented".
Evidence: A writes the file in an older Windows character set (the Windows code page), so some names may not survive, and opening it works on Windows only - elsewhere "Export Complete" is followed by "Export Error" (A-G31). Tier 3 (nice-to-have) in both audits' priority lists. MERGE_COMPARISON.md section 3, row 12.
Recommendation: Keep A's export, written in UTF-8 (the standard text encoding that holds any character), because it is small, works in A, and both screens already show the button.
Your answer: ____

### P3-11. Hand-written mouse-wheel rules                                  Part 3 · refs A-R146, A-G68
What this is about: on a long page, such as Step 4's list of 150 clients in the GSTR-9 season, the mouse wheel should scroll. In Tkinter (A's toolkit) the wheel does not scroll such pages by itself, so A wrote its own rules. Qt-based toolkits - PyQt5 (B's) and PySide6 (the one your Standard 12 names, Qt 6) - scroll by themselves.
A's way: one rule per scroll area (`A:gui/tabs/setup_tab.py:193-248`; `A:power_query_extractor/gui/extractor_window.py:434-478`), guarded by 64 tests; it broke twice before it held (`322717f` → `d64d3a9`, both 25-07-2026; the shared `4c42758` of 29-10-2025 → `bf2f1dd`, 27-07-2026).
B's way: Qt's own scrolling; nothing extra.
Other choices: if Part 8 keeps Tkinter, carry A's final rule and its tests.
Evidence: gaps remain in A (small wheel steps are dropped in the Step 4 log box - the on-screen list of what the program did; A-G68). MERGE_COMPARISON.md section 3, row 31; section 5, Category 16.1.
Recommendation: Drop, because it exists only to patch Tkinter and the window toolkit your Standard 12 names scrolls by itself.
Your answer: ____

### P3-13. B's run history and "Last Run" column                                  Part 3 · refs B-F17, B-R119, B-R120, A-R123, B-F25 (history part), B-G25, B-G59, B-G64 (part)
What this is about: B saves a record of every run in your user profile (your own folder on the PC) and shows on Step 2 when each client was last processed - for example when <Client A>'s ITC and Sales reports were last built. It also stores every figure Step 4 extracted, for every client.
A's way: no history; the records of a run are the summary workbook, the per-client text reports, the consolidated workbook of Step 4 and the log files (text files where the program records what it did) (A-R123).
B's way: `run_history.json` in `~/.gst_organizer/` (`B:models/run_history.py:16-148`; "Last Run" `B:gui/views/validation_view.py:381-385`; `c258047`, 07-02-2026).
Evidence: RAN (B audit and comparison section 5): "Last Run" appears only after the next scan; client records are never trimmed; a wrongly typed history file makes B abort at the end of the next run; Step 4 results are filed under a name no screen looks up, marked "refreshed" even when the refresh failed; the figures stay for ever outside the client folders (B-G64). A per-run Step 4 log file without the figures is P1-17; where client data may be written is contract C16 (Part 7). MERGE_COMPARISON.md section 3, row 33; section 5, Category 14.7; section 10.3.
Recommendation: Drop, because it keeps every client's figures indefinitely in a file outside the client folders, and the one thing it shows (Last Run) can be read from the dated run folders themselves.
Your answer: ____

### P3-16. Step 4 fills itself when the window opens                                  Part 3 · refs B-R126, B-G58 (part), A-R144 (part)
What this is about: when you open the tool, B's Step 4 already lists the clients and Version folders of the latest batch; A's Step 4 stays empty until you press its Scan button. Example: on a December morning you open the tool only to scan one late GSTR-3B download, and B first walks the whole FY 2025-26 batch on the network share.
A's way: Step 4's folder box follows Step 1's target as typed; nothing is read until you press Scan or a "full pipeline" run hands over to Step 4 (`auto_load=False`, `A:gui/tabs/extract_tab.py:71-80`; `bfcb24a`, 26-07-2026; pinned by the test `A:tests/test_main_window_tabs.py:124`).
B's way: at start-up it reads the last SAVED target folder and walks it on the window's own line of work, so the window waits (`B:gui/pyqt_main_window.py:171-190, 419-504`; `b0e7d0e`, 09-02-2026).
Other choices: keep B's fill but run it in a background worker that shows which folder it read - the share is then still read at every start.
Evidence: a slow or unreachable share delays B's start-up (B-G58, READ); a target typed but not yet saved is ignored (B-G28, RAN). A's `bfcb24a` gives its reason: scanning at start-up "would hit the network on every launch" (A audit 8.4.4). The Step 4 layout and where its folder walk runs are P1-13; which batch counts as "latest" is P1-14. MERGE_COMPARISON.md section 5, Category 16.13.
Recommendation: Drop the start-up fill (A's way: Step 4 reads the share only when asked), because you chose this deliberately in A's latest work and pinned it with a test, so that opening the tool never waits on the network.
Your answer: ____

### P3-19. Fold-away help cards                                  Part 3 · refs B-F33 (collapsible part), A-F31 (collapsible part)
What this is about: help boxes that fold away with one click - for example the card telling staff how GSTR-2B and IMS downloads must be named - so a 1366×768 laptop screen keeps room for the controls.
A's way: one fold-away box on Step 2 (`A:gui/widgets/collapsible_frame.py:7`, used at `A:gui/tabs/validation_tab.py:47`).
B's way: two fold-away cards on Step 1, "INSTRUCTIONS" and "EXPECTED FILE NAMES / MODES", folding at each scan (`B:gui/views/setup_view.py:17-82, 469-517`; `7e09e5c`, 08-02-2026).
Evidence: B's card text is typed by hand and its mode lines contradict the mode buttons ("Re-Run: Overwrite existing"); at the opening window size the Scan button is below the visible part of Step 1 in both (RAN for the comparison, section 6). The wording of the expected-file-names card is a GST rule ⚠ (Part 2); the other help texts are P1-20. MERGE_COMPARISON.md section 3, row 31; section 6.
Recommendation: Keep fold-away help cards, with texts written from the new tool's real behaviour, because they keep help one click away without crowding small screens.
Your answer: ____

### P3-21. One-time import of the old settings file                                  Part 3 · refs B-F2, B-R118
What this is about: the first time B opens on a PC where the Tkinter program was used - say an article assistant's PC that ran A last season - it copies the old settings (the four paths, mode, include-name tick, length, dark mode) into its own settings file, once.
A's way: nothing to import; A's settings file `~/.gst_organizer_cache.json` is exactly the file B imports.
B's way: `B:models/cache_manager.py:93-131`, called at `B:gui/pyqt_main_window.py:47`; marker file `.migrated` (`c258047`, 07-02-2026).
Evidence: RAN for the comparison (section 5, Category 14.3): A's recent-folders list and workflow choice are lost; a saved length of 0 becomes 15; a damaged old file is retried at every start; deleting the marker re-imports the old settings over the current ones (B audit 2.4.1 row 6; Section 3, B-F2 trace, item 5). The new tool keeps its settings in `%LOCALAPPDATA%\<App>\` (Windows' per-user folder for program data, Standard 13). MERGE_COMPARISON.md section 3, row 32; section 5, Category 14.4.
Recommendation: Drop, because typing four paths once per PC carries less risk than an import that gets two values wrong.
Your answer: ____

#### Asked in another part - not repeated here (for the assembler)

One-sided features that a Part 1 or Part 2 item already decides. The first draft of Part 3 had them as separate items (old local ids in brackets); they were merged so that each question is asked once, with one treatment. Where the draft added a point the other item lacks, it is noted for the assembler to fold in. P3-3 (install offer), P3-8 (export) and P3-16 (Step 4 start-up fill) stay in Part 3 because P1-1, P1-8 and P1-13 send those questions here.

- [P3-2] Start-up check of the add-ons (A-F1 check part, A-R124, B-F1, B-R124, A-G7, A-G22 part, B-G32) → **P1-1**, same recommendation (A's check); P1-1 already carries the draft's two points (a damaged pywin32 passes A's check, A-G7; without a console A dies at `A:main.py:17` before any check, A-G22). The install offer stays here as P3-3, which P1-1 names as "also asked in Part 3".
- [P3-6] Per-client "client name in folder names" switch (A-F10, A-R137, A-R40 per-client part, B-F5, B-R36, A-G19 part) → **P1-9**. Different recommendation: the draft recommended keeping a per-client switch that the table shows exactly as the run uses it (B lost it only as a side effect of `a62d3cc`); P1-9 recommends the global tick only and offers "Defer" as another choice - keep one recommendation, P1-9's, and the draft's view may go under its "Other choices". Also to settle: MERGE_COMPARISON.md section 4, 4.4.4 tags this setting [GST RULE] (the folder text it changes is written into the Links cells, like the category folder names of P2-7), while section 5, 6.3 and both registries (A-R40, B-R35, B-R36) say it is not; when in doubt it is a GST rule, which would mean "No recommendation - your call" in P1-9.
- [P3-9] Show the files the scan could not use (A-F9, B-F9, A-R61 Variations part, B-R95, A-G67, B-G20) → **P2-2** (treated as a GST rule there: no recommendation) and **P1-7** (an on-screen list of rejected files with reasons). The draft's recommendation for A's list is withdrawn so that the item has one treatment; P2-2 already states the technical point that B's sheet says "No file variations found" when there are some (RAN).
- [P3-10] Hover help (A-R147, A-F31 tooltip part, B-F33 tooltip part) → **P1-20**, same recommendation (A's hover tips, texts rewritten). Point to fold in: A's tips are not kept inside the screen near its edges (A audit 11.4.4, S3b-67).
- [P3-12] A's speed-ups on the network share (A-R110, A-R111) → **P1-6** (signatures read 16 at a time; same recommendation, A's reading) and **P1-11** (run folder worked out once per run; same recommendation, A).
- [P3-14] A text log of each Step 4 run next to its workbook (B-F25 log part, B-R122, A-F25 part, B-G64 part) → **P1-17**, same recommendation (written as the run goes, reasons but no figures).
- [P3-15] Stop button for Step 4 (B-F24 Stop part, B-R112 Step 4 part, A-G50 part, B-G70 part) → **P1-18**, same recommendation (Stop in Steps 3 and 4). Points to fold in: Stop checked inside every wait (B audit 11.4.3, T12), and a stopped run reported as "stopped", not as "Extraction complete" with a green tab dot (B-G70).
- [P3-17] Live scan log on Step 1 (B-F7, B-G22 part, B-G56 part) → **P1-6**, same recommendation (a live count and percentage, not a line per file), and **P1-7** (B's counts that do not add up).
- [P3-18] Coloured dots on the tabs (B-F29 dots part, B-R140 dots part, B-G71 part) → **P1-2**, same recommendation (no dots: they turn green after failures, RAN).
- [P3-20] One theme file for all colours (B-R144 and A-R148 approach, B-G67, B-G69 part, A-G64 part) → **P1-21**, same recommendation (B's one colour sheet, no colours fixed inside single screens). Whether dark mode is built in the first build at all is P5-1.
- Cross-part pointer: P8-8 sends the separate `launch_extractor.py` start files to "Part 4 ('The separate launchers')"; the launcher question itself is P1-22, and Part 4 keeps only the old entry file `extractor_main.py` (P4-7).

## Part 4 - Candidates to drop

Dead code (code that exists but is never used by the running program), unused settings, unfinished or abandoned approaches and superseded files, from both audits' Section 2.8 (dead code) and Section 12E (abandoned approaches), grouped so that each item can be answered at once. Nothing is copied from A or B into the new tool - the build writes everything new - so "Drop" means: the new tool will not have it and the Brief will not ask for it. Where a dead piece holds an idea worth having, the idea is offered separately in Part 5.

### P4-1. Settings and lists that nothing ever reads                                  Part 4 · refs A-R13, A-R14, A-R15, A-R37, A-R44, A-R53, A-R70, A-R84, A-R93, A-R107, A-R130, A-R150, A-R151, A-R153, B-R14, B-R15, B-R16, B-R33, B-R45, B-R49, B-R64, B-R85, B-R92, B-R106, B-R107, B-R108, B-R109, B-R130, B-R131
What this is about: values typed into the code that no part of the program reads. Example: the "safety" settings promise 3 retries and a 10 KB minimum size for each client's ITC and Sales report, but no report is ever retried or size-checked.
A's way: in `A:utils/constants.py` - `EXCEL_EXTENSIONS` (:196), `TEMPLATE_EXTENSIONS` (:197), `MAX_FILE_SIZE_MB = 100` (:198), `EXCEL_SIGNATURES` (:199-202), `EXCEL_SAFETY` (:234-245), `FOLDER_STRUCTURE['level2']` (:83), the mode icons (:178, 183, 188), the fonts (:138-144), the colour "white" (:153), nine unused message texts (:214-227); in the extraction settings `auto_refresh_if_missing`, `show_timestamp` and `FIELD_DISPLAY_NAMES` (`A:power_query_extractor/config/cell_mappings.py:250-251, 261`); unused version labels (`A:core/__init__.py:12-13` and three more).
B's way: the same constants at `B:utils/constants.py:179-185, 217-228, 83`; the mode names, texts and icons (:157-173); the whole window table `GUI_CONFIG` (:126-151) and all message texts `MESSAGES` (:191-211), which died with the Tkinter screens in `a62d3cc`; the same extraction settings (identical file) and version labels.
Evidence: A's are never read at any commit since the first one, `edac6a3` (31-08-2025) (A audit 12E.5); B's are the same, except `GUI_CONFIG`, `MESSAGES` and the mode names and texts, which the Tkinter screens read until `a62d3cc` (07-02-2026) deleted those screens (B audit 2.8.5-2.8.6, 12E.5). Dropping them changes nothing the tool does: the template types and extraction cells really used are your rules in P2-9 and Part 6. `TEMPLATE_EXTENSIONS` is a third list that disagrees with both template pickers; A's own March 2026 audit even reported part of `EXCEL_SAFETY` as used, which was never true (A audit 2.9 #15). MERGE_COMPARISON.md section 8.3.
Recommendation: Drop all, because nothing reads them and their values contradict what the code really does, which misleads anyone - or any AI - reading the code later.
Your answer: ____

### P4-2. Helper functions that are never called                                  Part 4 · refs A-R18, A-R34, A-R49, A-R50, A-R63, B-R18, B-R30 (get_date_only part), B-R41, B-R46, B-R47, B-R58
What this is about: small functions written but never used: unique names instead of overwriting (`get_unique_path`), a file fingerprint that could prove a copied GSTR-2B file matches its source (`calculate_file_hash`), relative paths for the Links cells (`get_relative_excel_path`), today's date (`get_date_only`), deleting a Version folder left empty after a client fails (`cleanup_on_error`), a run count (`get_summary`), and a client-name check whose result nobody reads.
A's way: `A:utils/helpers.py:29-31, 451-470, 529-547, 557-574`; `A:core/file_organizer.py:381-426`; name check stored, never shown (`A:core/file_parser.py:108-119`); still only in A (B's rewrite deleted them with the Tkinter screens): three window shortcuts, `UIHelpers.update_status` and two settings helpers (A audit 2.8.3 #5-#10); also an unused progress callback (2.8.4 #14).
B's way: the same functions at `B:utils/helpers.py:27-29, 344-363, 422-440, 450-467`; `B:core/file_organizer.py:364-409`; the same unread name check (`B:core/file_parser.py:104-115`) and progress callback.
Evidence: none has a caller at `edac6a3`, at the split `47a1492`, or today (A audit 2.8.13, RAN; B audit 2.8.3); nine of the shared names were re-checked in both (MERGE_COMPARISON.md section 8.3).
Recommendation: Drop, because none has ever run; the ideas worth having (a content check of copies, clean-up after a failure) are offered as new work in Part 5.
Your answer: ____

### P4-3. Dead or unreachable code in the Excel and Step 4 parts                                  Part 4 · refs A-R92, A-R94, A-R95, B-R72, B-R91, B-R93, A-G74, B-G79
What this is about: Excel and consolidation code that looks like protection or a feature but can never act (a "branch" is a piece of code that runs only in one case; "unreachable" means no case ever gets there). Example: a comment says "Don't quit if Excel was already open", but the test behind it is always true, so the Excel in which you have <Client A>'s working open is quit every time.
A's way: the "don't quit" test and the clean-up on discard (`A:core/excel_handler.py:58-64, 241-246`); a never-reachable "Open method of Workbooks class failed" branch and an unreachable `return` (`A:power_query_extractor/core/report_processor.py:169-181, 383-386`); a `cleanup()` that does nothing (:801-804); two older consolidated layouts, never called (`A:power_query_extractor/core/data_consolidator.py:250-352, 354-414`); description lines no setting uses (:52-65); a doubled header block (:73-83, A only).
B's way: the same at `B:core/excel_handler.py:57-63, 206-211`; `B:power_query_extractor/core/report_processor.py:122-137, 749-752`; `B:power_query_extractor/core/data_consolidator.py:51-65, 237-339, 341-399` (B removed the unreachable `return` and the doubled header in `c258047`).
Evidence: all READ in both audits (A audit 2.8.3-2.8.4; B audit 2.8.3-2.8.4); all from `edac6a3` (shared). Leaving your own Excel open comes from your Standard 12's private Excel (P8-6), not from this code; the older layout `_write_section` is the only code that would put a failure reason into the workbook - that idea is P5-3. MERGE_COMPARISON.md section 8.3.
Recommendation: Drop, because none of it can run, and what it pretends to do is provided properly elsewhere (Standard 12, P5-3).
Your answer: ____

### P4-4. Never-used code that holds a GST rule                                  Part 4 · refs A-R30, B-R28, A-R85, B-R84, A-R16 (category and folder labels), B-R13 (category and folder labels)
What this is about: three pieces of code that hold GST rules but never run. (1) A check that would flag a client lacking a "critical" return type ⚠ - "Annual Report" or "GSTR-3B Export" - as "Missing Critical Files". (2) "Calculated columns" in the consolidated workbook - subtract, add, multiply or divide two cells (a blank counts as 0, division by zero gives "Div/0"). (3) A "category" and "folder" label on each of the six file-name patterns ⚠ (the Annual Report's label "root" disagrees with where those files really go, the Version folder).
A's way: (1) `A:core/file_parser.py:338-371` (list :359); (2) `A:power_query_extractor/core/data_consolidator.py:130-136, 188-218`; (3) `A:utils/constants.py:17-18, 25-26, 33-34, 41-42, 49-50, 57-58`.
B's way: (1) `B:core/file_parser.py:366-399` (list :387); (2) `B:power_query_extractor/core/data_consolidator.py:118-124, 176-206`; (3) the same lines in `B:utils/constants.py`.
Evidence: technical point only - (1) is never called, (2) is used by none of the 22 extraction settings (all have no calculation), (3) is copied into every file record and never read; all three since `edac6a3` (31-08-2025, shared), unchanged in both. Whether any of them should become a rule of the new tool is your call. MERGE_COMPARISON.md section 8.3.
Recommendation: No recommendation - your call
Your answer: ____

### P4-5. A's own leftovers                                  Part 4 · refs A-R129, A-R117
What this is about: pieces A's July 2026 work left behind, plus older clutter A kept. Example: a Step 4 notice "Power Query refresh is unavailable ... needs pywin32" that can no longer appear when pywin32 is missing, and would blame pywin32 for any other error - say a damaged file on the day you refresh <Client A>'s reports.
A's way: the notice (`A:gui/tabs/extract_tab.py:21-30, 67-69, 84-105`; `c463256`); helpers kept only for the tests (`find_excel_files`, `validate_excel_file`, `_default_exists`, and a test-only shortcut in the standalone window; A audit 2.8.11); unused leftovers - an unused copy of the default refreshed-file suffix, a stored colour never read, two colour keys never used (A audit 2.8.5-2.8.7); the recent-folders list written at each save and never read (`A:gui/utils/cache_manager.py:32-43`, since `edac6a3`); comments pasted from AI chat answers such as "# ADD THIS LINE" (A audit 12E.5).
B's way: none of these: B's rewrite `c258047` replaced these screens and removed most pasted comments (B's own leftovers are P4-6).
Evidence: the notice cannot appear for a missing pywin32 (RAN A audit 2.8.4 #6); the helpers above are reached only by A's tests (A audit 2.8.11). Two more pieces of A's clutter are decided elsewhere: the 14 "[DEBUG]" lines on the Step 3 screen, kept by `064c634` as "a separate UX decision" (P1-12), and the empty `logs\` and `temp\` folders A creates in its program folder at every start (P1-1 and P8-8). A audit 2.8, 12E.4-12E.5.
Recommendation: Drop, because each is unreachable, kept only for tests, or clutter that nothing uses.
Your answer: ____

### P4-6. B's own leftovers                                  Part 4 · refs B-R105, B-R117, B-R132, B-R143
What this is about: pieces B's February 2026 rewrite created or left behind and never used - for example a file that records the season's recent source folders but is never read back. Another example: `models/signals.py`, a "central signal hub" (a signal is Qt's message from one part of the program to another) whose own description says all communication goes through it - nothing ever loads it.
A's way: none of these (A has no dead file or class, A audit 2.8.1-2.8.2).
B's way: `B:models/signals.py` (18 unused signals, `c258047`); an unused fold-away widget `B:gui/widgets/collapsible_group.py`; the `B:power_query_extractor/gui/` folder, which holds only a one-line package marker since `a62d3cc` deleted the window inside it; unused helpers on the scan and processing records (`B:models/app_state.py:57-88`), `ProcessingView.reset`, `ThemeManager.toggle`, a colour table, an `#info` button style and a grey tab-dot colour (B-R132); `session_state.json` written at every close and never read (`B:gui/pyqt_main_window.py:604-607`; readers `B:models/cache_manager.py:82-89` never called); helpers dead since `7e09e5c` and `a62d3cc` (`get_file_info`, `validate_excel_file`, `truncate_path`); an empty `%TEMP%\gst_excel_temp` folder (in Windows' folder for temporary files) made at every run and never used (`B:core/excel_handler.py:52-53`).
Evidence: RAN (B audit 2.8): never loaded, never created or never read; B's own docstring and commit message describe several of them as working ("Signal/slot architecture replaces 305 self.app.xxx callback references", `c258047`; B audit 2.9 #10, #16). Whether to build "reopen where you left off" properly is P5-2. MERGE_COMPARISON.md section 8.3.
Recommendation: Drop, because nothing uses them and they describe a design B never had, which misleads anyone reading the code.
Your answer: ____

### P4-7. The old extractor entry file `extractor_main.py`                                  Part 4 · refs B-F31, B-R128, B-G45, A-G32 (part), A-R118 (part), A-R121 (part)
What this is about: up to the split (30-10-2025) the Step 4 extractor was a separate program, and its own start file, `power_query_extractor/extractor_main.py`, is still in both projects. Whether the new tool keeps a second way to open Step 4 on its own (`launch_extractor.py`) is P1-22; this item is only the old file. Example: someone types `python -m power_query_extractor.extractor_main` (a command that starts one part of a program from the console) to refresh last week's ITC reports.
A's way: A's `launch_extractor.py` uses it, and it can also be started alone, but only with `python -m` - started directly it fails (A-G32). When your home settings file is missing it takes the target folder from any of three files in the folder it was started from (`A:power_query_extractor/extractor_main.py:36-63`, A-R118), and it writes its own `pq_extractor.log` there (A-R121).
B's way: `B:power_query_extractor/extractor_main.py` cannot even be loaded since `a62d3cc` deleted the window it imports; nothing calls it, and its log and look-up files are never created (B-F31, B-R128, B-G45).
Evidence: RAN (A audit 2.8.1 #2: direct start fails, `-m` works; B audit B-G45: `ModuleNotFoundError`). A stray settings file in the start folder can decide which folder A's version scans (A audit 11.4.4, S3b-60). Under Standard 13 the tool is started only through its `Run [App Name].bat` launcher (P8-8), whatever P1-22 decides. MERGE_COMPARISON.md section 3, rows 29-30; section 8.3.
Recommendation: Drop it in both, because B's copy cannot run at all and A's is a console-only extra way into Step 4 that reads its settings from stray files.
Your answer: ____

### P4-8. The old documents inside A, and the shared README                                  Part 4 · refs MERGE_COMPARISON.md section 10, D3, D4, D7, D9 (no feature IDs)
What this is about: four documents in A and one in both describe an older state of the program: A's March 2026 audit `PROJECT_AUDIT_ANNUALSTATEMENT.md`, the design note `docs/superpowers/specs/2026-07-25-merge-extractor-design.md`, `tests/README.md`, the AI coding tool's setting `.claude/settings.local.json`, and `README.md`.
A's way: the March audit (`fc279af` 04-03-2026, rewritten `f7eb8f1` 05-03-2026, never updated since); the design note (`b276542`, 25-07-2026); `tests/README.md` lists 4 of 18 test files; `.claude/settings.local.json` still tracked.
B's way: only `README.md` (byte-identical to A's, from 29-10-2025, still describing two separate programs); `.claude/settings.local.json` removed in `c258047`.
Evidence: the March audit uses a company-like name in its worked examples (lines 130-232), treated as a real client until you say otherwise (P10-7 asks you), and holds your own folder path - it must never be copied; the new audits replace it and record where every old document disagrees with the code (A audit Appendix X4; B audit 2.9). The design note's two unbuilt plans are the "Continue to extraction" button (P5-6) and moving the extractor's log set-up out of `extractor_main.py` (side-stepped in A since `bfcb24a`, A audit 12E.4, and moot once that file goes, P4-7). MERGE_COMPARISON.md section 10 (D3, D4, D7, D9).
Recommendation: Drop all five from the new project, because the new audits replace them, your pipeline writes a new README later (Step G), and the March audit must stay out because of the client-like name.
Your answer: ____

## Part 5 - Candidates to defer

Improvements that the audits' blueprints (Section 11 of each: priority tiers, known improvements, suggested skills) or their lists of unfinished features (Section 12E.4), and A's design note, suggest - but which a first working build does not need. "Defer" puts the item in the Brief's Parking Lot with today's date; nothing is lost. Fixes for faults that hit real work (Excel safety, Resume, path lengths, the 9-column mismatch and the like) are not here: they belong to the contracts of Part 7 and the tech stack of Part 8.

### P5-1. Dark mode                                  Part 5 · refs A-F28, B-F28, A-R148, B-R144, A-G64, B-G69
What this is about: a switch that turns the whole window dark, for example for late-evening work before the GSTR-9 due date. Both projects have one; neither works fully. How it would be built is P1-21; this item is only whether the first build needs it.
A's way: a tick in the title strip that records each widget's colours and replays them (`A:gui/utils/dark_mode_manager.py:16-294`; rebuilt and patched three times on 26-07-2026; 11 tests).
B's way: View > Dark Mode applies one dark style sheet - a single block of colour and font settings for the whole window (`B:gui/theme/theme_manager.py:25-424`).
Other choices: build it now, in the way chosen in P1-21.
Evidence: RAN (both audits): in A a scan in dark mode gives white text on pale rows and black log lines on dark grey; in B Steps 1-2 stay light and status lines turn dark-on-near-black (A-G64 / B-G69). Both audits rate it Tier 3, the lowest (A audit 11.3: "appearance only"; B audit 11.3). MERGE_COMPARISON.md section 3, row 28.
Recommendation: Defer, because it is appearance only and, with the one colour sheet P1-21 recommends, it can be added later without touching any screen's logic.
Your answer: ____

### P5-2. Reopen where you left off                                  Part 5 · refs A-R117, B-R117, B-F27 (session part)
What this is about: opening on the last tab and window size, and offering recent source folders - e.g. a made-up `Inputs-FY2526` folder - from a list. Both projects save such a list that nothing reads.
A's way: a recent-folders list written at each save, never read (`A:gui/utils/cache_manager.py:32-43`, since `edac6a3`).
B's way: `session_state.json` with recent folders, last tab and window size, written on close and never read; the window always opens on Step 1 at 1400×900 (`B:gui/pyqt_main_window.py:604-607`; `B:models/cache_manager.py:82-89`; `c258047`).
Evidence: half-built in both (A audit 12E.4; B audit 12E.4); the four paths themselves are already remembered by both settings files (P1-4). The unread lists themselves are dropped in P4-5 and P4-6; this item is whether to build the idea properly later. MERGE_COMPARISON.md section 5, Category 14.3.
Recommendation: Defer, because the four paths are already remembered, so this saves little in a first build.
Your answer: ____

### P5-3. Failure reason inside the consolidated workbook                                  Part 5 · refs A-R94, B-R93 (idea)
What this is about: when a client's refresh fails, the consolidated workbook today shows only ✗ for that client; a never-used older layout would have written the reason ("Error Details") next to it, so you see why without opening a log.
A's way: the reason is shown on screen per report ("⚠️ ITC Report - <reason>", `A:power_query_extractor/gui/extractor_window.py:917-938`); the reason column exists only in the never-called layout (`A:power_query_extractor/core/data_consolidator.py:250-352`).
B's way: the screen shows only "Result: ITC: FAIL" (`B:gui/controllers/extraction_controller.py:76-84`); the same never-called layout (`B:power_query_extractor/core/data_consolidator.py:237-339`).
Evidence: the old code is dropped in P4-3, so this is new work. Whether the figures of a failed refresh show as "Missing" is a GST rule (Part 6) and is not touched here. MERGE_COMPARISON.md section 6.
Recommendation: Defer, because the reason already shows on screen (P1-13) and, if you agree to P1-17, in the run folder's log file.
Your answer: ____

### P5-4. Content check of copied files                                  Part 5 · refs A-R50, B-R47 (idea)
What this is about: after copying a client's GSTR-2B file into its folder, both projects check only that the copy has the same size. A never-used helper would compare an md5 fingerprint (a short code computed from every byte of a file, which changes if any byte changes).
A's way: `calculate_file_hash`, never called (`A:utils/helpers.py:529-547`); copies are checked by size (A-R47).
B's way: the same (`B:utils/helpers.py:422-440`; size check B-R42).
Evidence: never called at any commit; the copy problem actually found - a half-written copy left under its final name (A-G11 / B-G26) - is cured by copying to a temporary name and renaming, not by a fingerprint; a fingerprint reads every file a second time over the share (HIGH). MERGE_COMPARISON.md section 8.5.
Recommendation: Defer, because the size check plus a temporary name covers the failures seen so far without reading every file twice.
Your answer: ____

### P5-5. Clean-up after a client fails                                  Part 5 · refs B-R41 (idea; A's same function has no rule ID)
What this is about: if processing of <Client A> fails halfway, its new Version folder and any copies stay on disk. A never-called function would, in Fresh mode (the processing mode that starts a new batch), delete that Version folder - but only when no file other than names starting with "_" is in it, so it would remove empty leftovers, not half-done copies (`A:core/file_organizer.py:398-407`).
A's way: `cleanup_on_error`, never called (`A:core/file_organizer.py:381-407`).
B's way: the same (`B:core/file_organizer.py:364-390`).
Evidence: never called since `edac6a3` in either project; automatic deletion on a network share is risky to get wrong (a wrong folder deleted is worse than one left over). A audit 12E.4; B audit 12E.4.
Recommendation: Defer, because a clear "failed" entry for the client in the run's summary makes leftovers easy to spot, while deleting on the share needs care.
Your answer: ____

### P5-6. A "Continue to extraction" button after an "organise only" run                                  Part 5 · refs A-F5 (part)
What this is about: A's design note of 25-07-2026 planned, for "organise only" runs, a button at the end of Step 3 that takes you on to Step 4 when you decide to go on.
A's way: not built; a log line instead - "Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query." (`A:gui/handlers/processing_handler.py:146-148`); planned at `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:78-81` (`b276542`).
B's way: no choice; always moves to Step 4.
Evidence: only the design note mentions it (RAN, `git log -S`: only `b276542`); it matters only if you keep A's workflow choice (P3-5).
Recommendation: Defer, because clicking the Step 4 tab does the same today.
Your answer: ____

### P5-7. Start the refresh automatically after Step 3                                  Part 5 · refs A-F5 (text part), A-G66 (part)
What this is about: a true "full pipeline" in which Step 4 starts refreshing every report by itself after Step 3 - for example overnight for 50 clients.
A's way: "full pipeline" opens Step 4 and lists the clients but never starts the refresh, although its label says "Also refresh Power Query and extract the data" (`A:utils/constants.py:170`; `A:gui/handlers/processing_handler.py:136-175`).
B's way: moves to Step 4; never starts the refresh.
Evidence: today's refresh presses keys (Alt+A, R, A - English Excel's shortcuts ⚠) into whatever window is in front and trusts fixed waits (A-G55 / B-G62, A-G3 / B-G6); how the new tool refreshes is still to be settled (P8-7) and tested on an office PC; whether the choice between "organise only" and "full pipeline" stays at all is P3-5. MERGE_COMPARISON.md section 6; section 9 (H3, H20).
Recommendation: Defer, because an unattended refresh is safe only once the refresh method has been proven on an office PC.
Your answer: ____

### P5-8. A settings screen for the GST configuration                                  Part 5 · refs none (idea from B audit 11.4.3, T10)
What this is about: the file-name patterns, the Links cells and the 22 extraction cells live in Python files that you edit with AI help, guided by a "HOW TO EDIT THIS FILE" text. A screen could let you change them without touching code - for example when the portal renames a download or a template's GSTR-9 table moves to another cell ⚠.
A's way: `A:utils/constants.py` (patterns, Links cells) and `A:power_query_extractor/config/cell_mappings.py:1-27` (editing instructions).
B's way: the same two files (`cell_mappings.py` byte-identical).
Evidence: the B audit's blueprint keeps the files and adds a check when they load (B audit 11.4.3, T10); such a check also catches the setting that can silently overwrite another (A-G78 / B-G78). The values themselves are your GST rules ⚠ (Part 6) and are not changed by this item.
Recommendation: Defer, because you change these rarely, and a check when the file loads gives most of the safety a screen would.
Your answer: ____

### P5-9. Read the refreshed figures without a second Excel                                  Part 5 · refs A-F23 (reading method), B-F22 (reading method)
What this is about: after refreshing a client's ITC report, both projects open the saved copy again in a second, hidden Excel just to read 22 cells ⚠ (such as the one for the consolidated sheet "Table 8D Limits"). Your Standard 12 allows reading such a workbook with openpyxl instead, without starting Excel (reading is safe; only saving with openpyxl loses Power Query).
A's way: a second Excel per report (`A:power_query_extractor/core/report_processor.py:527-636`).
B's way: the same (`B:power_query_extractor/core/report_processor.py:478-587`).
Evidence: Standard 12: "Reading a Power Query workbook with openpyxl is fine" (`rebuild_pipeline_FINAL.md:827`); both blueprints advise comparing all 22 cells both ways on an office PC before the second Excel goes, because why openpyxl was not used is not recorded (A audit 11.4.3 T5, which also offers re-opening the copy in the same private Excel; B audit 11.4.3 T5). Which cells are read is your GST rule ⚠ (Part 6) and does not change. MERGE_COMPARISON.md section 7.
Recommendation: Defer, because the first build can read through Excel as both projects do, and switch only after the office-PC comparison shows identical figures.
Your answer: ____

### P5-10. Skill files written from the lessons of A and B                                  Part 5 · refs none (A audit 11.6, B audit 11.6)
What this is about: both audits suggest 13 and 14 "skills" - instruction files (SKILL.md) that Claude loads for one kind of task, such as "Excel through COM, one safe session" (COM is Windows' way for one program to drive another) or "one naming rule, one finder" - so later projects do not relearn these lessons.
A's way: A audit 11.6, SK1 to SK13.
B's way: B audit 11.6, SK1 to SK14.
Evidence: the Excel and start-up rules several of them repeat are already in your CLAUDE.md (Standards 12 and 13); a skill is most useful when written from what a build has proved on the office PCs.
Recommendation: Defer, because the skills are best written after the first working build, from what it proved.
Your answer: ____

#### For the closing list

None found from Parts 3 to 5. Checked: every map row these parts drew on is "only A", "only B", or dead or half-built in both (MAP.md sections 1 and 2; MERGE_COMPARISON.md sections 3.3, 5 Categories 13-16 and 8.3); a row that is dead in both is a drop candidate (Part 4), not "the same and working". Also checked against the decision candidates of every comparison writer (DECISION_CANDIDATES.md) and against the drafts of Parts 1, 2, 6 and 7 to 10: each one-sided candidate is either an item here or listed under "Asked in another part".

#### Refs covered

By the items of this file:

Features: A-F1 (offer part), A-F4, A-F5, A-F12, A-F23 (reading method), A-F28, A-F31 (collapsible part); B-F2, B-F10, B-F17, B-F22 (reading method), B-F25 (history part), B-F27 (session part), B-F28, B-F31, B-F33 (collapsible part).

Rules: A-R13, A-R14, A-R15, A-R16 (category and folder labels), A-R18, A-R30, A-R34, A-R37, A-R44, A-R49, A-R50, A-R53, A-R63, A-R70, A-R84, A-R85, A-R92, A-R93, A-R94, A-R95, A-R107, A-R113, A-R117, A-R118 (part), A-R121 (part), A-R123, A-R125, A-R129, A-R130, A-R133, A-R136, A-R142, A-R144 (part), A-R146, A-R148, A-R150, A-R151, A-R153, A-R154; B-R13 (category and folder labels), B-R14, B-R15, B-R16, B-R18, B-R28, B-R30 (get_date_only part), B-R33, B-R41, B-R45, B-R46, B-R47, B-R49, B-R58, B-R64, B-R72, B-R84, B-R85, B-R91, B-R92, B-R93, B-R105, B-R106, B-R107, B-R108, B-R109, B-R117, B-R118, B-R119, B-R120, B-R126, B-R128, B-R130, B-R131, B-R132, B-R140 (grey dot part), B-R141 (note part), B-R143, B-R144.

Gaps: A-G20 (part), A-G31, A-G32 (part), A-G49, A-G53, A-G59, A-G60, A-G64, A-G66 (part), A-G68, A-G74, A-G75 (part); B-G25, B-G45, B-G58 (part), B-G59, B-G64 (part), B-G69, B-G73 (part), B-G79.

Other material: A's `tests/` folder (P3-1); A's `PROJECT_AUDIT_ANNUALSTATEMENT.md`, `docs/superpowers/specs/2026-07-25-merge-extractor-design.md`, `tests/README.md`, `.claude/settings.local.json` and the shared `README.md` (P4-8; MERGE_COMPARISON.md section 10, D3, D4, D7, D9); both audits' Section 11.6 skill ideas (P5-10).

Asked in another part (see the list at the end of Part 3; each is in that item's own refs):

Features: A-F1 (check part, P1-1), A-F9 (P2-2), A-F10 (P1-9), A-F25 (log part, P1-17), A-F29 (P1-22), A-F31 (tooltip part, P1-20); B-F1 (P1-1), B-F5 (P1-9), B-F7 (P1-6), B-F9 (P2-2), B-F24 (Stop part, P1-18), B-F25 (log part, P1-17), B-F29 (P1-2), B-F30 (P1-22), B-F33 (tooltip part, P1-20).

Rules: A-R40 (per-client part, P1-9), A-R61 (Variations part, P2-2), A-R110 (P1-6), A-R111 (P1-11), A-R124 (P1-1), A-R128 (logs/temp part, P1-1 and P8-8), A-R137 (P1-9), A-R147 (P1-20); B-R36 (P1-9), B-R95 (P2-2), B-R112 (Step 4 part, P1-18), B-R122 (P1-17), B-R124 (P1-1), B-R127 (P1-22), B-R140 (dots, P1-2).

Gaps: A-G7, A-G22 (P1-1); A-G19 (part, P1-9); A-G50 (part, P1-18); A-G67 (P2-2); A-G69 (part, P1-12); B-G20 (P2-2); B-G22 (part, P1-7); B-G32 (P1-1); B-G44 (P1-22); B-G56 (part, P1-6); B-G67 (P1-3); B-G70 (part, P1-12); B-G71 (part, P1-13).
