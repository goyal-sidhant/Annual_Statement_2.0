# Decision candidates and flags from the comparison writers

Collected from the MAP, C3, C4, C5a, C5b, C6, C78 and C910 writers (23-10-2026). Candidates overlap; the Phase 3 drafters merge duplicates. Checkers corrected the sections afterwards; where a candidate and a checked section disagree, the section wins.

## MAP

### Decision candidates
- IMS Reco duplicate '(N)' names: A accepts them and links the '(1)' copy first; B rejects them. Refs A-R2/B-R7, A-R57/B-R56, A-R27/B-R25, A-G6/B-G3, A-G38/B-G47, B-G48, A commit 5c8ce95. [GST RULE] ⚠
- Category folder names: A uses GSTR3B / ITC / Sales; B uses 'GSTR-3B Exports' / 'Other ITC related files' / 'Sales related files'. The names change the folder text written into the Links cells and depend on whether the real templates read that text. Refs A-R39/B-R35, A-R41/B-R37, A-R54/B-R54, A-R55/B-R55, A-G41/B-G54, A commit 7347383. [GST RULE] ⚠ (A-R41/A-R54/A-R55 are tagged; the naming itself is technical)
- Which cells feed the IGST/CGST/SGST columns of the consolidated sheets 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B'. Both projects show Diff GSTR-3B J15/K15/L15 there. Refs A-G1/B-G1, A-R80..A-R82/B-R79..B-R81, A-R88/B-R83. [GST RULE] ⚠
- Refresh success test Info!BB2 (same in both) and how a refresh is confirmed before saving. Refs A-R73/B-R70/B-R71, A-R71/B-R66/B-R67, A-G3/B-G6, Standard 12. [GST RULE] ⚠ for the BB2 cell; the refresh method is technical
- Whether figures read from a report whose refresh failed are shown or replaced by 'Missing' (same in both). Refs A-R89/B-R88, A-G17/B-G10. [GST RULE]
- Links cells for file types a client lacks keep the template's text, and only the first file of a type is linked (same in both). Refs A-R54..A-R57/B-R54..B-R56, A-G5/B-G2, A-G6/B-G3, B-G29. [GST RULE] ⚠
- Both reports built for every client whatever files it has; a client counted 'Success' even when its reports fail (same in both). Refs A-R60/B-R50/B-R60, A-G19/B-G19. [GST RULE] for which reports are made; the success counting is technical
- Expected six file types, duplicate rule and completeness (same in both; confirm for 2.0). Refs A-R26..A-R28/B-R24..B-R26, A-R30/B-R28. [GST RULE] ⚠
- State-code table, fallback codes, hyphen ban in names and spelling variants (same in both). Refs A-R21/A-R22/B-R21/B-R22, A-G33..A-G36/B-G16/B-G17, A-G77. [GST RULE] ⚠
- No period or financial-year check (same in both). Refs A-R25/B-R23, A-G39/B-G49/B-G23. [GST RULE] ⚠
- Dry Run gate: A requires both templates and a target; B requires only a ticked client. Refs A-R51/B-R48/B-R133, A-F11/B-F11. [GST RULE] by A's tag
- Step 1 help card: A builds it from the real patterns; B types it by hand. Refs A-R16/B-R13/B-R141, A-G66/B-G72. [GST RULE] by tag ⚠
- Template paths and options: A re-reads them live during a run; B takes a snapshot at Start. Refs A-R60/B-R50, A-G14/B-G28. Technical (A-R60 carries a GST tag)
- Client key handling on Step 2: A turns number-looking keys into numbers; B stores the key with the row. Refs A-R17/B-R17, A-G25. Technical fix of a GST-tagged rule
- GUI toolkit: Tkinter (A) or PyQt5 (B). Refs A-F2/B-F1 and the whole screen layer; note the B audit's remark that Standard 13's embeddable Python lacks tkinter. Technical
- pywin32 handling: A guards it, checks at start and offers to install (A-R64, A-R124, A-R125, A-G7, A-G53, A-G59); B needs it just to start (B-R124, B-G32). Technical
- Scan acceptance: A drops small or invalid files silently; B lists them. Also the ~$ skip, case-insensitive extension matching and A's parallel signature reads. Refs A-R8..A-R10/B-R1..B-R3, A-R109/A-R110, A-G37/B-G46/B-G22. Technical
- Showing the problem-file list: A only in the summary workbook; B count only, with an empty summary sheet. Refs A-F9/B-F9, A-R12/B-R12, A-G67/B-G20. Technical (the hint text is GST-tagged)
- Private/Limited abbreviation in file names: A whole-word and any case; B case-sensitive substring. Refs A-R19/A-R20/B-R19/B-R20, B-G18. Technical
- Step 3 report creation: A builds locally and copies once; B writes in place and leaves an unfilled copy on failure. Refs A-R98/B-R52/B-R105, B-G5, A-G12/B-G26. Technical
- Step 4 batch and version choice: date modified, newest version (A) or highest name, all versions (B). Refs A-R42/B-R39, A-F19/B-F18, A-G45/B-G12. Technical
- Refreshed-copy recognition and Skip refresh: configured suffix (A) or fixed '_Refreshed' (B), and reuse of copies from failed refreshes in both. Refs A-R68/A-R69/B-R63/B-R136, A-G2/B-G7, A-G16/B-G9. Technical
- Settings storage and migration: A's single file or B's folder with settings, session and history plus a one-time import; Standard 13 location. Refs A-R114..A-R119/B-R115..B-R118, A-G24/B-G36. Technical
- Run history and Extraction_Log (B only), including how long extracted figures are kept in the user profile. Refs B-F17/B-F25, B-R119/B-R120/B-R122, B-G25/B-G59/B-G64. Technical
- Workflow choice (A only) or always switching to Step 4 (B). Refs A-F5/A-R154. Technical
- Live setup checklist (A only). Refs A-F4/A-R113/A-R133, A-G20/A-G49. Technical
- Per-client include-name toggle and long-name warnings (A only; dead in B). Refs A-F10/A-R136/A-R137/B-F5/B-R36. Technical
- Client-list export (A works; B's button is dead). Refs A-F12/B-F10, A-G31/B-G73. Technical
- Threading, Stop and close behaviour. Refs A-F17/B-F16/B-F34, A-G13, A-G27/B-G40, A-G48/B-G58, A-G65/B-G70/B-G74, A-G70. Technical
- Step 4 wait box range: 5-60 with unchecked typed values (A) or 5-120 enforced (B). Refs A-R72/B-R65/B-R68, A-G26. Technical
- Consolidated workbook location and row label: the folder from the last scan with the client folder name (A), or the run folder shown with '<client folder> / <version folder>' (B). Refs A-R86/A-R87/B-R86/B-R87. Technical
- Shared Excel-safety defects to fix in 2.0 whichever base is kept. Refs A-G54/B-G63, A-G55/B-G61/B-G62, A-G43/B-G53, A-G28/B-G38, A-G71/B-G74/B-G82, A-G50/B-G57/B-G39. Technical
- Shared file-safety defects. Refs A-G8/B-G15, B-G14, A-G9/B-G27, A-G10/B-G13, A-G11/B-G26, A-G41/B-G54. Technical (A-G8 is GST-tagged)
- Logging: location in the start folder, client data in logs, cp1252 line loss, no log from B's launcher, all warnings switched off. Refs A-R120..A-R122/B-R121/B-R123/B-R127, A-G56/B-G64/B-G44, A-G61/B-G31. Technical
- Start-up without a console and error visibility. Refs A-R127/B-R125, A-G22/B-G34, A-G30/B-G35. Technical
- Tests: A's 288 tests pin A's behaviours; 7 of A's pure-logic tests fail on B's code (IMS '(N)' and the abbreviation). Decide which behaviours the 2.0 tests should pin. Refs A-R2, A-R19, A-G60. Technical (the IMS item is a GST rule)

### Flags for other sections
- The two registries tag several matching pairs differently. [GST RULE]: A-R48 no vs B-R57 yes; A-R51 yes vs B-R48 no; A-R58 no vs B-R53 yes; A-R60 yes vs B-R60 no; A-R66 no vs B-R61 yes. ⚠: A-R16 no vs B-R13 yes; A-R22 no vs B-R22 yes; A-R41 yes vs B-R37 no; A-R58 no vs B-R53 yes. MAP.md uses 'yes' wherever either registry says yes; the comparison sections should do the same.
- Evidence the comparison sections can reuse: the business core is shared and verified function by function in /root/merge-scratch/run/MAP/funccmp_out.txt (52 identical, 14 differ only in comments, log-file lines or error types, 16 really different, 4 only in A). cell_mappings.py is byte-identical. In utils/constants.py only the IMS Reco pattern, the category folder names and WORKFLOW_MODES differ.
- Worked example for A-G6 / A-R57 and the decision list: in A, an IMS duplicate 'ImsReco-...-31032025 (1).xlsx' is linked into Links!B11 instead of the plain file, because a space sorts before a dot. The client also becomes 'Has duplicates'. In B the duplicate is a variation and the plain file is linked (RAN, P:excel.prepare).
- On Linux, B's wildcard search misses an upper-case .XLSX, while A finds it on every system. Any Linux test of 2.0 should mirror Windows' case-insensitive matching. On Windows both find the file (HIGH, inferred, not run).
- B's one-time settings import (B-R118) reads ~/.gst_organizer_cache.json, which is exactly the file A writes at bf2f1dd (A-R114). If the author ran both programs on one PC, B took A's paths once and ignored workflow_mode and recent_folders. This matters for 2.0's settings migration.
- B-G14 (two source files that clean to the same name) is the same defect in A (RAN with scripts/collide.py; identical in all three modes), but the A audit mentions it only inside A-G8.
- Shared defects that the other audit did not register as gaps: A-G46 (drive letter versus network path) and A-G79 (unverified COM assumptions) also exist in B, which has no separate gap for them. B-G76 (Skip refresh passed through a module-wide setting) and B-G77 (CoUninitialize while Excel objects are still held) also exist in A, which has no gap for them.
- Step 2 resets all ticks at every scan in both programs: A starts every row unticked, B ticks every row. Only the B audit counts this as a gap (B-G73).

### Corrections to the audits or the map
- A audit, Section 1 (PROJECT_AUDIT_AnnualStatement_oldv1.md line 7) and Section 12A (line 10582) say 'A is probably the original repository, renamed (MEDIUM)'. GitHub's creation dates contradict this: B's repository was created on 31-08-2025, three minutes before edac6a3, and A's on 07-02-2026. The same audit's Section 12 (line 11155), the B audit S12own and the task brief all give these dates. B's repository is therefore the original; only the order of the later renames is inferred.
- A registry A-R139 lists A's Dry Run lines but leaves out '🧪 ════════ DRY RUN COMPLETE ════════', which A prints at A:gui/handlers/processing_handler.py:66 (B-R135 correctly includes the same line for B).
- B audit B-G42 (S67) lists the state-code call at B:core/excel_handler.py:550 among the places where an error 'ends the whole processing run'. That call runs inside the worker's per-report try (B:gui/controllers/processing_controller.py:223-239), so an error there fails only that report; the same holds in A (A:gui/handlers/processing_handler.py:304-305). Only the client look-up and the ETA call (B:gui/controllers/processing_controller.py:85-90; A:gui/handlers/processing_handler.py:232-236) sit outside the per-client protection.
- A registry A-R9 says Excel's ~$ lock files are dropped by the 1024-byte limit. That holds for ordinary lock files of about 165 bytes, but a ~$ file of 1024 bytes or more with a ZIP signature is listed by A and becomes a variation (RAN, probe P:scan.variations, '~$Big-lock.xlsx'). B skips every name starting with ~$ (B:utils/helpers.py:259-260).

### Questions for the author
- IMS Reco files saved by a browser as a duplicate, e.g. 'ImsReco-<Client A>-<State>-31032025 (1).xlsx'. Should 2.0 count such a file as an IMS Reco file, as A does, or reject it, as B does? When the plain file and the '(1)' copy both exist, which one should feed Links B11? A currently links the '(1)' copy.
- Do the real ITC and Sales templates read each folder from the Links cells, or do their queries contain a fixed folder name such as 'GSTR-3B Exports' or 'Other ITC related files'? This decides whether A's short folder names work with the templates. Dummy copies of the templates would let this be checked.
- Which cells should fill the IGST/CGST/SGST columns of the consolidated sheets 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B'? Today both versions show the Diff GSTR-3B J15/K15/L15 values there.
- Which version did the office actually use after February 2026: B's v4.0 or A's July 2026 version? Do batches on the office share have long or short category folder names? Step 4 of 2.0 may need to read both.
- Was either program run on a PC where the other had been used? B imports A's settings file (~/.gst_organizer_cache.json) once, and only on its first start.
- Should Dry Run refuse to run until both templates and the target are chosen (A) or only need a ticked client (B)?
- Should Step 4 list every Version folder of a client (B) or only the newest (A)? And should 'newest' mean 'date modified' or the date in the folder name?

## C3

### Decision candidates
- [GST RULE] ⚠ IMS Reco file names with a trailing '(N)' after the date: A accepts them (5c8ce95, 23-07-2026); B rejects them as a variation (edac6a3 pattern). This also decides which file feeds Links B11 and whether the client shows 'Has duplicates'. Refs: A-F7/B-F6, A-R2/B-R7, A-R57/B-R56, A-R27/B-R25, MAP 2.1 item 1.
- [GST RULE] ⚠ Category folder names, which are also written into Links cells B2/B7/B10 (ITC) and B2/B8 (Sales): A has GSTR3B/ITC/Sales (7347383), B has GSTR-3B Exports / Other ITC related files / Sales related files (edac6a3). They depend on what the templates' queries expect. Refs: A-F13/B-F12, A-R41/B-R37, A-R54, A-R55/B-R54, B-R55, MAP 2.1 items 6 and 8.
- Technical: a path-length guard before folders and files are created (A-G41 / B-G54). A's own commit 7347383 calls it 'the durable fix'. Not a GST rule.
- [GST RULE] What must be chosen before Dry Run: A requires both templates and a target (shared edac6a3 gate); B requires only a ticked client (c258047). Refs: A-F11/B-F11, A-R51/B-R48/B-R133, MAP 2.1 item 7.
- [GST RULE] ⚠ A merged Links target cell on the openpyxl path: A skips it silently, B fails the report. Refs: A-F14/B-F13, A-R58/B-R53, MAP 2.1 item 9.
- [GST RULE] How Step 4 recognises ITC/Sales reports: A uses a lower-case prefix test, then a search; B uses a search for both. They differ only for a hand-renamed upper-case .XLSX on Linux. Refs: A-F20/B-F18, A-R59/B-R51/B-R61, MAP 2.1 item 10.
- [GST RULE] Which template builds each client's report when the Step 1 box changes mid-run: A re-reads it per client; B takes a snapshot at Start. Refs: A-F15/B-F14, A-R60/B-R50, MAP 2.1 item 11.
- [GST RULE] ⚠ Content of the 'expected file names' help card: A uses the pattern description strings, B a hand-typed list. Neither mentions .xls, '(N)' or the hyphen ban. Refs: A-F31/B-F33, A-R16/B-R13/B-R141, MAP 2.1 item 3.
- [GST RULE] Client key on the way back from Step 2, and the status display / 'Complete Only' test. A's '0042' crash (A-G25) is technical. Refs: A-F8/B-F8, A-R17/B-R17, A-R29/B-R27, MAP 2.1 items 4-5.
- [GST RULE] ⚠ Hint texts for unrecognised names (identical in both). The technical decision is whether to show the list at all: A shows it in the summary workbook's Variations sheet; B's sheet is always empty and false. Refs: A-F9/B-F9, A-R12/B-R12, A-G67/B-G20.
- [GST RULE] ⚠ 9-column consolidated mismatch: 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' are filled from Diff GSTR-3B!J15/K15/L15. The code is identical in both, and which cell is meant is the author's call. Refs: A-F24/B-F23, A-R88/B-R83, A-G1/B-G1.
- [GST RULE] ⚠ The Info!BB2 success test (identical in both) and the fact that a refresh cannot be confirmed. Refs: A-F21/B-F19/B-F20, A-R73/B-R70, A-G3/B-G6, A-G2/B-G7.
- [GST RULE] A client without Sales-type files gets an unfilled Sales report that is counted as a success, in both. Refs: A-F15/B-F14, A-G5/B-G2.
- Possible [GST RULE] (when in doubt): consolidated rows, one per client (A) or one per ticked version folder (B). Refs: A-F24/B-F23, A-R65/B-R61.
- Technical: window toolkit and job model. A uses Tk, scans on the window thread and runs daemon workers that touch the window. B uses Qt, with QThread workers and signals, but has no guard across job kinds. Refs: A-F2, A-F7, A-F17, A-F21 / B-F1, B-F34; A-G13, A-G48, A-G70 / B-G74.
- Technical: start-up dependency checks. A checks, offers to install and keeps pywin32 optional off Windows; B cannot start without pywin32. Standard 13 (excepthook, .bat/pythonw start). Refs: A-F1/B-F1, A-G22, A-G53, A-G59 / B-G32, B-G34.
- Technical: build reports on a local staging copy (A, 12d505f) or in place (B, which leaves an unfilled template under the report name). Refs: A-F14/B-F13, A-G12/B-G5.
- Technical, Standard 12: both take over, hide and quit the user's own Excel, and both refresh by SendKeys. Refs: A-G54/B-G63, A-G55/B-G61.
- Technical: Step 4 batch and version selection. A uses newest date modified, B the highest day-first name; neither is reliable. Refs: A-F19/B-F18, A-R42/B-R39, A-G45/B-G12.
- Technical: recognise refreshed copies by the configured suffix (A, 2aaa2b1, with tests; over-matches '_R{timestamp}', A-G16) or by fixed text (B). Also whether to store the suffix (B does; A does not, A-R119). Refs: A-F22/B-F21, A-R67 to A-R69.
- Technical: settings file location, when settings are saved, and the one-time import of A's settings file. Refs: A-F27/B-F27/B-F2, A-R114 to A-R119, A-G24/B-G36.
- Technical plus confidentiality: Step 4 Stop button, Step 4 log file and run history. Only B has them, and they store every figure for ever. Refs: A-F25/B-F24, B-F25, B-F17, A-G56/B-G64, B-G25.
- Technical: keep A's live Setup checklist (A-F4) and the workflow choice (A-F5). Fix their misleading texts and the hand-over (A-F5 defects).
- Technical: keep or drop the per-client folder-name toggle (A-F10/B-F5, A-R40/B-R35). A's interplay defects are A-G19.
- Technical: export of the client list (A-F12/B-F10, A-G31/B-G73).
- Technical: the client list's default ticks. A ticks none; B ticks all, incomplete clients too. Refs: A-F8/B-F8.
- Technical: input safety of the client-folder length box (A-F6/B-F4, A-R36/B-R32, A-G26). The Level-2 state code (A-R35) is a GST rule and is identical in both.
- Technical: name cleaning of 'Private'/'Limited'. The map tags A-R19/B-R19 as not GST. A uses whole words in any case; B is case-sensitive anywhere, so its folder and file names disagree.
- Technical: silent drops versus counted variations at the scan, the scan's speed and the upper-case .XLSX miss on Linux. Refs: A-F7/B-F6, A-G37, A-G48 / B-G22, B-G46, B-G55.
- Technical: dark mode approach (A-F28/B-F28, A-G64/B-G69), tooltips and help (A-F31/B-F33, A-G66/B-G72, B-G67 invisible buttons), launcher and log file (A-F29, A-F30/B-F30, B-F31, B-F32, A-G32/B-G44, B-G45).

### Flags for other sections
- Section 5 / C5a Category 6 (folder layout): A's comment at A:utils/constants.py:85-90 presents modelled path-length figures as measured. Commit 7347383 gives the measured run as 270 chars and 1094 of 3030 files over ~218, calls the margin 'thin (215 vs 218)' and says a path-length guard 'remains the durable fix'. Neither project has such a guard (A-G41 / B-G54).
- Section 5 / C5b Category 14, and the Step 3 log: when the summary workbook cannot be written, A's Step 3 log shows '❌ Failed to create summary report' (A:gui/handlers/processing_handler.py:469-473). B's Step 3 log shows nothing, because the IDENTICAL routine returns False instead of raising (B:core/excel_handler.py:313-346) and B logs only success (B:gui/controllers/processing_controller.py:189-193). The error reaches only the program log file.
- Decision list and history sections: B's a62d3cc says 'All functionality has been reimplemented', yet three features were lost in the move to PyQt5 and are kept only in A. They are the client-list export (A-F12 / B-F10, button never connected), the per-client folder-name toggle (A-F10 / B-F5, B:models/app_state.py:83 never written) and the Variations sheet (A-F9 / B-F9, B passes [] at B:gui/controllers/processing_controller.py:73; regression against 47a1492:gui/handlers/processing_handler.py:161).
- B audit Section 7: B's tab dots (B-F29) turn green after an empty scan, a failed run or a stopped run, are never reset, and the Step 2 dot never turns green. The B audit flags these but gives them no gap number of its own.
- Tests sections: A's full suite on Linux gives 1 failed, 287 passed. The failure is test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only (a Windows path read on Linux). Run on B's code, A's pure-logic tests give 7 failed, 66 passed (4 IMS '(N)' tests, 3 name-cleaning tests), and test_file_ops.py cannot be collected ('cannot import name scan_excel_files'). RAN in /root/merge-scratch/run/C3/.
- Any section on the consolidated workbook: A gives one row per client and B one row per ticked version folder ('client / version'). This decides which row a figure lands in, so the decision list should treat it as a possible GST rule (the 'when in doubt' rule). C3 describes it without judging it.
- Settings / Standard 13 sections: B imports exactly the file A writes today (~/.gst_organizer_cache.json; B:models/cache_manager.py:93-131) and ignores A's workflow_mode key. The import loses recent folders through a key mismatch (recent_folders against recent_source_folders) and turns a saved length of 0 into 15.
- UI section: B's Step 1 option reads 'Include client name in brackets for Level 4 folders' and its note promises a warning for names over 10 characters and 'individual client settings' (B:gui/views/setup_view.py:369, 385). B has neither.

### Corrections to the audits or the map
- MAP §1 row A-F13 / B-F12 says only the three category folder names differ. That is incomplete. Copied file names, and the client name in brackets, also differ when 'Private'/'Limited' is written in capitals or sits inside a word. A replaces whole words in any letter case; B replaces case-sensitively anywhere in the text (B:utils/helpers.py:78). RAN /root/merge-scratch/run/C3/sanitize_out.txt: 'ClientA PRIVATE LIMITED' gives A ClientA_Pvt_Ltd and B ClientA_PRIVATE_LIMITED; 'Privateer Unlimited Co' stays unchanged in A and becomes Pvter_Unlimited_Co in B. The map's rule row A-R19 / B-R19 already records this; the feature row does not.
- MAP §1 row A-F31 and MAP §4.1 (322717f) say the file-name card is 'built from the real patterns'. More precisely, it is built from the description strings stored next to each pattern, not from the regular expressions (A:gui/tabs/setup_tab.py:494-503 reads info['description']; the texts are at A:utils/constants.py:19,27,35,43,51,59). So the card can still drift: it omits .xls, the '(N)' forms and the freedom of letter case (A audit §3 A-F31, finding 69).
- MAP §1 row A-F24 / B-F23 says 'B writes into the run folder shown on Step 4'. More precisely, it is the run folder shown when Start Extraction was pressed. The worker stores it at B:gui/controllers/extraction_controller.py:29 and uses it at :92. A rebuild of the Step 4 list during the run does not move the workbook (RAN B audit, B-F34 Run 1).
- MAP §1 row A-F27 leaves out that A's settings file also holds recent_folders (A:gui/utils/cache_manager.py:31-43), which nothing reads (A-R117). The row is incomplete, not wrong.
- Commit 12d505f's message says 'On failure the destination is no longer left with a half-written template copy'. That holds only for a failed Excel session. If the final copy to the network breaks, a partial file is still left behind: A:core/excel_handler.py:215 raises, :226-228 returns False and nothing is cleaned up (A audit A-G12). This limits what the commit message claims; it is not an audit error.
- A's code comment A:utils/constants.py:85-90 calls 235 chars / 27 files over (old names) and 215 / none (short names) 'Measured on a real run'. Commit 7347383's message says these are modelled figures that assume a planned shorter share root, and that the measured run gave 270 chars and 1094 of 3030 files over ~218 (READ git show -s 7347383). The A audit already records this under A-R41 / A-G41, and MAP 2.1 item 6 quotes the message correctly. The error is in the code comment only.
- The earlier draft of C3 row 8 said the client-key and status rules are 'identical'. MAP 2.1 items 4-5 list them as not the same, because of the way back from the table (A-R17: A rebuilds the key from the screen text, A:gui/handlers/client_handler.py:169-170; B keeps it, B:gui/views/validation_view.py:371) and because of the display and the 'Complete Only' test (A-R29: contains, A:gui/handlers/client_handler.py:152, against equals, B:gui/views/validation_view.py:446). Fixed in C3.

### Questions for the author
- Do the ITC and Sales templates' Power Query queries read the category folder path from the Links cells, or do they contain a fixed folder name? How to check: open a real template, go to Data > Queries, and look at the Source step of each query.
- Do browser duplicate downloads such as 'ImsReco-<client>-<state>-<date> (1).xlsx' occur in the office? If both a plain and a '(1)' file exist, which one should feed Links B11?
- Which path-length figures fit the office share today? A's commit 7347383 gives 270 chars measured on the current root, while its code comment gives 235/215, which assume a planned shorter share root.
- In the consolidated workbook, do you want one row per client (A) or one row per version folder (B)?
- Should Step 4 show the latest batch by date modified (A) or by name (B), or should the user pick the batch? How often are older summaries opened in Excel after a newer run (this changes date modified)?
- Should Dry Run refuse to preview until both templates and the target are chosen (A), or preview with only a client ticked (B)?
- Which template file types are used in practice: .xltx (A's picker), .xlsm (B's picker) or .xlsx (both)?
- Is a per-client 'include client name in folders' option wanted, or only the global tick?
- May a merged tool keep a Step 4 log file and a run history that copy every extracted figure into plain-text files outside the reports (B-G64)?
- What fills Info!BB2 in the templates, and is a non-empty BB2 enough to accept a refresh?
- Should the consolidated look-up for 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' use Diff GSTR-3B!J15/K15/L15 (both projects today) or the cells set in the mapping (A-G1/B-G1)?

## C4

### Decision candidates
- [GST RULE] IMS Reco file names with a browser '(N)' after the 8-digit date: accept or not, and which of two IMS files goes into ITC Links B11 ⚠ (4A-5; A-R2 / B-R7; A-R27/A-R57 / B-R25/B-R56; A-G6, A-G38 / B-G48). A's 4 tests pin acceptance.
- [GST RULE] Category folder names written into the ITC/Sales Links cells ⚠: short GSTR3B/ITC/Sales (A) or long 'GSTR-3B Exports'/'Other ITC related files'/'Sales related files' (B). First check the templates' queries (4A-27; A-R39, A-R41, A-R54, A-R55 / B-R37, B-R54, B-R55; A-G41 / B-G54).
- [GST RULE] Shortest allowed client-folder name, and whether the state code must always survive the cut. A accepts typed values below its 15 floor; B's box refuses them (4A-20, 4B-20; A-F6, A-R35, A-R36, A-G21, A-G26 / B-F4, B-R32, B-G24).
- [GST RULE] Whether the per-client 'client name in folder names' option returns. If it does, store it under '<client>-<state code>' as A's 33d0baf does (4A-12; A-F10, A-R40, A-R17 / B-F5, B-R36).
- [GST RULE, same formula in both] Step 2 should hand back the scan key unchanged (B's way) to avoid A's 'Fatal error' for number-like client names (4B-10; A-R17, A-G25 / B-R17).
- [GST RULE, display] Wording of the 'Expected file names' card: built from the patterns (A) or hand-typed (B). Neither mentions .xls or the '(N)' forms (4A-21; A-R16 / B-R13, B-R141).
- [GST RULE] Whether a report may be produced when a Links value cannot be written (openpyxl path): A skips silently, B fails and leaves an unfilled copy (4B-9; A-R58 / B-R53).
- Technical: scan speed. One folder listing, cached per-file details, and signatures read 16 at a time (A), combined with the scan in a worker thread with live progress (B) (4A-1, 4A-18, 4A-19, 4B-1, 4B-19; A-R109, A-R110, A-F7, A-G48 / B-F6, B-F34, B-G58).
- Technical: which rejected files the scan reports. Tiny, wrong-signature and unreadable files are counted (B) or vanish (A); skip ~$ lock files; show the names, not only counts (4B-17, 4B-18; A-R8 to A-R10, A-G37, A-G67 / B-R1 to B-R4, B-G20, B-G22).
- Technical: build Step 3 reports in a local staging copy and copy once (A), so nothing half-made is left at the destination (4A-2; A-R98, A-G12 / B-G5, B-G26).
- Technical: Step 4's choice of batch and version. Run folder by date modified (A) or by day-first name (B); newest version only (A) or every version selectable (B) (4A-7, 4B-25; A-F19, A-R42, A-G45 / B-F18, B-R38, B-R39, B-G12).
- Technical: recognise refreshed copies by the configured suffix in the list and in Skip Refresh (A), without A's over-matching for short suffixes (4A-26; A-R68, A-R69, A-G16 / B-R136, B-G9).
- Technical: pywin32 at start-up. Start without it off Windows, require it on Windows with a clear reason, report packages plainly, fix B's wrong advice and missing requirements line. The install offer conflicts with Standard 13's 'nothing installed on users' PCs' (4A-15, 4A-32 to 4A-36; A-F1, A-R124, A-R125, A-R129, A-G7, A-G53, A-G59 / B-F1, B-R124, B-G32).
- Technical: start without a console (pythonw), and never create folders in the program folder (B's way) (4B-5, 4B-6; A-G22, A-G23 / B-G33, B-G34).
- Technical: all workers talk to the window only through signals (B's pattern; A's Step 3 still touches Tk directly) (4B-2, 4A-23; A-G70, A-R143 / B-F34).
- Technical: Step 4 needs a Stop button and must refuse a second run (B), which A lacks (4B-3, 4B-4; A-G50, A-G65 / B-F24, B-G71, B-G74).
- Feature choice: workflow choice organise-only / full pipeline (A), or always jump to Step 4 (B) (4A-30, 4B-31; A-F5, A-R154, A-G66 / B-F18).
- Feature choice: live Step 1 checklist and hover help (A only) (4A-20, 4A-41; A-F4, A-F31, A-R147, A-G20, A-G49).
- Feature choice: run history and a per-run Extraction_Log file (B only). Note that B's run history keeps extracted figures in plain JSON in the home folder (4B-12, 4B-26; B-F17, B-F25, B-R119, B-R122, B-G25, B-G59, B-G64).
- Technical: carry A's test suite, after the GST-rule decisions above fix the expected behaviour (4A-9; A-R126, A-G60).
- Technical: name cleaning of Private/Limited. Whole words in any case (A), versus substring replacement that also shortens run-together names (B). This is a naming convention, not a GST rule (4A-10, 4A-17; A-R19, A-R20, A-G40 / B-R19, B-R20, B-G18, B-G80).
- Technical: log hygiene. Remove the on-screen [DEBUG] lines (B did), and stop writing Links values (client paths) at INFO in both (4A-13, 4B-7; A-R122, A-G56, A-G69 / B-R121, B-G64).
- Technical: code clean-up B did and A did not. Bare excepts, paste markers, unreachable return, duplicated header block, and the tracked .claude/settings.local.json (4B-8; A-G74 / B-G79).
- Technical: dark-mode approach. A records and replays colours; B swaps stylesheets. Both leave parts light (4A-31, 4A-37 to 4A-39; A-F28, A-R148, A-G64 / B-F28, B-G69).

### Flags for other sections
- Excel-safety / Standard 12 section: A's 12d505f (local staging) is A's only Excel-safety gain after the split. Both projects still take over the user's own Excel with GetActiveObject and always Quit it (A-G54 = B-G63). B's in-place report building leaves an unfilled template copy under the report's name on failure (B-G5; RAN C4 core probe, key staging_12d505f), and B's Step 4 then offers it for refreshing (B:gui/pyqt_main_window.py:461).
- Rules section C5 / MAP 2.1 #9: the merged-Links-cell difference was RAN in C4 (scripts/merged_probe.py). A returns success with B2 empty and B4/B5/B7/B8 filled. B returns failure and leaves the template copy with no Links value written. The path is unreachable in B (the window needs pywin32) and in A runs only without a working pywin32.
- Tests section: A's 4 pure-logic test files re-run on a B copy give 7 failed, 66 passed (4 IMS '(N)', 3 Private/Limited). test_file_ops.py cannot load on B (no scan_excel_files). The same 95 pass on A (RAN C4, out2/Atests_on_B.txt, Atests_file_ops_on_B.txt, Atests_on_A.txt). This matches notes2/A_tests.md §11.
- Features section C3 rows 7 / 34 (scan): strace over 50 made-up files shows A lists the source folder once and B three times. Per file, both make 1 metadata request and 1 open on Linux. The Windows saving of A's cached stat rests on Python's documentation (not run on Windows). Speed probe with a 20 ms artificial wait per open, 100 files: A 0.15 s, B 2.03 s.
- Features section C3 rows 19 / 20: Step 4 run-folder choice re-run (RAN C4). A uses the 01-02-2026 batch by date modified; B uses the 31-01-2026 batch by its day-first name (B:gui/pyqt_main_window.py:432-448). B lists every Version folder; A only the newest (A:power_query_extractor/gui/extractor_window.py:641-648). Custom suffix _Done_{timestamp}: A reuses the existing copy on Skip Refresh; B makes a new copy and shows 'Pending'.
- Gap map: B's Step 4 is not rebuilt after an extraction that ends in the error path (B:gui/pyqt_main_window.py:563-565). This is minor and not registered as a gap.
- Gap map / C3 row 1: B has no 'Press Enter' pause at all (B:main.py:49-59), so a double-clicked console would close at once with its message (MEDIUM, not run on Windows). A's _pause (A:main.py:45-54) does not help without a console, because A dies earlier at A:main.py:17 (RAN C4).
- Repository hygiene: A's repository still tracks .claude/settings.local.json (git ls-tree A/main). B removed it in c258047 and ignores .claude (B:.gitignore:82). A still has 23 bare except: lines, 8 'ADD THIS' paste markers, an unreachable return (A:power_query_extractor/core/report_processor.py:382-385) and a duplicated header block (A:power_query_extractor/core/data_consolidator.py:73-80). B has none of these.

### Corrections to the audits or the map
- MAP §3 row A-G21 says 'typed values unchecked (8 or less drops the state code)'. That is true for two-letter state codes only. For a three-letter fallback code (an unknown state 'Xyz' gives 'XYZ'), a maximum of 9 also drops it: create_client_state_key('ClientA Traders','Xyz',9) gives 'ClientA T' (RAN C4; the function is identical at A:utils/helpers.py:620-644 and B:utils/helpers.py:513-537).
- B audit S12own Phase 7 item 5, and S3b line 36, say Step 4's list is rebuilt 'after every extraction'. The code rebuilds it only after a completed extraction (B:gui/pyqt_main_window.py:561). The error path _on_extraction_error (B:gui/pyqt_main_window.py:563-565) does not rebuild it (READ).
- MAP §4.1 row a603c02 says the commit 'Speeds up scanning a network folder without changing results'. That holds on Windows only. On Linux and macOS, names with capital extensions such as .XLSX are found from that commit on (RAN A audit on the parent commit and a603c02; RAN C4 at the snapshot: A finds GSTR3B-ClientB-Delhi-May.XLSX, B does not, B:utils/helpers.py:255-257).
- Confirmation, not an error: the length-box finding (A-G21) was re-run on A's REAL Step 1 screen (scripts/lenbox_real_A.py, built with the stand-in app object of A's own smoke tests). The interrupted earlier attempt at this section had used a stand-in Spinbox. Typing 8 keeps 8 (A:gui/tabs/setup_tab.py:440-442, read at A:gui/handlers/processing_handler.py:194), and two made-up clients then share one folder 'ClientA '. B's real Step 1 refuses the typed 8 and stays at 35 (B:gui/views/setup_view.py:373-376).

### Questions for the author
- IMS Reco: should a file named like 'ImsReco-<client>-<state>-DDMMYYYY (1).xlsx' count as that client's IMS reconciliation file? When both the plain file and the '(1)' copy are present, which one should be linked in the ITC template (Links B11)? A links the '(1)' copy; B ignores it.
- Do your ITC and Sales templates' Power Queries take every folder path only from the Links cells, or do any queries name 'Other ITC related files', 'Sales related files' or 'GSTR-3B Exports' themselves? How to check: open each template's Power Query editor and search for those names. This decides whether A's shorter folder names are safe with your templates.
- What is the shortest client-folder name you want to allow, and must the state code (e.g. '-MH') always stay in the folder name? In A, typing 8 into the length box puts two different clients into one folder.
- Do you need the per-client 'include client name in folder names' choice on Step 2, or is the single global tick enough? B removed the per-client choice; A has it and fixed its key.
- In Step 4, do you ever need to refresh an older Version folder of a client (B lists every version), or only the newest (A)?
- Which 'Refreshed File Suffix' do you use in practice? A's commit 2aaa2b1 says a custom suffix was 'Reported' as broken; B recognises only '_Refreshed'.
- Do the office PCs start the tool with a console window (python main.py, or a double-click on main.py) or without one (pythonw)? From which folder? The log file lands in that folder, and A cannot start at all without a console.
- Was the Windows check that A's commit 12d505f asked for ever done? Do reports built in %TEMP% and copied to the share open and refresh normally?
- Do you want a per-run Extraction_Log text file beside the consolidated workbook (B), and a run history? B's history stores extracted figures in a plain file in your home folder.
- Do you want the workflow choice 'organise only' / 'full pipeline' (A), or should Step 3 always hand over to Step 4 (B)?

## C5a

### Decision candidates
- IMS Reco file names with a browser '(N)' after the date: accept (A) or reject (B). Includes the effect on 'Has duplicates' and on which IMS file Links B11 names. Refs A-R2/B-R7, with A-R27/B-R25 and A-R57/B-R56. GST RULE ⚠.
- Which file of a type feeds a Links cell when duplicates exist: the first in sorted path order, where a '(1)' copy sorts first (same in both). Should duplicates be flagged or chosen differently? Refs A-R57/B-R56, A-R27/B-R25. GST RULE.
- File discovery: case of extensions, '~$' names, folders named like workbooks, and too-small / non-Excel / unreadable files (silent in A, listed as variations in B). Refs A-R8/B-R1, A-R9/B-R2, A-R10/B-R3, B-R4. Not a GST rule.
- Show the list of unused source files with naming hints to the user (A: summary workbook Variations sheet; B: never shown). Refs A-R12/B-R12, A-R61/B-R94-B-R95, A-R11/B-R5. GST RULE ⚠ (the hint texts name return types and date formats).
- Step 1 example-file-name card: generated from the pattern descriptions (A) or hand-typed (B). Should it mention '.xls', the '(N)' forms and the hyphen ban? Refs A-R16/B-R13/B-R141. GST RULE ⚠.
- How Step 2 hands the client key to Step 3: rebuilt from table text, which crashes on number-like names such as 0042 (A), or stored with the row (B). Refs A-R17/B-R17. GST RULE tag (state-code key); the crash is technical.
- Private/Limited abbreviation in file names: whole-word, any case (A) or case-sensitive substring (B). Also whether file names and the client-folder key should abbreviate the same way, including underscore-joined names. Refs A-R19/B-R19, A-R20/B-R20, A-R35/B-R32. Not a GST rule (A-R35 state-code part is GST).
- Step 2 status display: row tint with icon (A) or three text colours (B); 'Complete Only' contains vs equals. Refs A-R29/B-R27 (with A-R138/B-R134). GST RULE tag (display only).
- Scan statistics: what counts as found / unparsed / variations, and where the counts are shown. Refs A-R31/B-R29. Not a GST rule.
- Max client-folder length: allowed range, validation of typed values, greyed-but-used box, and whether the state code may ever be cut. Refs A-R36/B-R32/B-R138, A-R35. Not a GST rule (state-code part of A-R35 is GST).
- Category folder names: short GSTR3B / ITC / Sales (A, for Excel's ~218-character path limit ⚠) or long names (B). Decide together with the templates' queries. Refs A-R39/B-R35, A-R41/B-R37 (GST RULE ⚠), A-R54/A-R55/B-R54/B-R55 (GST RULE ⚠).
- Per-client 'include client name in folders' control: keep A's Step 2 toggle, or B's global-only setting with its misleading note. Refs A-R40/B-R35/B-R36. Not a GST rule.
- Step 4 choice of run folder (newest date modified, highest name, or user-chosen) and of versions (newest only, or all). Refs A-R42/B-R39, A-R43/B-R38. Not a GST rule (decides which data set is reported).
- Processing-mode texts and Resume semantics (Resume skips only within the same minute; B's card says 'Re-Run: Overwrite existing'). Refs A-R43/A-R44/B-R44/B-R45. Not a GST rule.
- Dry Run check: require templates and target as Start does (A) or not (B). Refs A-R51/B-R48/B-R133. GST RULE (A registry tags it).
- Template file types offered and accepted (.xltx/.xlsx vs .xlsx/.xlsm; the report is always .xlsx; checklist only in A). Refs A-R52/B-R48/B-R51, A-R53/B-R49. Not a GST rule.
- Handling of a Links cell inside a merged range on the openpyxl path (skip silently vs fail and leave a template copy). Refs A-R58/B-R53, with A-R96/B-R97. GST RULE (B registry) ⚠.
- Step 4 report recognition: prefix test in any letter case (A) or glob search (B). Refs A-R59/B-R51/B-R61. GST RULE.
- During Step 3, read the template paths live (A) or snapshot them at Start (B); also the client processing order. Refs A-R60/B-R50/B-R60. GST RULE (A-R60, B-R50 tagged).
- Assemble each report in a local temporary copy and copy it once (A), or copy the template to the final name first (B; leaves a broken 'report' on failure). Refs A-R98/B-R52, A-R99/B-R59. Not a GST rule (file safety).

### Flags for other sections
- UNREGISTERED rule difference (for category 16 and the decision list): Step 3 client order. A follows the order in which the scan first met each client in the sorted file list (A:gui/handlers/client_handler.py:30, 165). B sorts by client key (B:gui/views/validation_view.py:367, since c258047). RAN: 'Zeta Traders' before 'Alpha Traders' in A, the reverse in B. This affects which client wins when two clients share a folder, and the running totals in organisation reports.
- Within B, Step 3's Re-Run/Resume picks the run folder by newest date modified (B:core/file_organizer.py:167), but Step 4 picks it by highest name (B:gui/pyqt_main_window.py:432-448). The two steps can work on different runs (RAN modes_probe and step4_probe). Relevant to the Section 6/7 gap lists.
- B leaves an unopenable or unfilled template copy under the final report name when report creation fails (B:core/excel_handler.py:122-129, 227-231). B's Step 4 then lists it as a report (B:gui/pyqt_main_window.py:461). RAN step3broken. Relevant to C5b (category 9, B-R61) and the gap lists.
- The Variations list is never shown in B: neither on screen nor in the summary workbook (B:gui/controllers/processing_controller.py:73; the worker is not given ScanResult.variations, B:gui/pyqt_main_window.py:366-371). Relevant to the feature comparison (B-F9 / A-F15) and Section 3.
- B's Step 1/2 statistics show 'Unparsed: 2' next to 'Variations: 5' for the same scan, because rejected files count as variations but not as files found (RAN display_probe). Relevant to the screens section (category 16).
- In A, typing a max length of 8 or less (A:gui/tabs/setup_tab.py:440-442; no validation) drops the state code. One client's registrations in two states then share one client folder (RAN: 'ClientA-' for Delhi and Maharashtra). Typing text such as 'abc' stops Step 3 with 'Fatal error' (A:gui/handlers/processing_handler.py:194). Relevant to the gap map and to C5b category 16.
- Running B's window code on Linux requires stand-in pywin32 modules even to import it (scripts/fake_win32.py). This confirms B-R124 for the dependencies section; any future test of B's GUI code needs the same.
- Two spellings of one client in one state ('<Client> PRIVATE LIMITED' / '<Client> Private Limited'): in A, copies and report names coincide and overwrite each other, while both are reported 'Success' (RAN collide_probe). In B only the client folder is shared. Relevant to the gap map (A-R35/A-R59 collisions).

### Corrections to the audits or the map
- MAP.md row A-R58/B-R53 and 2.1 item 9 say a merged Links cell 'makes the report fail in B'. That holds only when the Links cell lies inside a merged range but is not its first cell. RAN merged_probe: with B2 as the first cell of merged B2:C2, both projects write B2 normally and succeed. With B2 inside merged A2:B2, A skips B2 silently (bare except, A:core/excel_handler.py:302-303). B fails ('MergedCell' object attribute 'value' is read-only; except (KeyError, ValueError) at B:core/excel_handler.py:251) and leaves the unfilled template copy on disk under the report name, which the MAP row does not mention.
- MAP.md row A-R42/B-R39 says 'client folders without such a version dropped in both'. A lists a client folder whose newest Version folder holds no report, with both ticks off; only folders with no Version-* folder at all are dropped (A:power_query_extractor/gui/extractor_window.py:641-659). B drops it. RAN step4_probe with ClientD-TN: offered in A, absent in B. The A audit's A-R42 text is correct.
- B registry and B audit (B-R1, B-R3) omit that B's scan lists a FOLDER whose name ends in .xlsx/.xls/.xlsm. B records it as the variation 'Cannot read file: [Errno 21] Is a directory: ...' (RAN core_probe on Linux; glob at B:utils/helpers.py:255-257 does not exclude folders; open at B:core/file_parser.py:195). On Windows it would probably be 'File too small to be valid Excel', because a folder reports size 0 (MEDIUM, not run). A drops such folders silently (regular-file test, A:utils/helpers.py:296).
- A registry row A-R39 repeats the code comment at A:utils/constants.py:85-90 as if 235 chars / 27 files and 215 / 0 were measured on the real 3030-file run. The message of commit 7347383 says the real run's worst path was 270 characters, with 1094 files (36%) over ~218. The other figures were modelled with a planned shorter share root. The code comment itself also says 'Measured on a real run'. The A audit Section 5 and A_cross_flags.md:330/344 already note this; the registry row still has the old wording.
- A's code comment at A:gui/tabs/setup_tab.py:478-479 (quoted in MAP 2.1 item 3) says the Step 1 card 'can never drift from the regexes the parser actually uses'. The card shows the hand-written 'description' strings (A:utils/constants.py:19-59), not the regexes. The IMS description (A:utils/constants.py:27) does not reflect the '(N)' that A's IMS regex (A:utils/constants.py:23) accepts.
- B registry and B audit (B-R27) say Step 2 'colours the whole row'. B colours the TEXT (foreground) of every cell in the row, not the row background (B:gui/views/validation_view.py:395-397 setForeground; RAN gui_probe_B). MAP's wording 'B colours the text' is correct.
- Commit 275d9a4's message (cited in A-R59 history) calls the new prefix-test helpers 'proven equivalent to the original glob logic'. This holds on Windows only. On Linux a hand-renamed ...XLSX report is ticked by A's prefix test but not found by A's own search ITC_Report_*.xlsx ('No ITC report found'). RAN step4_probe: A:power_query_extractor/gui/extractor_window.py:36 vs A:power_query_extractor/core/report_processor.py:112.

### Questions for the author
- [GST] When a client's folder holds both ImsReco-<client>-<state>-DDMMYYYY.xlsx and a browser re-download '... (1).xlsx', which file should the ITC report's IMS cells (Links B10/B11) use? A uses the '(1)' copy; B ignores it. Should a client whose only IMS file is a '(1)' download count as having IMS Reco?
- Do the ITC and Sales templates' Power Query queries read the folders from Links B2/B7/B10 (ITC) and B2/B8 (Sales)? Or does any query contain a fixed folder name such as 'Other ITC related files', 'GSTR-3B Exports' or 'Sales related files'? To check: open each template, Data > Get Data > Launch Power Query Editor, then Advanced Editor for each query, and search for those names. This decides whether A's short folder names work with your templates.
- Are any client names in your file names written as pure numbers, for example office codes with leading zeros like 0042? A's Step 3 stops on such a client.
- Should files the program cannot use (too small, not an Excel file, unreadable, folders named .xlsx, unrecognised names) be listed by name somewhere the staff will see, with the naming hint?
- Are your templates .xlsx, .xltx or .xlsm files? A's file picker offers .xltx/.xlsx, B's offers .xlsx/.xlsm, and the report is always saved as .xlsx.
- During a batch, does anyone change the template on Step 1? Should a batch always use the templates chosen when Start was pressed?
- In Step 4, should older Version folders of the same run (from a Re-Run) be offered (B) or only the newest (A)? Should the run folder be the newest by date in its name, by last change on disk, or chosen by you?
- Do you want 'PRIVATE LIMITED' in capitals abbreviated to 'Pvt_Ltd' in copied file names and report names (A), or kept as written (B)?
- What maximum client-folder length do you use? Should values below 15 (which can remove the state code from the folder name) be allowed?
- Which wording should Step 1 use for the example file names: with '.xlsx' and 'ClientName' (A), or the short form (B)? Should it mention that .xls files and '(1)' downloads are accepted?

## C5b

### Decision candidates
- Skip refresh: which existing copy counts as already refreshed and may be reused without re-checking. A-R68, A-R69 against B-R63 (with A-G16, B-G9, A-G2, B-G7). [GST RULE] (treated as one when in doubt: it decides which file the consolidated figures come from).
- Refresh success test Info!BB2 ⚠: same in both. Author to confirm the sheet, the cell and the accepted values. A-R73 / B-R70, B-R71. [GST RULE].
- The 22 extraction cells ⚠ and the 9-column mismatch (IMS Reco, Extra In Purchase and Extra In GSTR2B show Diff GSTR-3B J15/K15/L15): same in both. A-R74 to A-R83, A-R88 / B-R73 to B-R83 (A-G1 / B-G1). [GST RULE].
- 'Missing' for a failed refresh and the #,##0.00 display: same in both. A-R89, A-R90 / B-R88, B-R89. [GST RULE].
- Calculated columns, dormant, same in both (keep or drop). A-R85 / B-R84. [GST RULE] (dormant).
- Dry Run gate: whether templates and a target must be chosen before a preview. A-R134 (A-R51) against B-R133, B-R135. [GST RULE] per MAP 2.1 item 7.
- Help text listing the expected input file names ⚠: read from the patterns (A) or hand-typed (B). A-R16, A-R147 against B-R141. [GST RULE].
- Step 3 summary workbook Variations sheet (real list in A, always empty in B); this decides whether the unrecognised-file hint is ever shown. A-R61 against B-R94, B-R95; touches A-R12 / B-R12. [GST RULE] for the hint display, technical otherwise.
- Which reports and Version folders Step 4 offers and pre-ticks; one row per client or per Version folder in the consolidated workbook; meaning of 'Total Clients'. A-R65, A-R87, A-R91 against B-R61, B-R136, B-R87, B-R90 (folder choice is category 6, A-R42 / B-R39). Technical/layout, not a GST rule.
- Consolidated workbook location (last scan, or the folder fixed at Start), behaviour on write errors, and an extraction log file. A-R86 against B-R86, B-R122. Technical.
- pywin32 policy: guard in Step 4, start-up check, install offer, Step 4 loading card. A-R64, A-R124, A-R125, A-R129 against B-R124, B-R125. Technical.
- Step 3 report staging on local disk, or in place at the final name. A-R98 against B-R52, B-R105. Technical (file safety).
- Shared Excel session behaviour (takes over the user's Excel, Dispatch rather than DispatchEx, CoUninitialize placement) set against Standard 12. A-R97, A-R103 / B-R96, B-R101, B-R110. Technical, same in both.
- Step 4 wait box range (5-60 unchecked, or 5-120 enforced) and persistence of the Step 4 options. A-R72, A-R119 against B-R65, B-R116. Technical.
- Scan performance: one listing plus parallel signature reads, or three searches. A-R109, A-R110 against B-R1, B-R3. Technical.
- Run folder computed once per run. A-R111 (B recomputes per client). Technical.
- Background threads, re-entrancy and Stop for Step 4. A-R135, A-R143 against B-R111, B-R112. Technical.
- Settings storage, save points, recent folders and migration from A's file. A-R114 to A-R117 against B-R115 to B-R118. Technical.
- Run history and extraction log file (confidentiality: figures kept in the user profile). A-R123 against B-R119, B-R120, B-R122. Technical.
- Log file locations (current folder) and the extractor log. A-R120, A-R121 against B-R121, B-R127, B-R128. Technical.
- Console handling and program-folder writes at start-up. A-R127, A-R128 against B-R125. Technical.
- requirements.txt ⚠: minimum against exact versions; pywin32 listed or only commented. A-R126 against B-R124. Technical.
- Step 2 table: rows start unticked or all ticked; missing types named or only counted; sortable columns; Last Run column. A-R138 against B-R134. Technical/presentation.
- Features only in A to keep or drop: setup checklist, tooltips, workflow choice, export of the client list, long-name warnings, per-client include-name toggle, wheel rules. A-R133, A-R113, A-R147, A-R154, A-R142, A-R136, A-R137, A-R146. Technical.
- On-screen logs: plain text with time stamps (A), or HTML that drops text (B). A-R140 against B-R137, B-R139. Technical.
- Refresh-status display: per report type with the configured suffix, or per Version folder with '*_Refreshed*'. A-R145 against B-R136. Technical.
- Keyboard access and greyed buttons. A-R135 against B-R138. Technical.
- Step 4 folder source (follows Step 1 as typed, or the last saved target filled from disk at start-up). A-R144 against B-R39, B-R126. Technical.
- Window, title strip, version labels and dark mode. A-R131, A-R148, A-R155, A-R156, A-R130 against B-R129, B-R144, B-R130, B-R132. Presentation.

### Flags for other sections
- Gaps section (section 7 comparison): with suffix '_R{timestamp}' (and, by the same rule, '_{timestamp}'), A's Skip refresh treats the ORIGINAL Step 3 report as an existing refreshed copy. It reuses it, reports 'Skipped - Using existing refreshed file' as a success, and the consolidated workbook then shows unrefreshed figures under a tick (RAN C5b p9.py d2 S3; A:power_query_extractor/core/refresh_naming.py:23-66). A-G16 records the over-match; check that this Skip-refresh consequence is recorded too.
- Gaps section: in A a Step 4 scan during a Step 4 run redirects the consolidated workbook into the newly scanned target's run folder, although the workbook holds the first target's clients (RAN C5b p11a.py; A:power_query_extractor/gui/extractor_window.py:632, 945). Check whether an A-G gap records it.
- Gaps section: in 'rerun'/'resume' mode, B recomputes the run folder for every client. A newer 'Annual Statement-...' folder that appears mid-run (for example from a colleague's run) splits one B run across two run folders (RAN C5b p13.py; B:core/file_organizer.py:146-179). The MAP calls this 'same result unless the target changes during the run'. Check whether a B gap records it.
- Gaps / Excel-safety appendix (X1): in B's Step 3 COM path, when Excel cannot open the copied template, the unfilled template stays under the final report name (RAN C5b p12.py with the fake Excel; B:core/excel_handler.py:122-129). A leaves no file there (A:core/excel_handler.py:135-155, 212-215). The same happens on the openpyxl path for a template that is not a workbook (RAN).
- Gaps / confidentiality: B's run_history.json keeps the whole Step 4 result (every extracted figure, paths, BB2 text) indefinitely in the user profile, with pq_refreshed=True even when the refresh failed (RAN C5b p14.py; B:models/run_history.py:73-95).
- Category 6 (C5a) / settings: a client_name_max_length of 0 in A's settings file is restored as 35 by A. B migrates it as 0, its box shows 15, and B saves 15 on closing (RAN C5b p14.py). So client folder names are cut at different lengths after moving from A to B.
- Section 8 / Standard 12: both projects' Step 3 code takes over the user's own running Excel (GetActiveObject), hides it and calls Quit on it. Shown with the fake Excel in C5b p12.py ('excel.Quit USER-OWN-EXCEL' in both; A:core/excel_handler.py:123-128, 230-252; B:core/excel_handler.py:110-115, 203-211).
- Category 8 (C5a): the merged-Links-cell nuance above (only non-top-left cells inside a merge differ) affects MAP row A-R58 / B-R53 and MAP 2.1 item 9.
- Screens section (S10): B's on-screen logs drop leading spaces and any text between '<' and '>' (RAN C5b p16.py: '    Links B11 = <ImsReco-ClientA-Delhi-31032025> & done' became 'Links B11 = & done'). In A, a greyed Step 3 Start still reacts to a mouse press, and an enabled Start does not react to the keyboard (RAN C5b p166).
- Probe limitation, not a finding: A's Step 3 worker started in its own thread did not finish when the probe pumped the window with update() instead of mainloop(). The worker body was therefore run on the window's thread (C5b p16.py). Nothing is claimed about A's real threaded behaviour from this.

### Corrections to the audits or the map
- A audit Section 5.17, A-R152 says the four MESSAGES texts are 'shown as Warning boxes before Dry Run and Start'. The code shows three of them in error boxes titled 'Error': no_source (before a SCAN, A:gui/handlers/file_handler.py:123), no_templates (:135) and no_target (:139). Only no_selection is a 'Warning' box (A:gui/handlers/processing_handler.py:45, 83). RAN C5b p16.py: the scan gates and the Dry Run template gate produced showerror boxes titled 'Error'.
- MAP row A-R58 / B-R53 says 'a merged Links cell is skipped silently in A, makes the report fail in B'. B registry B-R97 and B audit S45 B-R97 say 'a merged target cell makes the report fail'. This holds only when the mapped cell lies INSIDE a merged range and is not its top-left cell (RAN C5b p12.py, A2:B2 merged, mapped cell B2): A returns True with B2 empty; B returns False and leaves the template under the report name. When the merge starts at the mapped cell (B2:C2), both projects write it and succeed (RAN, with the openpyxl in venv-A and venv-B).
- B registry B-R136 says 'a custom suffix always shows Pending'. The precise rule is: any suffix whose copies do not contain '_Refreshed' (B:gui/pyqt_main_window.py:466-479 searches for '*_Refreshed*'). B audit S45 B-R136 already words it this way. RAN C5b p145: '_PQ_{timestamp}' gave 'Pending'.
- B registry B-R63 says 'with any other suffix nothing is found and a normal refresh runs'. The precise rule is: any suffix whose copies do not match '<stem>_Refreshed_*.xlsx' (B:power_query_extractor/core/report_processor.py:100-101). For example '_Refreshed' with no timestamp is not found (RAN C5b p9.py d2 S5), while a suffix starting '_Refreshed_' still is. B audit S45 words it as 'a suffix that does not start with _Refreshed_'.

### Questions for the author
- Do the current ITC and Sales templates still fill sheet 'Info' cell BB2 (⚠) with the client name after a successful refresh, and should a value from another client or a value already present in the template count as success?
- Does the office ever change the 'Refreshed File Suffix' from _Refreshed_{timestamp}? If so, which suffix is used? This decides how much the Skip-refresh difference (9.4) matters.
- When 'Skip refresh' is ticked, may the tool reuse an existing refreshed copy without checking that its own refresh succeeded (both projects do so today)?
- Should the consolidated workbook show one row per client (newest Version folder only) or one row per Version folder, and what should 'Total Clients' count?
- Which cells should the IGST/CGST/SGST columns of 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' show (configured AA1/AD1/AG1, O1-Q1 and S1-U1 ⚠; today both projects show Diff GSTR-3B J15/K15/L15)?
- Should the Step 4 options (wait time, suffix, Skip refresh) be remembered between sessions, and what maximum wait is sensible (A allows 60, B 120)?
- Which Excel version and display language do the office PCs use? The Alt+A, R, A key tips and the error-dialog words assume an English Excel (⚠).
- Is 218 characters the path limit you actually meet in your Excel version, and is Windows long-path support enabled on the office PCs (⚠)?
- Should Step 2 start with all clients ticked (B) or none (A), and should it name the missing return files (A) or only count them (B)?
- Should a Dry Run require both templates and a target to be chosen first (A) or only a ticked client (B)?
- Which wording should the expected-file-names help card use (with or without .xlsx, mention of '(N)' duplicate forms and of .xls)?
- Is it acceptable that B's run history keeps every extracted figure for every client in the user's profile folder indefinitely?
- Will the tool be started without a console window (pythonw / a desktop shortcut), or from a network program folder that users cannot write to?

## C6

### Decision candidates
- Window toolkit for 2.0: A Tk vs B PyQt5; Standard 12 names PySide6 (A-F2 / B-F1; §6 D5). Not a GST rule.
- Start-up and launcher: A's console requirements report and pip install offer vs B with no check and no window without pywin32; Standard 13 calls for a .bat/pythonw launcher, no console, excepthook and LOCALAPPDATA (A-F1, A-R124-A-R127 / B-F1, B-R124, B-G32, B-G35; §6 D1, D11, 6.9). Not a GST rule.
- Step 1 design: A's live checklist with pill and greyed Scan vs B's static cards and SCAN LOG; Scan button below the fold in both (A-F4, A-R133 / B-F33, B-G68; §6 D12, D13, D22). Not a GST rule.
- Workflow after Step 3: A's organise-only/full choice vs B's always-move-to-Step 4; whether the refresh may ever start automatically (A-F5, A-R154 / B-F29; §6 D14, D44). Not a GST rule.
- Browse buttons and template picker filter (A *.xltx *.xlsx with a checklist accepting 4 endings vs B *.xlsx *.xlsm with 3 invisible buttons) (A-F3 / B-F3, B-G67; §6 D15, D16). Not a GST rule.
- Expected-file-names help text on Step 1: A's 'type + …ClientName…xlsx' vs B's short list without extension; whether to mention .xls and the (N) forms (A-R16, A-R147 / B-R13, B-R141; MAP 2.1 item 3; §6 G1, D24). [GST RULE] ⚠.
- Step 2 Missing column: name the missing return files (A) or show a count (B) (A-R138 / B-R134, on top of the shared completeness rule A-R26 / B-R24; §6 G3, D29). [GST RULE] ⚠.
- Step 2 status display and 'Complete Only': whole-row tint vs three text colours; 'contains Complete' vs 'equals Complete' (A-R29 / B-R27; MAP 2.1 item 5; §6 G2, D31). [GST RULE].
- Client key with state code: where it is shown, and the round-trip from the table (A rebuilds it from text and breaks on '0042'; B stores it) (A-R23, A-R35 / B-R22, B-R32; MAP 2.1 item 4; §6 G4). [GST RULE] ⚠.
- Default ticks after a scan: none (A) vs all (B) (A-R138 / B-R134; §6 D30). Not a GST rule.
- Per-client include-name toggle and long-name warnings: keep A's, or drop them and B's note promising them (A-F10, A-R136, A-R137 / B-F4, B-F5, B-R141, B-G72; §6 D19, D33). Not a GST rule.
- EXPORT LIST: keep A's working export or remove B's dead button (A-F12, A-R142 / B-F10, B-G73; §6 D35). Not a GST rule.
- Dry Run gate: require templates and target first (A) or only a ticked client (B) (A-R51 / B-R48, B-R133; MAP 2.1 item 7; §6 G5, D37). [GST RULE].
- Template paths read per client during a run (A) vs fixed at Start (B) (A-R60 / B-R50, B-R60; MAP 2.1 item 11; §6 G6, D46). [GST RULE].
- Step 3 log and progress display: time stamps, per-file lines, the 14 [DEBUG] lines, Clear Log, readable progress bar, Stop button look (A-R140, A-R141, A-R135 / B-R137, B-R138, B-G70; §6 D40-D42). Not a GST rule.
- Locking controls during runs and confirming close during any run (A-R135, A-R143 / B-R111, B-R138, B-G40, B-G74; §6 D10, D47). Not a GST rule.
- Step 4 layout and folder choice: A's folder box + Scan + two sub-tabs vs B's one page with a tree built from the saved target; batch chosen by date modified (A) vs by name, day-first (B) (A-F18, A-F19, A-R144 / B-F18, B-R39, B-G12; §6 D48-D51). Not a GST rule.
- Step 4 feedback: failure reason on screen (A) vs FAIL only (B); end-of-run question 'Open the consolidated report?' (A) vs none (B); a success/failure count (A-F24, A-R143 / B-F24, B-R139, B-G71; §6 D56, D57). Not a GST rule.
- Meaning of the Step 4 refresh status ('Refreshed' green in B, time in A) given that both copy before refreshing (A-F20, A-R145 / B-F18, B-R136, B-G7; §6 D58). Not a GST rule.
- Step 4 Stop button and Extraction_Log file in the run folder (B-F24, B-F25 / A-F25, A-R143; §6 D55, D59). Not a GST rule.
- Visible Excel plus Alt+A, R, A keystrokes in Step 4 (same in both) vs Standard 12's hidden Excel with RefreshAll (A-F21, A-R71, A-F26 / B-F19, B-R67, B-F26, B-G61, B-G62; §6 D60, 6.9). Not a GST rule. The BB2 success check it shows (A-R73 / B-R71, §6 G8) is a [GST RULE] ⚠, the same in both.
- Extracted figures printed in the on-screen log and in B's text log (A-F23 / B-F25, B-G64; §6 D61). Not a GST rule; confidentiality decision.
- Dark mode: keep or drop; if kept, cover all screens (A-F28, A-R148 / B-F28, B-R144, B-G69; §6 D62, D63). Not a GST rule.
- Status bar content and B's tab dots (meaning 'ran' vs 'succeeded') (A-F2, A-R155 / B-F29, B-R140; §6 D7, D8). Not a GST rule.
- Settings: what is remembered and when it is saved; B's one-time import of A's settings file (A-F27, A-R116 / B-F27, B-F2, B-R117; §6 D9). Not a GST rule.
- Window size and minimum for 1366x768 laptops (A-R131, A-R155 / B-R129, B-G68; §6 D4). Not a GST rule.
- Standalone extractor launcher: a separate window (A) vs the main window opened on Step 4 (B) (A-F29 / B-F30, B-F31; §6 D2). Not a GST rule.
- Rewrite help texts, tooltips and message wording against the merged code (A-R147, A-R152 / B-R141, B-R131, B-G72; §6 D64, D66, 6.6). Not a GST rule, except the file-name card (G1).

### Flags for other sections
- §3 (C3) Row 2 and MAP A-F2: qualify 'both switch tabs after a scan'. B also jumps to Step 2 after an empty scan; A does not (RAN C6 probeA/probeB P3).
- §5 Category 16 (C5b 16.4) gate table: consistent with C6 6.6. C6 adds two RAN cases: 'Start when the ITC template path does not exist' (A 'Error' 'ITC template file not found' / B 'Warning' 'Please select a valid ITC template.') and 'scan of an empty folder' (A stays on Step 1 / B jumps to Step 2).
- Step 4 / Excel-safety sections: A writes the failure reason per report on screen ('⚠️ ITC Report - <reason>', A:power_query_extractor/gui/extractor_window.py:917-938); B writes only 'Result: ITC: FAIL, Sales: FAIL' (B:gui/controllers/extraction_controller.py:76-84). RAN C6. This bears on B-G71.
- Step 4 status: both copy the report to '_Refreshed_' before refreshing, so after a refresh that fails inside Excel both show a time. B labels it 'Refreshed' in green (B-G7); A's 'ITC: <time>' is not proof of success either. Relevant to the decision on what the status should mean.
- Start-up / Standard 13 section: A's main.py creates empty `logs` and `temp` folders inside the program folder at every start (A:main.py:29-43; RAN C6: they appeared in the run copy). Standard 13 uses `-B` so that users' PCs write nothing into the shared app folder (rebuild_pipeline_FINAL.md:969-970).
- GST-rules section (MAP 2.1 item 3, C5a Category 1): history detail for B's file-name card. c258047 (07-02-2026) showed the long 'ClientName…xlsx' descriptions in an HTML help panel; 7e09e5c (08-02-2026) went back to the shared short list (edac6a3, 47a1492:gui/tabs/setup_tab.py:362-373).
- UI decision list: in BOTH projects the Scan button is below the visible part of Step 1 at the opening window size. A: page 1,521 px tall in a 640 px view. B: column 1,070 px tall in an 815 px view, SCAN FILES at 989 px. RAN C6.
- Confidentiality (Step 4 / logs section): both print every extracted value on screen; B also writes the values into Extraction_Log_<…>.txt in the run folder (MAP A-F23/A-F25, B-G64).

### Corrections to the audits or the map
- MAP §1 row A-F2 says 'Both switch tabs on their own after a scan'. In A this is true only when at least one client is recognised: an empty scan stays on Step 1 with status '💡 No valid GST files found' (A:gui/handlers/file_handler.py:83-91). B switches to Step 2 after every completed scan, including an empty one ('0 clients (0 complete)', status 'Scan complete: 0 clients, 0 files, 0.0 B'; B:gui/pyqt_main_window.py:304-307). RAN (C6) probeA P3 / probeB P3.
- MAP §1 row A-F31 says A has 'a file-name card built from the real patterns'. The card prints the separate `description` strings stored beside each regex in FILE_PATTERNS (A:gui/tabs/setup_tab.py:494-503; A:utils/constants.py:19-59), not anything generated from the regexes, so it omits .xls, any letter case and the '(N)' forms (as the A audit §10.5.10 says). MAP 2.1 item 3 words it correctly. The same strings exist unused in B (B:utils/constants.py:19-59); B showed them in c258047 (c258047:gui/views/setup_view.py:168-173) and replaced them in 7e09e5c with the shared short list.
- Clarification (not an error) to MAP §1 A-F20 / A audit §10.8.3: A also copies the report before refreshing it (A:power_query_extractor/core/report_processor.py:164, like B:power_query_extractor/core/report_processor.py:120), so on Windows a refresh that fails inside Excel also leaves A's 'Last Refresh Status' showing a time (the original report's date, kept by copy2 unless Excel saved the copy). A avoids it only when pywin32 is missing, because it then returns before copying: RAN (C6), the rows stayed 'Never refreshed', while B's stand-in run left '_Refreshed_' copies and showed 'Refreshed (23/10/26 17:32)'.
- Environment note on the B audit §10.0 Limits: it says the log emoji are drawn as small boxes in the Linux screenshots; in this run the B screenshots show small emoji pictures, while A's Tk screenshots show nothing for emoji. This is not a code issue and changes no finding.

### Questions for the author
- How do you and your staff start the tool today: from a console with `python main.py`, through `launch_extractor.py`, or another way? (Standard 13 expects a no-console launcher.)
- Should Step 1 check each path as it is typed (A's checklist), or only when Scan is pressed (B)?
- After Step 3, should the tool move to Step 4 by itself, stay on Step 3, or let you choose? Should the Power Query refresh ever start without a click?
- After a scan, should the clients start ticked (B) or unticked (A)?
- [GST RULE] On Step 2, should the missing return files be named (A: 'IMS Reco, Sales Reco') or counted (B: '2')?
- [GST RULE] Which wording of the expected file names should Step 1 show (A's '…-ClientName-State-Period.xlsx' with type names, or B's shorter list without extension)? Should it mention .xls and the '(N)' duplicate forms?
- [GST RULE] Should Dry Run require the templates and the target first (A) or not (B)?
- [GST RULE] Should 'Complete Only' tick rows whose status contains 'Complete' (A) or equals it (B)? Should 'Has duplicates' have its own colour?
- Do you use EXPORT LIST, the per-client '📁 Name' toggle, or the Step 2 keyboard keys?
- Step 4: one page with a client → version tree (B), or a folder box with a Scan button and two sub-tabs (A)? Should Step 4 let you pick which 'Annual Statement-…' batch to refresh?
- At the end of Step 4, do you want the question 'Open the consolidated report?', a count of reports that succeeded and failed, and a text log of the run in the batch folder?
- Should closing the window during any run ask first?
- Do you use dark mode?
- What screen sizes and Windows display scaling do the office PCs and laptops use? (B's minimum window is taller than a 1366x768 laptop screen.)
- During a Step 4 refresh, have you seen letters appear in the wrong window, or another program's pop-up close by itself? (Same code in both.)

## C78

### Decision candidates
- Window toolkit for 2.0: Standard 12 says PySide6-Essentials, pinned to 6.12.x while any PC runs Windows 10 (Std:833), and the Tech Stack allows PyQt5 only for an existing PyQt5 app kept as it is (Std:208-209). A is Tkinter, B is PyQt5/Qt 5. Refs A-F2, A-F28, A-F31, B-F1, B-F28, B-F29, B-F33. Not a GST rule.
- How users start the app: Standard 13's launcher (Run [App Name].bat with pythonw -B -s) plus error hooks and a message box, against today's console start. Refs A-F1, A-F29, A-F30, A-G22, A-G30, A-R127, B-F1, B-F30, B-F31, B-G34, B-G35, B-G45, B-R125, B-R127. Not a GST rule.
- Where settings, logs and any run history live (%LOCALAPPDATA%\<App>\ per Std:1039-1040). Also: keep or drop B's run history and session file, and whether extracted figures may be stored outside client folders. Refs A-R114 to A-R122, A-G23, A-G56, B-F17, B-F25, B-F27, B-R115 to B-R122, B-G24, B-G25, B-G33, B-G44, B-G59, B-G64. Not a GST rule.
- Dependency handling: exact pinned versions, pytest left out (Std:937-940), packages only in the shared _runtime (Std:924, 933-936). Keep or drop A's start-up check and pip install offer; list pywin32 properly. Refs A-R124, A-R125, A-R126, A-G53, A-G59, B-R124, B-G32. Not a GST rule.
- Python version to target (Std:931-932 says 3.14) and whether Tkinter must be in the runtime. Client names ending in a dot are cleaned differently on 3.14. Refs A-G40, A-R20, B-G80, B-R20. Not a GST rule.
- Power Query protection: add Standard 12's DataMashup check before any openpyxl save (Std:846-850). Choose between refusing to start and falling back when pywin32 is missing; decide whether the openpyxl fallback report path survives at all. Refs A-G7, A-R96, A-R124, A-R129, B-G31, B-G32, B-R97, B-R124. Not a GST rule.
- Rewrite the Excel sessions to Standard 12: DispatchEx, never the user's Excel, DisplayAlerts off at Quit, Quit protected, IgnoreReadOnlyRecommended plus the ReadOnly check, CoUninitialize after the Excel function returns, a timeout or Stop. Refs A-R97, A-R101, A-R102, A-R103, A-G28, A-G43, A-G50, A-G54, A-G71, B-R96, B-R98, B-R99, B-R101, B-R110, B-R112, B-G38, B-G39, B-G53, B-G63, B-G74, B-G77. Not a GST rule (the standard already decides most of it).
- Refresh method: Standard 12's RefreshAll + CalculateUntilAsyncQueriesDone + Refreshing check, against today's Alt+A, R, A key presses, fixed waits and desktop-wide dialog clicking. Refs A-R71, A-R72, A-R105, A-G3, A-G44, A-G55, A-G73, B-R66 to B-R69, B-R102, B-R103, B-G6, B-G51, B-G57, B-G61, B-G62. The mechanism is not a GST rule, but what proves a refresh worked (Info!BB2, A-R73 / B-R71) IS a GST rule ⚠, to be decided separately by the author.
- Reading Step 4 figures through Excel (COM) or with openpyxl (Std:827 allows reading with openpyxl). Refs A-F23, A-R83, B-F22, B-R73 to B-R83. The reading technology is not a GST rule; the sheets and cells read ARE GST rules ⚠ (cell_mappings) and are not affected by this choice.
- Tests for 2.0: whether to carry over A's pure-logic tests. Name cleaning is technical (A-R20 / B-R20), but the IMS '(N)' pattern tests pin a GST rule ⚠ (A-R2 / B-R7, 5c8ce95). Once the author decides which cell each column means (A-G1 / B-G1, A-R88 / B-R79 to B-R83, a GST rule ⚠), add a test that pins that decision. Partly a GST rule.
- Error handling: Standard 13 error hooks and a message box; no bare except; worker errors sent back through signals (B's pattern) rather than touching the window from the worker (A-G70). Refs A-G30, A-G70, A-R127, B-G35, B-G42, B-R125. Not a GST rule.
- File safety: keep A's local staging of Step 3 reports (A-R98, 12d505f) rather than B's in-place build (B-R52, B-G5). Also write settings, history, summary and consolidated workbooks to a temporary name and rename, and clean up partial copies. Refs A-G11, A-G12, A-G28, A-G9, B-G26, B-G27, B-G43, B-G82. Not a GST rule.
- Shared dead code that embodies GST rules: keep or drop. validate_results' 'critical files' list (Annual Report, GSTR-3B Export; A-R30 / B-R28) and the dormant calculated columns (A-R85 / B-R84) are GST rules ⚠, the author's call. The rest of the shared dead code (cleanup_on_error, get_unique_path, calculate_file_hash, _write_section, _create_details_sheet, etc.) is not a GST rule; B's dead signals.py, CollapsibleGroup and unimportable extractor_main.py can go. Refs A-R49, A-R50, A-R63, A-R94, A-R95, B-R41, B-R46, B-R47, B-R58, B-R93, B-F31.

### Flags for other sections
- For C12/summary: health verdicts (technical only). Tests: A, clearly. Error handling: B, narrowly. Dead code: A. Silent bugs: about even, A narrowly. File safety: A, narrowly. Closeness to Standards 12-13: neither; B nearer on toolkit, start-up and logs, A nearer on tests and on stating dependencies. Source: C78 §8.6.
- Any section quoting Standard 12 'met / not met' counts should use C78 §7.2's single scoring, the same for A and B. Making reports: 4 met / 8 not met / 8 n.a.; refresh: 6 / 11 / 3; reading: 6 / 6 / 8; the fallback misses all 3 rules that apply. The two audits' own counts differ only because they read three rules differently.
- Reusing A's tests on B's shared logic: 4 of the 7 failures pin A's IMS Reco '(N)' file-name pattern (A:utils/constants.py IMS pattern from 5c8ce95, 23-07-2026; B lacks the '(N)' part). That is a [GST RULE] ⚠ classification (A-R2 / B-R7). Carrying A's tests over unchanged would pin A's version of that rule, so the decision list must leave it to the author. The other 3 failures are name cleaning (33d0baf).
- File-safety difference for C3 and C5 (A-F14/B-F13, A-R98/B-R52): A builds Step 3 reports in %TEMP%\gst_excel_temp and copies them to the target once (12d505f). B copies the template to the final report name first, so a failed fill leaves an unfilled report that Step 4 then refreshes (B-G5). B still creates %TEMP%\gst_excel_temp but never uses it.
- Toolkit facts for the decision list. B's PyQt5 wheel metadata says GPL v3 (B audit 9.1), stated as a fact only, no legal opinion. B's Qt 5 runtime (PyQt5-Qt5) is a separate 50 MB copy, so keeping PyQt5 next to PySide6 would put two Qt copies in the shared _runtime, both locked while any app runs (Std:1051-1054).
- Standard 13 start-up: A dies silently under pythonw at A:main.py:17. B survives but hides every start-up error (print only) and aborts with exit code 134 on any unhandled window-handler error (B-G35). Neither has the report_fatal_error hook (Std:984-1029). Relevant to C3 row 1 (A-F1/B-F1) and to Phase 5 starting files.
- Both projects put gst_organizer.log in the current folder (A:main.py:24; B:main.py:20; unchanged since the shared code). Under Std:962 that is the shared app folder, holding client names and figures (A-G56/B-G64). B's run_history.json also keeps extracted figures outside the client folders, untrimmed (B-G59/B-G64). Relevant to any section on outputs and client data.

### Corrections to the audits or the map
- Task brief vs the standards file: the brief says the CLAUDE.md copy's Tech Stack gives 'PySide6 with the PySide6-Essentials package, pinned to 6.12.x while any PC runs Windows 10'. The Tech Stack (rebuild_pipeline_FINAL.md:208-209) names PySide6-Essentials and adds '[— or PyQt5, only for an existing PyQt5 app that is being kept as it is]'. The 6.12.x pin is in Standard 12's table at line 833, not in the Tech Stack (READ).
- B audit S89 §9.9 says every start compiles 'B's 7,805 lines'. That is a newline (wc -l) count. Counting last lines that have no line end gives 7,821, as B audit S2 §2.1 says (RAN C78: 38 program files, 7,821 lines). For A: 8,930 program lines and 2,163 test lines, which matches A audit §2.1.
- The two audits score the identical Standard 12 COM code differently; this is not a code difference. (1) Rule 'COM only for Power Query workbooks it writes/refreshes' for Step 4 reading: A audit X1.3.4 C1 says NOT MET, B audit X1.3.4 R1a says MET. The code is the same: A:power_query_extractor/core/report_processor.py:527-636 and B:…:478-587, a COM read of a refreshed copy closed without saving. (2) The same rule for the openpyxl fallback: A audit X1.3.5 says 'not applicable' and files the failure under C2; B audit says R1a NOT MET (A:core/excel_handler.py:282-307 and B:core/excel_handler.py:231-255 behave the same). (3) The 'never RefreshDate as proof' rule for Step 3 and Step 4 reading: A audit C14 says MET, B audit R5e says not applicable. C78 §7.2 uses one scoring for both and states the differences.
- Re-confirmed, no change needed (RAN C78): A 143 try / 147 handlers / 23 bare / 102 Exception / 41 pass-only / 7 finally; B 90 / 91 / 0 / 79 / 15 / 3. A's tests: 288 collected; 1 failed / 287 passed under xvfb; 14 failed / 241 passed / 33 skipped without a display. A's pure-logic tests on B: 7 failed / 66 passed, and test_file_ops.py fails with ImportError scan_excel_files. Under no-console streams A stops at A:main.py:17 with AttributeError 'buffer' and writes no log; B loads and writes gst_organizer.log in the start folder.

### Questions for the author
- Does any office PC still run Windows 10? Standard 12 pins PySide6 to 6.12.x while any does (rebuild_pipeline_FINAL.md:833).
- For 'Annual Statement 2.0': does the PyQt5 exception in the Tech Stack ('only for an existing PyQt5 app that is being kept as it is', line 209) apply, or is it a new app on PySide6?
- Which Python version is 'python.exe' on your PC (the one with pywin32)? Will 2.0 run only from the shared _runtime on Python 3.14, as Standard 13 describes?
- How do staff start the tool today: a console with 'python main.py', a double-click, or a shortcut? From which folder? (The log file lands in that folder in both old versions.)
- Has anyone ever found a Step 3 report whose Power Query queries had disappeared? (Both versions still contain an openpyxl path that strips them while reporting 'Success'.)
- Does anyone use B's run history: the 'Last Run' column, or run_history.json, which also keeps every extracted figure outside the client folders?
- Should the tool ever install Python packages itself, as A's start-up offer does, or should packages only ever be added to the shared _runtime by you?
- Do staff keep other Excel workbooks open, or use the PC, while Step 3 or Step 4 runs? (Both versions take over or press keys into whatever Excel or window is in front.)

## C910

### Decision candidates
- Consolidated-workbook column look-up: which cell feeds the IMS Reco, Extra In Purchase and Extra In GSTR2B columns (A-G1/B-G1; A-R80-A-R82/B-R79-B-R81). GST RULE.
- Proof of a good refresh: is 'Info!BB2 not empty/0/error' the test, and should a copy be saved before the test result is used (A-G3/B-G6, A-G2/B-G7, A-R71-A-R73/B-R66-B-R71). The BB2 test is a GST RULE; saving before the check and Skip refresh re-using a failed copy are technical.
- Links cells for file types a client does not have: blank them or keep the template's text (A-G5/B-G2; A-R54-A-R56/B-R54-B-R56). GST RULE.
- Which file feeds a report when a client has several of one type, including browser '(1)' duplicates (A-G6/B-G3, A-G38/B-G48; A-R2, A-R27, A-R57/B-R8, B-R25, B-R56). GST RULE.
- Client identity: spelling variants, Private/Pvt, letter case, and the folder-name length cut that can fold two clients together (A-G8/B-G15/B-G14, A-G35/B-G16, A-G21/B-G24; A-R17, A-R19, A-R35/B-R17, B-R19, B-R32). GST RULE (the client key and folder carry the state code).
- File-name patterns and hyphens (A-G33, A-G34/B-G17, A-G38/B-G47; A-R1-A-R7/B-R5-B-R11). GST RULE.
- State-code fallback and accepted state spellings (A-G36, A-G77/B-G16; A-R21-A-R23/B-R21, B-R22). GST RULE.
- Period and financial-year checks on file names (A-G39/B-G49, B-G23; A-R25/B-R23). GST RULE.
- How 'Missing' and failed-refresh figures appear in the consolidated workbook, and sheet-name matching (A-G17/B-G10, B-G52; A-R83, A-R89/B-R88, B-R90). GST RULE.
- Value types in the consolidated workbook: text figures and Excel error cells (A-G18, A-G79/B-G11; A-R90/B-R88, B-R89). GST RULE.
- Template Links sheet must exist and be checked before a run (A-G4/B-G4; A-R51, A-R58/B-R53, B-R59). Technical, not a GST rule.
- Meaning of Resume and Re-run, minute-based names, and which run folder Re-run uses (A-G10/B-G13, A-G9/B-G27, A-G45; A-R38, A-R43/B-R38, B-R43). Technical, not a GST rule.
- Excel safety under Standard 12: a separate hidden Excel, no take-over of the user's Excel, no keystrokes or dialog clicks into other programs, time limits, and clean-up order (A-G54/B-G63, A-G55/B-G61, B-G62, A-G44/B-G51, A-G50/B-G39, B-G57, A-G28/B-G38, B-G77). Technical, not a GST rule.
- Safe writing: temporary names then rename, partial copies, and not losing the whole consolidated workbook over one value (A-G11/B-G26, A-G12/B-G43, A-G29/B-G43). Technical, not a GST rule.
- Start-up and error visibility without a console, and log location, encoding and growth (A-G22/B-G34, A-G23/B-G33, A-G30/B-G35, A-G56/B-G64, B-G44, A-G52). Technical, not a GST rule (Standard 13).
- Honest counts and statuses: 'Success' with failed reports, Stop reported as complete, problem files shown only as a count (A-G19/B-G19, B-G21, A-G69/B-G70, A-G67/B-G20). Technical, not a GST rule.
- Guards against two runs or two program copies at once (A-G71/B-G74, B-G82). Technical, not a GST rule.
- Settings-file validation (A-G24/B-G36). Technical, not a GST rule.
- Long paths on the office share (A-G41/B-G54). Technical, not a GST rule; the path limits change over time.
- Validation of the cell-mapping configuration (A-G78/B-G78; A-R83-A-R85/B-R82, B-R84). GST RULE.
- Repository confidentiality, Section 10 items D1-D7: the run log and .pyc paths in the shared history of both public repositories, the company-like name and the author's folder path in A, the office share path in two A commit messages, and the author e-mails in the commit metadata. Not a GST rule; any action is the author's decision.

### Flags for other sections
- Rules/feature writers (C3, C5a/C5b): Skip refresh behaves differently in A and B when the folder listing puts a refreshed copy first (see the M5 row of section 9). Relevant to A-R67/A-R69 versus B-R63; both outcomes are wrong, in different ways.
- History writer: the consolidated-workbook mix-up (A-G1/B-G1) became live with the last shared commit 47a1492. The 'contains' look-up from edac6a3 was harmless while every output title had a space; at bd176bf all titles had spaces (RAN git show).
- Decision-list writer: section 9 groups by the higher of the two audit ratings, rounding in-between ratings up. Where the higher rating rests only on one project's part, the row is grouped by the shared part (rows M23, M32, L2, L4, L5, L8, L11, L13 say so). A-G19/B-G19 is the largest rating disagreement (LOW-MEDIUM versus HIGH for identical code).
- Section 10 visibility: the brief says both repositories were public on 21-10-2026. This session could only confirm public on 23-10-2026 (RAN gh api: visibility public, private false, forks 0, only branch main, tips = bf2f1dd and b0e7d0e). GitHub's API keeps no visibility history.
- The narrow-code-page log test (M11/M27/M37) used an ASCII locale (LC_ALL=C, PYTHONUTF8=0, PYTHONCOERCECLOCALE=0) as a stand-in for cp1252. Result: the line with the success symbol was missing from the log file in both projects.
- MAP §3 row A-G19 / B-G19, B-G21, B-G22 was split into H10, M8 and M9 in section 9, so section 9 has 75 rows for MAP's 73. Writers that count shared gaps should use 73 when they mean MAP rows.
- Section 10 item D9 (README example company name, README.md:128-133 since d674386) is in BOTH repositories (byte-identical README). The A audit lists it as a placeholder (X5.2), the B audit as CD3.

### Corrections to the audits or the map
- A audit A-G15 (and MAP row A-G15 / B-G8) says that when a refreshed copy is listed first, 'the run refreshes a copy of a copy'. That holds without Skip refresh. With Skip refresh ticked, A instead takes the listed-first copy as its OWN refreshed copy and reports success without refreshing. The search A:power_query_extractor/core/report_processor.py:139-145 globs '{stem}*.xlsx', which includes the file itself. Since 2aaa2b1, refresh_naming.is_refreshed_copy_of (A:power_query_extractor/core/refresh_naming.py:61-66) accepts any name that starts with that stem and contains '_refreshed_'. RAN C910 probe_step4 s6: A gave status success 'Skipped - Using existing refreshed file' on the copy itself; B made and refreshed '…_Refreshed_…_Refreshed_….xlsx'.
- B audit has no gap for the blank-client double count, although B's code has it identically: a file whose client part is blank is counted 'parsed' and is also put among the variations (B:core/file_parser.py:243, :339; A:core/file_parser.py:210, :311). RAN C910 scripts/blank_client.py: parsed 2, unparsed 1, variations 2 in both. A records it inside A-G19; MAP lists it as shared; the nearest B gap is B-G22.
- B audit B-G19 describes 'the summary workbook hides failures' (Client Status lists only clients that did not raise; failures appear only on the Errors sheet) as part of B's gap. A's code does exactly the same (A:gui/handlers/processing_handler.py:416-429, 443-451 versus B:gui/controllers/processing_controller.py:135-147, 174-181). The two audits rate this identical code LOW-MEDIUM (A-G19) and HIGH (B-G19). This is a rating disagreement, not a code difference.
- Clarification on B-G21 / A-G19: B's per-client 'Folders Created' counts the run folder again for every client (6, 12, 18 for three one-file clients), but A no longer does (6, 11, 16), because A caches the Level-1 folder since 275d9a4 (23-07-2026, A only; A:core/file_organizer.py:150-161). Both are still running totals (RAN C910 probe_steps123 section 9). The B audit's wording 'Fresh mode also counts the run folder again for every client' is true for B only.
- Clarification on A audit X5 item 4: a drive-letter search of A's PROJECT_AUDIT_ANNUALSTATEMENT.md hits lines 5, 231 and 663. Line 231 is not the author's folder; it is an example output path built with the company-like name, so it belongs to item 3 (it is in that item's line list). RAN C910, masked shapes only.

### Questions for the author
- [GST RULE] Should the IMS Reco, Extra In Purchase and Extra In GSTR2B columns of the consolidated workbook show IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1 and ExtraInGSTR2B!S1/T1/U1, as configured, or the Diff GSTR-3B J15/K15/L15 figures they show today in both projects (H1)? Have you noticed these sheets repeating the GSTR-3B difference?
- [GST RULE] Do your real ITC and Sales templates hold paths or file names in the Links cells for Annual Report and IMS (ITC B4/B5/B10/B11) and for Sales, Annual and Sales Reco (Sales B2/B3, B5/B6, B8/B9)? If they do, a client missing that file type keeps the template's values (H5).
- [GST RULE] Is 'Info!BB2 is not empty, not 0 and not an error' the test you want for a successful refresh (H3)? What does BB2 hold in your templates?
- Do the templates' Power Query steps read only the named file, or every file in the linked folder? If they read the whole folder, backups, '(1)' duplicates and partial copies are read too (M36). How to check: open each template's Power Query editor and read the source steps.
- Which Excel version and display language do the office PCs run, and do users keep using the PC during Step 4? (H17, H20)
- Is the company-like name used in the worked examples of A's PROJECT_AUDIT_ANNUALSTATEMENT.md a real client (Section 10, D3)?
- Were both repositories meant to be public (Section 10, 10.1)? Do you regard the office share path in the messages of A's commits 4f84ba9 and 7347383 as confidential (D5)?
