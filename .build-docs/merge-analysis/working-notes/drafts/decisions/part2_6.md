## Part 2 - GST rules applied differently

Every item in this part is a **GST rule**: something that decides a tax figure; which return, table, sheet or cell a figure or file belongs to; a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. When in doubt, a point is treated as a GST rule. These rules are yours alone. Each item shows what A does, what B does, the same made-up example run through both, when each version came in and any hint why, and stops there: **no recommendation, and no view on which version is right.** Where an item also has a purely technical side (a crash, a lost message), that is said as a "technical point" and is not a vote for either version.

Words used in Parts 2 and 6 (explained once here):

- **A** = `AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026; window built with **Tk**, short for Tkinter, Python's built-in window toolkit). **B** = `AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026; window built with **Qt**, through PyQt5, an add-on window toolkit). The names mislead: A is not the older one, and B is not newer or better. Both share their first 12 commits and split after `47a1492` (30-10-2025).
- A **commit** is one saved step in a project's history: a short code such as `5c8ce95`, a date and a message written by whoever made it. "Shared" means the version was already there before the split, so both projects inherited it.
- The four steps of both tools: **Step 1 Setup** (choose the source folder of downloaded returns, the two templates and the target folder), **Step 2 Validation** (the list of clients found, with their files), **Step 3 Processing** (copy each client's files into a folder tree and make its ITC and Sales reports), **Step 4 Extract** (refresh the reports in Excel and read fixed cells into one workbook). **Dry Run** is a preview of what Step 3 would do, without copying anything.
- A **template** is your ITC or Sales Excel workbook, which the tool copies once per client to make that client's report. Each template has a **Links sheet**: the tool writes folder paths and file names into fixed cells of it (for example ITC cell B11 = the client's IMS Reco file name), and the template's **Power Query** (Excel's built-in data-import tool, whose queries live inside the workbook) reads those cells to load the client's files. To **refresh** is to make Excel re-run those queries.
- Folders made by Step 3: a **run folder** `Annual Statement-<DDMMYY HHMM>` per Fresh run (see the modes below); inside it a **client folder** `<client>-<state code>` (for example `ClientA-DL`); inside that a **Version folder** `Version-<DDMMYY HHMM>` for each run; inside that three **category folders** (GSTR-3B, ITC and Sales files).
- Step 3 has three **processing modes**: **Fresh** (the normal one: a new run folder), **Re-Run** (a new Version folder inside the newest run folder) and **Resume** (the same, skipping files already in place). The **share** is the office's shared network drive where the season's files sit.
- The **consolidated workbook** is Step 4's single workbook that collects every client's extracted figures, one sheet per output (for example "Table 8D Limits").
- A file-name **pattern** is the rule the code uses to recognise a downloaded file from its name (technically a "regular expression", a text pattern matched against the name). A **variation** is the code's word for a source file it could not use.
- **RAN** = the real code of both projects was run on the same made-up inputs in a scratch copy and both results are quoted; **READ** = established by reading the code. **HIGH / MEDIUM** = confidence of a statement that could not be run here (this machine has no Excel). **⚠** = a value that can change over time (return and table names, portal file-naming, the state list, template cell positions, Excel limits) - please check it.
- `A:path:line` / `B:path:line` point to a line of A's or B's code. "Section 5, 1.1" and the like point to the full detail in MERGE_COMPARISON.md. All clients and figures below are made up (ClientA, ClientB, `0042`, Zeta Traders and the like); no real client appears.

Answer each item with "A", "B", "Not sure yet" or a short sentence. Items here are numbered P2-1 … for drafting; the final list renumbers them.

### P2-1. IMS Reco files saved with a "(1)" after the date                Part 2 · refs A-R2, B-R7 (with A-R27/B-R25, A-R57/B-R56; A-F7/B-F6; gaps A-G6/B-G3, A-G38/B-G47, B-G48)
What this is about: When the same IMS reconciliation is downloaded twice, the second copy is saved as `ImsReco-<Client A>-Delhi-31032026 (1).xlsx`. Should such a file count as the client's IMS Reco file - and when the plain file and the "(1)" copy are both in the folder, which one should the ITC report's Links cells B10/B11 ⚠ point to?
A's way: Accepts one "(1)" or " (1)" after the 8-digit date; the date part is still the 8 digits (A:utils/constants.py:23). With both files present the client becomes "Has duplicates" and B11 gets the "(1)" copy, because the "first file in name order" rule (same in both) puts a space before a dot (A:core/excel_handler.py:648-651). Came in with `5c8ce95` (23-07-2026, "Detect IMS Reco files with a trailing (N) duplicate suffix"); its message says such names "were not detected: the pattern required exactly 8 digits immediately before the extension" and that "all other types already tolerated the suffix".
B's way: Nothing may follow the date (B:utils/constants.py:23), so the "(1)" copy is a variation ("No matching pattern found"), shown only as a count. This is the original pattern from the first commit `edac6a3` (31-08-2025), never changed in B.
Other choices: Accept the "(1)" copy but link the plain file when both exist; or accept the "(1)" copy only when there is no plain file.
Evidence: RAN through both (MERGE_COMPARISON.md section 5, Category 1 (1.1); section 4, 4.4.1). Folder with `ImsReco-ClientA-Delhi-31032026.xlsx`, its " (1)" copy and ClientA's other five types: A - "Has duplicates", 8 files copied, B11 = `ImsReco-ClientA-Delhi-31032026_(1)`; B - "Complete", 7 copied, B11 = `ImsReco-ClientA-Delhi-31032026`. Only the "(1)" copy present: A - "Complete", B10/B11 filled; B - "Missing 1 files", B10/B11 keep the template's own text. Neither project checks which copy is the newer download. A's tests pin A's form (A:tests/test_file_patterns.py:46-67); 4 of them fail on B's code. Hint why: real IMS files arrived in this form (HIGH, A audit); where the "(1)" came from is not stated (MEDIUM: a browser saving twice).
Recommendation: No recommendation - your call.
Your answer: ____

### P2-2. Showing staff the files the tool could not use, with the naming hint                Part 2 · refs A-R12, B-R12 (with A-R61, B-R94, B-R95; A-F9/B-F9; gaps A-G67/B-G20)
What this is about: For every file whose name matches no pattern, both projects work out the same hint ⚠, for example `ImsReco-ClientName-State-DDMMYYYY.xlsx` for an IMS download saved as "IMS Reco-…" with a space. The question is whether staff ever see the list of unused files and their hints.
A's way: Writes the list into the "Variations" sheet of the Step 3 summary workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx`, with the hint under "Expected Pattern" (list handed over at A:gui/handlers/processing_handler.py:224; sheet A:core/excel_handler.py:554-585). This is the behaviour both had at the split (`47a1492:gui/handlers/processing_handler.py:161`), kept by A.
B's way: Works out the same hint (B:core/file_parser.py:316-334) but hands the summary an empty list (B:gui/controllers/processing_controller.py:73), so the sheet always reads "No file variations found"; on screen only a number ("Variations: n"). Came in with `c258047` (07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"); the message does not mention it (HIGH that the list was lost when the window was rewritten; no reason is given anywhere).
Other choices: Also show the list on screen after the scan (neither does today).
Evidence: RAN through both (section 5, Category 1 (1.6) and Category 8 (8.7)). Source folder with `Claims register ClientA.xlsx`: both work out the IMS hint (the word "claims" contains "ims"); A's Variations sheet has one row - file name · "No matching pattern found" · `ImsReco-ClientName-State-DDMMYYYY.xlsx` · size · "Rename file"; B's sheet says "No file variations found". The hint texts themselves are identical (Part 6, P6-1). Whether files too small or not really Excel are listed is a separate technical difference (section 5, 1.3-1.4).
Recommendation: No recommendation - your call.
Your answer: ____

### P2-3. The example file names shown on Step 1                Part 2 · refs A-R16, A-R147, B-R13, B-R141 (A-F31/B-F33; gaps A-G66/B-G72)
What this is about: Step 1 shows staff how each downloaded return must be named before a scan - the names your team types when they save the season's GSTR-2B, IMS, GSTR-3B, Sales, Sales Reco and Annual Report files ⚠. Which wording should they see?
A's way: The "EXPECTED FILE NAMES" card prints each type with the example text stored next to its pattern: "GSTR-2B Reco - GSTR-2B-Reco-ClientName-State-Period.xlsx", "IMS Reco - ImsReco-ClientName-State-DDMMYYYY.xlsx" and so on (A:gui/tabs/setup_tab.py:477-503; texts A:utils/constants.py:19-59). Came in with `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"); its code comment says the card "can never drift from the regexes [the patterns] the parser [the code that reads file names] actually uses", but it prints the hand-written example texts, not the patterns, so the IMS line does not show the "(1)" form A accepts (P2-1).
B's way: A hand-typed short list without extension or type names: "GSTR-2B-Reco-Client-State-Period", "ImsReco-Client-State-DDMMYYYY" … (B:gui/views/setup_view.py:488-499); the longer texts sit unused in B:utils/constants.py:19-59. The short list is the shared original (`edac6a3`); B's `c258047` (07-02-2026) first showed the longer `.xlsx` texts, and `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections") went back to the short list; no reason is given.
Other choices: Either wording, adding what neither card says today: `.xls` files are also accepted; a hyphen inside a client or state name breaks the reading (Part 6, P6-1); and the "(1)" form if you keep A's way in P2-1.
Evidence: RAN on both real screens (section 5, Category 1 (1.7); section 6.7, G1). Both name the same six types in the same order; A shows `.xlsx` and the type, B neither. Recognition itself does not depend on this card.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-4. The client key handed from Step 2 to Step 3, and where it is shown                Part 2 · refs A-R17, B-R17 (A-F8/B-F8; gaps A-G25, B-G73)
What this is about: Each client is identified by "<client as written in the file names>-<state code>", for example `0042-MH` for a client your office files under the code 0042 in Maharashtra. Both projects build this key the same way at the scan; they differ in how Step 2 hands it back when you tick clients and press Start.
A's way: Rebuilds the key from the text shown in the Step 2 table (A:gui/handlers/client_handler.py:161-173); Tk turns number-looking text into a number, so `0042` comes back as 42 and the key becomes `42-MH`, which does not exist. Step 2 shows the state name only; the key appears in the Step 3 log and as the Step 4 row names. Shared Tk code since `edac6a3` (31-08-2025).
B's way: Stores the key inside each row when the table is filled and hands it back unchanged (B:gui/views/validation_view.py:371, 412-421). Step 2 also shows the key in a "Folder Name" column (:379) - the key itself, not the shortened client folder, so `ClientB Private Limited-MH` is shown while the folder is `ClientB Pvt Ltd-MH` (gap B-G73). Came in with B's window rewrite `c258047` (07-02-2026); the message does not mention it.
Evidence: RAN (section 5, Category 2 (2.1); section 4, 4.4.5). Ticking `0042`, ClientA and ClientE: A hands Step 3 `42-MH`, `ClientA-DL`, `ClientE-DL`; B hands `0042-MH`, `ClientA-DL`, `ClientE-DL`. Start in A with `42-MH` first: "Fatal error: '42-MH'", nothing written for any client, no summary workbook. Technical point only: the key formula and the state codes are the same in both; A's way stops the run for number-like client names. Other number-like names were not run.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-5. How Step 2 shows each client's completeness                Part 2 · refs A-R29, B-R27, A-R138 and B-R134 (the "Missing" column only) (A-F8/B-F8)
What this is about: Step 2 lists each client as "Complete", "Missing <n> files" or "Has duplicates" - the rule that decides this is the same in both (Part 6, P6-2). What differs is how it is shown: the colours, what the "Complete Only" button ticks, and whether the missing return types are named - for example a client whose IMS Reco and Sales Reco files have not come in yet, or one with two GSTR-2B reconciliations for the year.
A's way: Status with an icon ("✅ Complete", "⚠️ Missing 2 files", "⚠️ Has duplicates"); the whole row is tinted pale green for Complete and pale orange for anything else, so "Has duplicates" looks like "Missing" (A:gui/handlers/client_handler.py:37-59); "Complete Only" ticks rows whose status contains "Complete" (:152); the "Missing" column names the types, for example "IMS Reco, Sales Reco" (:31). Shared Tk code (`edac6a3`; tint by each row's status since the shared `1c48722`, 29-10-2025).
B's way: Status without icon; the row's text is green (Complete), orange (Missing) or red-orange (anything else) (B:gui/views/validation_view.py:387-398); "Complete Only" ticks rows whose status equals "Complete" (:446); the "Missing" column shows a number (:377); a label reads "3 clients (1 complete)" (:407-410). Came in with `c258047` (07-02-2026) and `7e09e5c` (08-02-2026); no reason is given.
Other choices: A mix - for example B's three colours with A's named missing types.
Evidence: RAN on both real screens (section 5, Category 5 (5.1); section 6.7, G2 and G3). ClientE (all six types, two GSTR-2B Reco files): A "⚠️ Has duplicates" in the same tint as a client with missing files; B "Has duplicates" in red-orange. ClientB lacking IMS Reco and Sales Reco: A's Missing cell "IMS Reco, Sales Reco"; B's "2". With today's three status texts "Complete Only" ticks the same rows in both. Whether rows start ticked (A none, B all) is a separate technical point.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-6. Shortest allowed client-folder name, and whether the state code may be cut off                Part 2 · refs A-R36, B-R32, B-R138 (with A-R35; A-F6/B-F4; gaps A-G21/B-G24, A-G26)
What this is about: Each client folder is "<client>-<state code>", for example `ClientB Pvt Ltd-MH`, shortened to a maximum length set on Step 1 (default 35). The name builder is the same in both: if the maximum is so small that fewer than 6 characters would remain for the client, it cuts the whole name, state code included - so one client's Delhi and Maharashtra registrations can land in one folder. Which lowest value should be allowed, and may the state code ever be cut off? (How the box reacts to typed text, and B's greyed-out look, are technical and decided in P1-5.)
A's way: A number box whose arrows run 15-100 and which is always active, but a typed value is not checked and is used as typed (A:gui/tabs/setup_tab.py:434-443; read at A:gui/handlers/processing_handler.py:194); typed text such as "abc" stops Step 3 with "Fatal error". The shared box also refused typed values outside 10-100; A removed that check in `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"), and the message does not mention it.
B's way: A Qt number box held within 15-100 (B:gui/views/setup_view.py:374-382): a typed 8 is refused, a stored 8 becomes 15. It is greyed out unless "Include client name" is ticked, but its value is used anyway (B:gui/pyqt_main_window.py:203). Range 5-100 in `c258047` (07-02-2026), 15-100 since `7e09e5c` (08-02-2026); no reason is given.
Other choices: Allow any value but never cut the state code (that would change the shared name builder).
Evidence: RAN (section 5, Category 6 (6.1); section 4, 4.4.3). Maximum 8 (possible only in A): ClientA/Delhi and ClientA/Maharashtra both get the folder `ClientA-` (state code gone, one shared folder). B's lowest value 15: `ClientA Private Limited`/Maharashtra → `ClientA Pvt -MH`. A three-letter made-up state code (P6-3) is lost already at 9. Technical point: A's box also accepts text that stops the run.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-7. Names of the three category folders - also written into the Links cells ⚠                Part 2 · refs A-R39, A-R41, A-R54, A-R55, B-R35, B-R37, B-R54, B-R55 (A-F13/B-F12, A-F14/B-F13, A-F15/B-F14; gaps A-G41/B-G54)
What this is about: Inside each Version folder the tool makes three folders: GSTR-3B exports; ITC files (GSTR-2B Reco and IMS Reco); Sales files (Sales and Sales Reco). Their full paths are written into the templates' Links cells - ITC B2, B7, B10 and Sales B2, B8 ⚠ - which the templates' queries read. So the names matter to your templates, not only to staff opening the folders.
A's way: `GSTR3B`, `ITC`, `Sales` (A:utils/constants.py:85-95). Came in with `7347383` (26-07-2026, "Shorten category folder names to fit Excel's path limit"). Its message reports a real run of 3,030 output files on the office share whose worst path was 270 characters, with 1,094 files (36%) over Excel's limit of about 218 characters ⚠; it warns "the margin is thin" and that a path-length guard "remains the durable fix". The code comment beside the names gives modelled figures (235 / 215) that assume a planned shorter share root.
B's way: `GSTR-3B Exports`, `Other ITC related files`, `Sales related files` (B:utils/constants.py:85-89) - the shared original names from `edac6a3` (31-08-2025), never changed in B.
Other choices: A third set of names. Whatever you choose, batches already on the share carry the long names (every B batch, and A batches before 26-07-2026), so Step 4 of the new tool may meet both.
Evidence: RAN (section 5, Category 6 (6.2, 6.4) and Category 8 (8.3); section 4, 4.4.2). ClientA, run at 21-10-2026 10:30: ITC B2 = `…\Version-211026 1030\GSTR3B` (A) / `…\Version-211026 1030\GSTR-3B Exports` (B); B7 and B10 = `…\ITC` / `…\Other ITC related files`; Sales B2 and B8 = `…\Sales` / `…\Sales related files`. Which file type goes into which folder is the same in both (P6-4). NOT SURE whether your templates' queries take every folder only from these cells or also name a folder themselves - the templates are in neither repository. How to check: open each template in Excel, Data > Get Data > Launch Power Query Editor, open Advanced Editor for each query and search for "GSTR-3B Exports", "Other ITC related files" and "Sales related files".
Recommendation: No recommendation - your call.
Your answer: ____

### P2-8. What must be chosen before a Dry Run                Part 2 · refs A-R51, A-R134, A-R139, B-R48, B-R133, B-R135 (A-F11/B-F11)
What this is about: Before you process 40 clients you may press Dry Run to preview the copies and the "Would create: ITC Report / Sales Report" lines. Must both templates and the target folder be chosen before the preview runs?
A's way: The same check as Start: both template files chosen and present, and a target typed - otherwise "Please select both ITC and Sales template files" and no preview (A:gui/handlers/file_handler.py:132-151, called at A:gui/handlers/processing_handler.py:40). This is the shared check (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`, 31-08-2025).
B's way: Only a ticked client is needed (B:gui/pyqt_main_window.py:326-343); Start still checks templates and target (:345-361). B's Dry Run came without the check in `c258047` (07-02-2026); no reason is given.
Evidence: RAN on both real screens (section 5, Category 8 (8.1); section 6.7, G5). No templates chosen, three clients ticked, Dry Run: A shows the error box and no preview; B prints the whole preview, promising "📊 Would create: ITC Report" and "💰 Would create: Sales Report" for every client, ending "No files were actually moved or created". Start behaves alike in both.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-9. Which template file types can be picked                Part 2 · refs A-R52, B-R48, B-R51 (gaps A-G42/B-G30, B-G66)
What this is about: Your ITC and Sales templates may be plain workbooks (`.xlsx`), Excel templates (`.xltx`, `.xltm`) or macro workbooks (`.xlsm`). The Browse button decides which types staff can pick with the mouse. Whatever is picked, the report is always saved under an `.xlsx` name (same in both). Listed here "when in doubt": the type list is tagged a GST rule in neither audit, only through the report name, which carries the state code.
A's way: Browse offers `*.xltx *.xlsx` (A:gui/handlers/file_handler.py:31, 42) - the shared original; A's Step 1 checklist (added in `322717f`, 25-07-2026) marks other types "Not an Excel template", but Start accepts any existing file.
B's way: Browse offers `*.xlsx *.xlsm`, label "(.xlsx or .xlsm)" (B:gui/views/setup_view.py:221, 612, 619); the filter came with `c258047` (07-02-2026) and the label with `7e09e5c` (08-02-2026); no reason is given. Start accepts any existing file.
Other choices: Offer all four Excel types in the picker.
Evidence: RAN (section 5, Category 8 (8.2)). `.xlsx`, `.xltx` and `.xlsm` templates are accepted by Start in both; so are `.xlsb` and even `.txt` (A's checklist flags them; nothing stops them). NOT SURE whether a report made from an `.xlsm` or `.xltx` template and saved as `.xlsx` opens normally (MEDIUM; how to check: run one client on Windows with such a template and open the report). It helps to know which type your templates are.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-10. A Links cell inside merged cells, when a report is made without Excel ⚠                Part 2 · refs A-R58, B-R53 (with A-R96, B-R97; A-F14/B-F13)
What this is about: On a PC where Python cannot drive Excel, both projects make the reports with **openpyxl** (a Python library that writes Excel files without Excel; when it saves, it drops the template's Power Query and any drawn shapes such as text boxes and buttons). On that path a Links cell that lies inside merged cells (several cells joined into one, here A2:B2 with B2 the mapped cell) cannot be written. Should the report still be made with that cell empty, or fail?
A's way: Skips that one cell without a word and counts the report as created (a catch-all that swallows any error, A:core/excel_handler.py:297-303). Shared code since `edac6a3`. In A this path runs wherever **pywin32** (the add-on that lets Python drive Excel on Windows) cannot be loaded.
B's way: The whole report fails and the unfilled template copy stays on disk under the report's name (B:core/excel_handler.py:246-252, 258). Came in with `c258047` (07-02-2026), whose message lists "Fixed all 26 bare except: clauses with specific exception types" (a bare except is the catch-all described under A's way). B's window cannot start without pywin32, so B never reaches this path.
Other choices: Stop that report and name the cell that could not be written. This item matters only if the new tool keeps a way of making reports without Excel (a Part 8 decision).
Evidence: RAN (section 5, Category 8 (8.4); section 4, 4.4.7). B2 inside merged A2:B2: A - B2 empty, B4, B5, B11 written, "created", no message; B - fails, log "'MergedCell' object attribute 'value' is read-only", B4, B5, B11 still hold the template's placeholders (the stand-in text the test template had in those cells). When the merge starts at B2 (B2:C2) both write it. Through Excel both are identical. Whether your templates have merged Links cells is not known.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-11. Which template a batch uses if the Step 1 box is changed mid-run; order of clients                Part 2 · refs A-R60, B-R50, B-R60, B-R134 (A-F15/B-F14; gaps A-G14/B-G28)
What this is about: During a Step 3 run for 40 clients someone notices that last year's ITC template is selected and picks the new one on Step 1. Should the remaining clients use the new template, or the whole batch the one chosen when Start was pressed? The two projects also process clients in a different order.
A's way: Re-reads both template paths (and the per-client folder-name choices) from the screen for every client (A:gui/handlers/processing_handler.py:193, 301, 310, 332). Order: the order in which the scan first met each client in the alphabetically sorted file list (A:gui/handlers/client_handler.py:30, 165). Both are the shared Tk code (`edac6a3` … `47a1492`).
B's way: Takes one snapshot of all settings when Start is pressed (B:gui/pyqt_main_window.py:352, 363-371). Order: sorted by client key, unless you click a column heading - then the order on screen (B:gui/views/validation_view.py:324, 367, 412-421). Both came with `c258047` (07-02-2026); no reason is given.
Other choices: B's snapshot with A's order, or A's live reading with B's order.
Evidence: RAN (section 5, Category 8 (8.6); section 6.7, G6). Three clients; the ITC template changed from "v1" to "v2" after ClientA: A - ClientA from v1, ClientB and ClientD from v2 (two templates in one batch, no warning); B - all three from v1. Clients `Zeta Traders` (first file an Annual Report) and `Alpha Traders`: A does Zeta Traders first, B Alpha Traders first. B's re-sorting after a heading click was READ only (HIGH), not run. The order matters only when two clients share a folder or for the running totals in the per-client text reports. Both make both reports for every client (P6-5).
Recommendation: No recommendation - your call.
Your answer: ____

### P2-12. How Step 4 recognises the ITC and Sales reports                Part 2 · refs A-R59, B-R51, B-R61 (A-F20/B-F18; gaps A-G47/B-G46 in part)
What this is about: Step 4 lists, per client, the reports Step 3 made: `ITC_Report_<client>_<state code>_<DDMMYY_HHMM>.xlsx` and `Sales_Report_…` (names identical in both, P6-4). The two projects differ only for a report someone renamed by hand with a capital `.XLSX`.
A's way: Ticks a report when its name, in lower case, starts `itc_report_` / `sales_report_` and ends `.xlsx` (A:power_query_extractor/gui/extractor_window.py:20-42), then finds the file with a **wildcard search** (`*` stands for any text) `ITC_Report_*.xlsx` (A:power_query_extractor/core/report_processor.py:111-125). Came with `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor" - cache = keep a result to reuse it, level-1 folder = the run folder, PQ = Power Query), for speed on the network share; the message calls the new helpers (two small pieces of code that list each folder once) "proven equivalent to the original glob logic" (glob = the wildcard search), which holds on Windows only.
B's way: Uses the wildcard search `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` both for the list and for the file (B:gui/pyqt_main_window.py:461-462; B:power_query_extractor/core/report_processor.py:72-89), from B's own Step 4 page in `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking").
Evidence: RAN on Linux (the test machine's operating system, which, unlike Windows, treats `.XLSX` and `.xlsx` as different) (section 5, Category 8 (8.5)). A report renamed `…_0335.XLSX`: A lists and ticks it, then reports "No ITC report found"; B does not list it. On Windows both find it (HIGH: Windows searches ignore letter case; not run). The tool itself always writes a lower-case `.xlsx`, so only hand-renamed files are affected.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-13. "Skip refresh": which existing refreshed copy may be reused                Part 2 · refs A-R68, A-R69, B-R63 (with B-R136; A-F22/B-F21; gaps A-G2/B-G7, A-G15/B-G8, A-G16/B-G9)
What this is about: Step 4 copies each report to a **refreshed copy** named with a **suffix** (default `_Refreshed_{timestamp}`, where `{timestamp}` becomes the date and time of the copy, `DDMMYY_HHMM`; editable on Step 4), refreshes that copy and reads its figures. With "Skip refresh" ticked, an existing refreshed copy is reused instead of refreshing again. So this decides which file the consolidated figures come from: today's refresh, an older copy, or - in one case below - the never-refreshed report.
A's way: Recognises copies by the suffix currently in the box, plus the old marker `_refreshed_` (A:power_query_extractor/core/refresh_naming.py:14-66; reuse at A:power_query_extractor/core/report_processor.py:132-159). Came in with `2aaa2b1` (25-07-2026, "Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one"); its message says that with a customised suffix "'skip refresh' found nothing and re-refreshed every report instead of reusing the existing files - the opposite of what the option is for".
B's way: Looks only for `<report name>_Refreshed_*.xlsx`, whatever the box says (B:power_query_extractor/core/report_processor.py:96-115). This is the shared original (`edac6a3`, 31-08-2025); the editable suffix box came in the shared `bd176bf` (30-10-2025) without changing the search.
Other choices: Reuse a copy only after checking that its own refresh had passed (neither project checks today).
Evidence: RAN (section 5, Category 9 (9.4)). Suffix `_PQ_{timestamp}` with yesterday's `_PQ_` copy present: A reuses it; B refreshes again (about 39 s of Excel per report at the default wait). Suffix `_R{timestamp}`, original report only: A takes the ORIGINAL Step 3 report as a refreshed copy (`_r` is inside `_report_`) and reports success without refreshing; B refreshes. Suffix `_Refreshed` without a time stamp: A reuses the copy; B refreshes and overwrites it. With the default suffix both behave alike. In both, a copy left by a FAILED refresh is reused as a success (shared gap A-G2/B-G7). How the Step 4 status column shows refreshes is a separate technical point.
Recommendation: No recommendation - your call.
Your answer: ____

### P2-14. Consolidated workbook: one row per client, or one row per Version folder                Part 2 · refs A-R87, A-R91, B-R87, B-R90 (with A-R65, B-R61; A-F24/B-F23, A-F20/B-F18)
What this is about: Step 4 writes every client's figures into the consolidated workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, one sheet per output (for example "Reco Diff" or "Table 6J Match"). If a client was processed twice in one batch (a Re-Run adds a second Version folder), should its figures appear once or once per Version folder? Listed here "when in doubt": it decides how many rows of figures one client has (section 3 counts the row layout as technical).
A's way: One row per client, labelled with the client folder (for example `ClientA-DL`), from the client's newest Version folder only; "Total Clients" counts clients (A:power_query_extractor/gui/extractor_window.py:649-655). Shared Tk extractor (`edac6a3`).
B's way: One row per ticked Version folder, labelled `<client folder> / <version folder>` (label B:gui/views/extraction_view.py:384); "Total Clients" counts Version folders (B:gui/pyqt_main_window.py:518-528). Came with `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking"); no reason is given.
Other choices: One row per client, from a Version folder you choose.
Evidence: RAN (section 5, Category 9 (9.3) and Category 11 (11.3)). ClientA with two Version folders, ClientB with one: A gives two rows (`ClientA-DL`, `ClientB-MH`), "Total Clients" 2; B, after Select All, gives three rows - `ClientA-DL / Version-311225 1000`, `ClientA-DL / Version-301225 0900`, `ClientB-MH / Version-311225 1000` - and "Total Clients" 3. The cells read and the figures in each row follow the same rules in both (P6-6, P6-7). Which batch and which Version folders Step 4 offers is a separate technical item.
Recommendation: No recommendation - your call.
Your answer: ____

## Part 6 - Rules both projects share

These rules are the same in both projects (checked function by function, and run where possible: MERGE_COMPARISON.md section 5), so there is nothing to choose between A and B. The question is only whether each rule is **still correct today** for your practice - the portal's file names, your templates' cells, the state list, the return tables. Rows marked ⚠ hold a value that can change over time. Each table is one item: answer **"all correct"**, or **"all correct except …"** with the row numbers and what is wrong (your words are enough; we will turn them into rules for the new tool). "Not sure yet" is fine. A few rows are **dormant**: written in both projects but never used; for those, say "drop" or "keep" in your except-list. No recommendation in this part.

### P6-1. File-name patterns and how a name is read ⚠                Part 6 · refs A-R1, A-R3, A-R4, A-R5, A-R6, A-R7, A-R11, A-R12 (texts), B-R5, B-R6, B-R8, B-R9, B-R10, B-R11, B-R12 (texts) (A-F7/B-F6; gaps A-G33, A-G34/B-G17, A-G38/B-G47)
What this is about: How the tool recognises each downloaded return from its file name, for example `GSTR3B-<Client A>-Maharashtra-Apr.xlsx` as ClientA's April GSTR-3B. Five of the six patterns are identical; the sixth (IMS Reco) differs and is item P2-1.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | GSTR-2B Reco: `GSTR-2B-Reco-<client>-<state>-<period>.xlsx`; the period may contain hyphens (e.g. `Apr25-Mar26`) | A:utils/constants.py:14-21 | B:utils/constants.py:14-21 | ⚠ | ___ |
| 2 | GSTR-3B Export: `GSTR3B-<client>-<state>-<month>.xlsx`; a leading "(1) " is accepted; `GSTR-3B-…` (with a hyphen) is not recognised | A:utils/constants.py:30-37 | B:utils/constants.py:30-37 | ⚠ | ___ |
| 3 | Sales: `Sales-<client>-<state>-<start month>-<end month>.xlsx` (exactly four parts); a file named `Sales-Reco-…` is read as a Sales file of a client called "Reco" | A:utils/constants.py:38-45 | B:utils/constants.py:38-45 | ⚠ | ___ |
| 4 | Sales Reco: `SalesReco-<client>-<state>-<period>.xlsx` | A:utils/constants.py:46-53 | B:utils/constants.py:46-53 | ⚠ | ___ |
| 5 | Annual Report: `AnnualReport-<client>-<state>-<year>.xlsx` | A:utils/constants.py:54-61 | B:utils/constants.py:54-61 | ⚠ | ___ |
| 6 | All patterns: any letter case; `.xlsx` or `.xls` only (an `.xlsm` file is found but never recognised); a trailing " (1)" on the five types above simply becomes part of the period or month text (e.g. month "Apr (1)") | patterns above | patterns above | ⚠ | ___ |
| 7 | Patterns tried in a fixed order (GSTR-2B Reco, IMS Reco, GSTR-3B, Sales, Sales Reco, Annual Report), first match wins, parts trimmed. A hyphen inside a client or state name shifts the parts (a client written `Alpha-Beta Traders` is mis-read) | A:core/file_parser.py:82-134 | B:core/file_parser.py:78-124 | | ___ |
| 8 | A name that matches no pattern is listed "No matching pattern found"; a match with an empty client or state "Could not parse filename" | A:core/file_parser.py:130-132, 209-220 | B:core/file_parser.py:120-122, 245-252 | | ___ |
| 9 | Hint for an unrecognised name (whether staff see it: P2-2): name contains "gstr" and "2b" → GSTR-2B Reco form; "ims" → IMS Reco form; "gstr" and "3b" → GSTR-3B form; "sales" and "reco" → Sales Reco form; "sales" → Sales form; "annual" → Annual Report form; otherwise "Check file naming convention" | A:core/file_parser.py:288-306 | B:core/file_parser.py:316-334 | ⚠ | ___ |

Evidence: RAN (section 5, Category 1, rows A-R1, A-R3 to A-R7, A-R11, A-R12): 15 test names give the same results in both apart from the IMS "(1)" names; the hint function is identical. The hyphen and `Sales-Reco-` cases are shared gaps A-G33/A-G34 = B-G17.
Recommendation: No recommendation - your call.
Your answer: ____

### P6-2. Expected return files and completeness ⚠                Part 6 · refs A-R26, A-R27, A-R28, A-R29 (status rule), A-R30, B-R24, B-R25, B-R26, B-R27 (status rule), B-R28 (A-F8/B-F8)
What this is about: Which return files every client is expected to have, and how Step 2 decides "Complete" - for example whether a client with twelve monthly GSTR-3B files but two GSTR-2B reconciliations is complete.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | Six types expected for every client: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report | A:utils/constants.py:68-75; A:core/file_parser.py:252-260 | B:utils/constants.py:68-75; B:core/file_parser.py:280-288 | ⚠ | ___ |
| 2 | GSTR-3B Export may have any number of files ("monthly filings - unlimited"); any other type with two or more files gets the remark "Multiple <type> (<n> files)" | A:core/file_parser.py:262-271 | B:core/file_parser.py:290-299 | ⚠ | ___ |
| 3 | Completeness % = types found ÷ 6 × 100, one decimal (it counts types, not files) | A:core/file_parser.py:273-275 | B:core/file_parser.py:301-303 | | ___ |
| 4 | Status: "Complete" (nothing missing, no duplicates) / "Missing <n> files" (wins over duplicates) / "Has duplicates" (how it is shown on screen: P2-5) | A:core/file_parser.py:277-286 | B:core/file_parser.py:305-314 | | ___ |
| 5 | Dormant: a check that calls Annual Report and GSTR-3B Export "critical" types; nothing ever calls it - say "drop" or "keep" | A:core/file_parser.py:338-371 (list :359) | B:core/file_parser.py:366-399 (list :387) | ⚠ | ___ |

Evidence: RAN (section 5, Category 5): identical function in both; for example a client with two GSTR-2B Reco files gets "Multiple GSTR-2B Reco (2 files)" in both, and completeness 100.0 / 33.3 / 16.7 come out alike.
Recommendation: No recommendation - your call.
Your answer: ____

### P6-3. State codes, the client key, GSTIN/PAN and the period ⚠                Part 6 · refs A-R17 (scan), A-R21, A-R22, A-R23, A-R24, A-R25, A-R35, B-R17 (scan), B-R21, B-R22, B-R23, B-R32 (gaps A-G8/B-G15, A-G35, A-G36, A-G77/B-G16, A-G39/B-G49, B-G23)
What this is about: How the state part of a file name becomes a code (`Maharashtra` → `MH`), how one client is told apart from another, and what the tool does - and does not do - with GSTIN, PAN and the period in the file names.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | State name → two-letter code from a fixed table of 39 names giving 36 codes (e.g. delhi DL, maharashtra MH, tamil nadu TN, odisha/orissa OD, puducherry/pondicherry PY, ladakh LA); the text is trimmed and lower-cased first; letter codes only - the numeric state codes are not used anywhere | A:utils/constants.py:251-292; A:utils/helpers.py:592-606 | B:utils/constants.py:234-275; B:utils/helpers.py:485-499 | ⚠ | ___ |
| 2 | A state not in the table gets a made-up code: two or more words → first letters of the first two (`Unknown Place` → UP, the same code as Uttar Pradesh); one word → its first three characters in capitals (`Tamilnadu` → TAM), taken from the untrimmed text (" Xyz" → " XY") | A:utils/helpers.py:608-618 | B:utils/helpers.py:501-511 | ⚠ | ___ |
| 3 | Client key = "<client exactly as written in the file name>-<state code>": spelling, spaces and capitals must match, so `ClientB Pvt Ltd` and `ClientB Private Limited` are two clients (how Step 2 hands the key back: P2-4) | A:core/file_parser.py:226-236 | B:core/file_parser.py:258-268 | | ___ |
| 4 | The CODE goes into the client key, the client folder and the report names; the state TEXT as written goes into the per-client text report, the Step 2 table and the summary workbook | A:core/file_parser.py:229-236; A:utils/helpers.py:628-631; A:gui/handlers/processing_handler.py:293 | B:core/file_parser.py:261-268; B:utils/helpers.py:521; B:gui/controllers/processing_controller.py:113 | | ___ |
| 5 | Client folder = client name with whole words "Private" → "Pvt" and "Limited" → "Ltd", then "-<state code>", cut to the maximum length; if fewer than 6 characters would remain for the client, the whole name is cut and the state code is lost (the maximum itself: P2-6) | A:utils/helpers.py:620-644 | B:utils/helpers.py:513-537 | | ___ |
| 6 | No GSTIN or PAN is read, checked or worked out anywhere | (none - searched) | (none - searched) | | ___ |
| 7 | The period parts of the names (GSTR-2B and Sales Reco period, IMS 8-digit date, GSTR-3B month, Sales start and end month, Annual Report year) are kept as plain text and never checked: no financial-year or tax-period logic; an impossible date such as `99999999` is accepted; files of different years for one client are grouped together | A:core/file_parser.py:97-106 | B:core/file_parser.py:93-102 | ⚠ | ___ |

Evidence: RAN (section 5, Categories 2-4 and 6): 39 names and 36 codes in both; `Delhi`, `" delhi "`, `Tamil Nadu`, `Orissa`, `Ladakh` give DL, DL, TN, OD, LA in both; the name builder gives `ClientB Pvt Ltd-MH` in both. Shared gaps: one client spelled several ways becomes several incomplete clients (A-G35 = B-G16); two different clients can share one folder (A-G8 = B-G15); no period check (A-G39 = B-G49, B-G23).
Recommendation: No recommendation - your call.
Your answer: ____

### P6-4. Folder, file and report names                Part 6 · refs A-R32, A-R33, A-R38, A-R41 (type → folder), A-R46, A-R48, A-R59 (names), A-R61 (name), A-R62, A-R86 (name), B-R30, B-R31, B-R34, B-R37 (type → folder), B-R40, B-R43 (backup name), B-R51, B-R57, B-R86 (name), B-R94 (name) (A-F13/B-F12, A-F16/B-F15; gaps A-G9/B-G27, A-G76/B-G81)
What this is about: The names of everything Step 3 and Step 4 create on the share, which staff open during the annual return season - apart from the three category folder names, which differ (P2-7).
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | Run folder `Annual Statement-<DDMMYY HHMM>` inside the target folder; a second Fresh run in the same minute reuses it | A:utils/constants.py:82; A:core/file_organizer.py:163-196 | B:utils/constants.py:82; B:core/file_organizer.py:146-179 | | ___ |
| 2 | Client folder `<client>-<state code>` (rule P6-3 row 5) | A:utils/helpers.py:620-644 | B:utils/helpers.py:513-537 | | ___ |
| 3 | A new `Version-<DDMMYY HHMM>` folder for every run, in every mode | A:utils/constants.py:84; A:core/file_organizer.py:102-108 | B:utils/constants.py:84; B:core/file_organizer.py:98-104 | | ___ |
| 4 | Time stamps are clock times, day first, two-digit year, to the minute (`211026 1030`; with `_` inside file names) - not tax periods | A:utils/helpers.py:25-35 | B:utils/helpers.py:23-33 | | ___ |
| 5 | Which file goes where: GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder itself | A:core/file_organizer.py:302-315 | B:core/file_organizer.py:285-298 | ⚠ | ___ |
| 6 | Copies keep the source name, cleaned (for example spaces → `_`; how "Private" / "Limited" are shortened differs between A and B and is decided in P1-10); after a successful copy the cleaned name is what the Links cells get; a skipped or failed file keeps its original name there | A:core/file_organizer.py:233-246, 270-274 | B:core/file_organizer.py:216-229, 253-257 | | ___ |
| 7 | Reports `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` and `Sales_Report_…` in the Version folder | A:gui/handlers/processing_handler.py:291-298, 355-356 | B:gui/controllers/processing_controller.py:111-130, 220 | | ___ |
| 8 | A file about to be overwritten in a Fresh run is first kept as `<name>_backup_<DDMMYY_HHMM><ext>` in the same folder | A:utils/helpers.py:217-234 | B:utils/helpers.py:187-204 | | ___ |
| 9 | Per-client text report `_Organization_Report_<client>_<state>.txt` in the Version folder | A:core/file_organizer.py:317-379 (name :330) | B:core/file_organizer.py:300-362 (name :313) | | ___ |
| 10 | Step 3 summary workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` in the run folder | A:gui/handlers/processing_handler.py:463 | B:gui/controllers/processing_controller.py:188 | | ___ |
| 11 | Consolidated workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` - year first, unlike every other name (where it is saved differs: a technical item) | A:power_query_extractor/core/data_consolidator.py:176 | B:power_query_extractor/core/data_consolidator.py:164 | | ___ |

Evidence: RAN (section 5, Categories 6-8 and 11): `Annual Statement-211026 1030`, `Version-211026 1030`, `…_backup_211026_1030.xlsx` and `ITC_Report_ClientA_DL_211026_1030.xlsx` in both (the cleaned client name inside copy and report names can differ through the "Private"/"Limited" rule, a technical item). Shared gaps: a repeat in the same minute overwrites earlier output (A-G9 = B-G27); one run is named by two separate clock readings (A-G76 = B-G81).
Recommendation: No recommendation - your call.
Your answer: ____

### P6-5. The template Links-sheet cells ⚠                Part 6 · refs A-R54, A-R55 (cells), A-R56, A-R57, A-R58 (Excel path), A-R60 (both reports), B-R50, B-R53 (Excel path), B-R54, B-R55, B-R56 (A-F14/B-F13, A-F15/B-F14; gaps A-G4/B-G4, A-G5/B-G2, A-G6/B-G3)
What this is about: Which cell of each template's Links sheet receives which folder or file name - the hand-over between the tool and your templates' queries. Only the folder text differs between the projects (P2-7) and, for IMS, which file is named (P2-1); the cells and rules below are the same.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | ITC report, sheet `Links`: B2 GSTR-3B folder (always); B4 Version folder and B5 Annual Report file name; B7 ITC folder and B8 GSTR-2B Reco file name; B10 ITC folder and B11 IMS Reco file name; B3, B6, B9 never written | A:utils/constants.py:102-114; A:core/excel_handler.py:603-656 | B:utils/constants.py:96-108; B:core/excel_handler.py:540-588 | ⚠ | ___ |
| 2 | Sales report, sheet `Links`: B2 Sales folder and B3 Sales file name; B5 Version folder and B6 Annual Report file name; B8 Sales folder and B9 Sales Reco file name; B4, B7 never written | A:utils/constants.py:115-125; A:core/excel_handler.py:657-708 | B:utils/constants.py:109-119; B:core/excel_handler.py:590-622 | ⚠ | ___ |
| 3 | Form: full folder paths with `\`; file names without `.xlsx`; written as text | A:utils/helpers.py:418-422, 553-555 | B:utils/helpers.py:311-315, 446-448 | | ___ |
| 4 | A client lacking a file type: that cell is not written, so the template's own text stays in it (e.g. a client without an Annual Report leaves ITC B4/B5 as the template has them) | A:core/excel_handler.py:189-197, 297-303 | B:core/excel_handler.py:159-167, 246-252 | | ___ |
| 5 | Several files of one type: only the first in sorted full-path order is named (of `…-FY2025-26 (1).xlsx` and `…-FY2025-26.xlsx`, the "(1)" file comes first); the others are copied but never linked | A:utils/helpers.py:361; A:core/excel_handler.py:628-630 (and the five similar lines to :698-700) | B:utils/helpers.py:262; B:core/excel_handler.py:561-562 (and the five similar lines to :617-618) | | ___ |
| 6 | Finding the sheet when Excel makes the report: `Links` in any letter case; else the first sheet whose name contains "link"; else a NEW empty `Links` sheet is added - and the report still counts as made. Without Excel (openpyxl): only `Links` in any letter case; if there is none, nothing is written and the report still counts as made | A:core/excel_handler.py:162-183, 285-310 | B:core/excel_handler.py:131-153, 233-258 | ⚠ | ___ |
| 7 | Both an ITC report and a Sales report are made for every ticked client, whatever files it has (a client with only GSTR-3B files still gets a Sales report with no Links value written) | A:gui/handlers/processing_handler.py:296-411 | B:gui/controllers/processing_controller.py:116-130, 204-239 | | ___ |

Evidence: RAN (section 5, Category 8 (8.3), rows A-R56, A-R57): with a made-up template holding "PLACEHOLDER-B3" etc., B3 keeps its placeholder and file names carry no extension in both; ClientD with two GSTR-2B Reco files gets B8 = `…-FY2025-26_(1)` in both; ClientD with only GSTR-3B and GSTR-2B files gets only ITC B2, B7, B8 and no Sales cell, in both. Shared gaps: no error for a template without a Links sheet (A-G4 = B-G4); missing types keep template text (A-G5 = B-G2); only the first file is linked (A-G6 = B-G3).
Recommendation: No recommendation - your call.
Your answer: ____

### P6-6. Extraction cells and the consolidated-workbook columns ⚠                Part 6 · refs A-R74, A-R75, A-R76, A-R77, A-R78, A-R79, A-R80, A-R81, A-R82, A-R83, A-R85, A-R88, B-R73, B-R74, B-R75, B-R76, B-R77, B-R78, B-R79, B-R80, B-R81, B-R82, B-R83, B-R84 (A-F23/B-F22, A-F24/B-F23; gaps A-G1/B-G1, A-G17/B-G52, A-G78/B-G78)
What this is about: After the refresh, Step 4 reads 22 fixed cells from each client's refreshed ITC and Sales reports into the consolidated workbook - your GSTR-3B and GSTR-1 differences, turnover, the reconciliation differences, the GSTR-9 Table 6J and Table 8D checks, the IMS figures and the extra-in-purchase / extra-in-GSTR-2B figures. The file that holds this list (`power_query_extractor/config/cell_mappings.py`) is byte-for-byte the same in both projects, so every line reference below is the same in A and B. **A technical fact, the same in both: of the 22 configured columns, 9 - the IGST, CGST and SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" (rows 7-9) - show the `Diff GSTR-3B` J15/K15/L15 values instead of their configured cells.** Which cells are right is your call.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Output sheet ← report, input sheet and cells (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | "GSTR-3B Difference" ← ITC report, sheet `Diff GSTR-3B`: J15 → Diff IGST, K15 → Diff CGST, L15 → Diff SGST | A:power_query_extractor/config/cell_mappings.py:35-58 | B: same file and lines | ⚠ | ___ |
| 2 | "GSTR-1 Difference" ← Sales report, sheet `Diff - GSTR-1`: N1 → Diff Taxable Value, O1 → Diff IGST, P1 → Diff CGST, Q1 → Diff SGST | A:…/cell_mappings.py:61-89 | same | ⚠ | ___ |
| 3 | "Turnover GSTR-1" ← Sales report, sheet `Diff - GSTR-1`: I1 → Taxable Value | A:…/cell_mappings.py:92-102 | same | ⚠ | ___ |
| 4 | "Reco Diff" ← ITC report, sheet `Reco`: C37 → IGST Diff, D37 → CGST Diff, E37 → SGST Diff | A:…/cell_mappings.py:105-127 | same | ⚠ | ___ |
| 5 | "Table 6J Match" ← ITC report, sheet `T6 of R9` (by its name, Table 6 of GSTR-9): K29 | A:…/cell_mappings.py:130-140 | same | ⚠ | ___ |
| 6 | "Table 8D Limits" ← ITC report, sheet `T8 of R9` (by its name, Table 8 of GSTR-9): G6 | A:…/cell_mappings.py:143-153 | same | ⚠ | ___ |
| 7 | "IMS Reco" ← ITC report, sheet `IMS_Reco`: configured AA1 → IGST, AD1 → CGST, AG1 → SGST - **today these three columns show J15/K15/L15 instead** | A:…/cell_mappings.py:156-178 | same | ⚠ | ___ |
| 8 | "Extra In Purchase" ← ITC report, sheet `ExtraInPurchase`: configured O1 → IGST, P1 → CGST, Q1 → SGST - **today show J15/K15/L15 instead** | A:…/cell_mappings.py:181-203 | same | ⚠ | ___ |
| 9 | "Extra In GSTR2B" ← ITC report, sheet `ExtraInGSTR2B`: configured S1 → IGST, T1 → CGST, U1 → SGST - **today show J15/K15/L15 instead** | A:…/cell_mappings.py:206-228 | same | ⚠ | ___ |
| 10 | Input sheet names must match exactly, letter case included (a sheet renamed `diff gstr-3b` gives no values, so "Missing") | A:power_query_extractor/core/report_processor.py:546-601 (exact name :564-568) | B:power_query_extractor/core/report_processor.py:496-564 | ⚠ | ___ |
| 11 | Column look-up: each column takes the first extracted field whose name CONTAINS the column title, else the exact field - this is what puts J15/K15/L15 into the 9 columns of rows 7-9 | A:power_query_extractor/core/data_consolidator.py:137-150 | B:power_query_extractor/core/data_consolidator.py:125-138 | | ___ |
| 12 | Dormant: "calculated columns" (subtract, add, multiply or divide two cells; 'Error', 'Div/0'); no mapping uses one - say "drop" or "keep" | A:power_query_extractor/core/data_consolidator.py:130-136, 188-218 | B:power_query_extractor/core/data_consolidator.py:118-124, 176-206 | | ___ |

Evidence: RAN through both (section 5, Category 10 (10.2, 10.3)): a made-up ITC report with J15 = 101, AA1 = 501, O1 = 601 gave, in both, "GSTR-3B Difference" Diff IGST 101, "IMS Reco" IGST 101 (not 501) and "Extra In Purchase" IGST 101 (not 601); with J15 empty, the configured 501 and 601 appeared. The mismatch dates from the last shared commit `47a1492` (30-10-2025, which added rows 4-9). For rows 7-9 you might answer, for example, "except 7-9: keep what they show today", "except 7-9: use the configured cells" or "except 7-9: Not sure yet". A latent trap (one that does no harm today but would if the list were edited), the same in both: a second mapping of the same cell would silently replace the first (A-G78 = B-G78).
Recommendation: No recommendation - your call.
Your answer: ____

### P6-7. When a refresh counts as successful, and how figures are shown ⚠                Part 6 · refs A-R66, A-R73, A-R89, A-R90, B-R61, B-R69, B-R70, B-R71, B-R88, B-R89 (A-F21/B-F19, B-F20, A-F24/B-F23; gaps A-G3/B-G6, A-G15/B-G8, A-G17/B-G10, A-G18/B-G11, A-G79)
What this is about: How Step 4 decides that a client's report refreshed properly, and what the consolidated workbook shows when it did not - for example whether a ClientB row shows "Missing" or its figures.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | "The report" to refresh is the first file the folder listing returns for `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` in the Version folder (an older refreshed copy can come first on some disks) | A:power_query_extractor/core/report_processor.py:111-125 | B:power_query_extractor/core/report_processor.py:72-89 | | ___ |
| 2 | A refresh counts as successful when no Excel error box was seen AND a sheet named `Info` (any letter case) exists AND its cell BB2 is not empty, not blank, not 0, and not an Excel error text (#N/A, #REF! …). BB2 is never compared with the client: another client's name, `00` or a large number also pass. The code calls BB2 the client name; nothing in either project writes it | A:power_query_extractor/core/report_processor.py:638-685 (cell :656) | B:power_query_extractor/core/report_processor.py:589-636 (cell :607) | ⚠ | ___ |
| 3 | The refreshed copy is saved whatever the test says, and its figures are read anyway | A:power_query_extractor/core/report_processor.py:345-374, 187-188 | B:power_query_extractor/core/report_processor.py:300-366 | | ___ |
| 4 | If the test fails, every figure of that report shows "Missing" in the consolidated workbook although it was read; an empty cell also shows "Missing"; 0 stays 0 | A:power_query_extractor/core/data_consolidator.py:128-129, 152-154; A:power_query_extractor/config/cell_mappings.py:252 | B:power_query_extractor/core/data_consolidator.py:116-117, 141-142; B:power_query_extractor/config/cell_mappings.py:252 | | ___ |
| 5 | Non-zero numbers are shown with two decimals and thousands separators (`#,##0.00`); the stored value is not rounded; a number stored as text stays text; an Excel error in a figure cell may arrive as a large negative number (MEDIUM) | A:power_query_extractor/core/data_consolidator.py:160-162 | B:power_query_extractor/core/data_consolidator.py:148-150 | | ___ |

Evidence: RAN through both with a stand-in for Excel - a fake Excel that records what the program asks of it, since this machine has no Excel (section 5, Category 9 (9.7, 9.8) and Category 11 (11.4)): BB2 empty, blank, `0`, `0.0`, `#N/A` → fail; `ClientA`, `ClientB` (another client), `00`, `12.5`, `-2146826246` → pass, identically. What real Excel does could only be READ. Whether `Info!BB2` is still the cell your current templates fill after a refresh is for you to confirm ⚠. Shared gaps: the refresh cannot be confirmed (A-G3 = B-G6); a copy of a copy may be refreshed (A-G15 = B-G8); figures that exist shown as "Missing" (A-G17 = B-G10).
Recommendation: No recommendation - your call.
Your answer: ____

### P6-8. Waits, the refreshed-copy suffix, key presses and Excel limits ⚠                Part 6 · refs A-R67, A-R71, A-R72 (waits), A-R104, A-R105, A-R112, B-R62, B-R66, B-R67, B-R68, B-R100, B-R102, B-R103, B-R114 (A-F21/B-F19, A-F22/B-F21, A-F26/B-F26; gaps A-G41/B-G54, A-G44/B-G51, A-G50/B-G57, A-G55/B-G61, B-G62)
What this is about: Fixed values the two projects share for driving Excel: how long they wait, how a refreshed copy is named, which keys start a refresh, which words mark an Excel error box, and the path length treated as Excel's limit. How the new tool should drive Excel is a Part 8 decision (your Standard 12 sets most of it); here only confirm whether these values still match your office PCs and Excel.
A's way: the same as B's; the table gives where each rule lives.
B's way: the same as A's.

| # | Rule (both projects) | A | B | ⚠ | Still correct? |
|---|---|---|---|---|---|
| 1 | Refreshed copy = `<report name><suffix>.xlsx`, default suffix `_Refreshed_{timestamp}`, where `{timestamp}` becomes `DDMMYY_HHMM` (the suffix box itself, and which copies Skip refresh reuses: P2-13) | A:power_query_extractor/core/report_processor.py:51, 127-130, 164 | B:power_query_extractor/core/report_processor.py:28, 91-94, 118-120 | | ___ |
| 2 | The refresh is started by pressing Alt+A, R, A in a visible, maximised Excel - Data → Refresh All in an English Excel; these key tips (the letters Excel offers as keyboard short-cuts) depend on the Excel version and display language | A:power_query_extractor/core/report_processor.py:278-293 | B:power_query_extractor/core/report_processor.py:234-249 | ⚠ | ___ |
| 3 | Fixed waits per refreshed report: 19 s + the wait w + the larger of 10 s and w → 39 s at the default 10 s, before Excel's own start-up (the range of the wait box differs: a technical item) | A:power_query_extractor/core/report_processor.py:185, 253-363 | B:power_query_extractor/core/report_processor.py:139-141, 207-319 | | ___ |
| 4 | An Excel error box is recognised by English words in a window title ('microsoft excel', 'error', 'cannot', 'failed' …) on any window of the desktop and answered with its OK / Close / Yes / Continue button | A:power_query_extractor/core/report_processor.py:417-525 | B:power_query_extractor/core/report_processor.py:368-476 | ⚠ | ___ |
| 5 | 218 characters treated as Excel's path limit - used only to choose another way of opening a long path and to explain errors; nothing measures paths before folders and files are created | A:core/excel_handler.py:356; A:power_query_extractor/core/report_processor.py:720, 795 | B:core/excel_handler.py:293; B:power_query_extractor/core/report_processor.py:668, 743 | ⚠ | ___ |
| 6 | Step 3 pauses 0.5 s after copying the template and 0.5 s after closing it (1 s per report) | A:core/excel_handler.py:155, 210 | B:core/excel_handler.py:125, 180 | | ___ |

Evidence: RAN with a stand-in for Excel (section 5, Category 9 (9.5, 9.6, 9.9) and Category 12): the same call order, keys and waits in both (39 s per report at 10 s; 139 s at 60 s); what real Excel does is READ only. A's own commit `7347383` measured paths up to 270 characters on the share. Shared gaps: keys and clicks reach whatever program is in front (A-G55 = B-G61, B-G62); Excel set-ups the code does not expect, such as another display language (A-G44 = B-G51); fixed waits with no time-out (A-G50 = B-G57); long paths not measured (A-G41 = B-G54).
Recommendation: No recommendation - your call.
Your answer: ____

#### For the closing list

Same in both and working - no decision needed (object if you disagree):

- Source files are only read and copied - never moved, renamed, changed or deleted - and each copy's size is compared with the source (A-R47 · B-R42; A:utils/helpers.py:172-215 = B:utils/helpers.py:142-185; RAN, identical function). The shared gap of a partial copy left after an interruption (A-G11 = B-G26) is not part of this line.

#### Refs covered

Part 2 items (P2-1 to P2-14):

- Rules: A-R2, A-R12, A-R16, A-R17, A-R29 (display), A-R36, A-R39, A-R41 (folder names), A-R51, A-R52, A-R54 (folder text), A-R55 (folder text), A-R58 (openpyxl path), A-R59 (Step 4 recognition), A-R60 (live reads, order), A-R61 (Variations sheet), A-R65 (part), A-R68, A-R69, A-R87, A-R91, A-R134 (Dry Run gate), A-R138 (Missing column), A-R139, A-R147 (file-name card); B-R7, B-R12, B-R13, B-R17, B-R27 (display), B-R32 (via the length box), B-R35, B-R37 (folder names), B-R48, B-R50 (snapshot part), B-R51, B-R53 (openpyxl path), B-R54 (folder text), B-R55 (folder text), B-R60, B-R61 (Step 4 recognition, part), B-R63, B-R87, B-R90, B-R94, B-R95, B-R133, B-R134 (Missing column, order), B-R135, B-R136 (part), B-R138 (length box), B-R141. Mentioned for context only: A-R27, A-R35, A-R57, A-R96, B-R25, B-R56, B-R97.
- Features (GST part only; their technical parts belong to other parts): A-F6, A-F7, A-F8, A-F9, A-F11, A-F13, A-F14, A-F15, A-F20, A-F22, A-F24, A-F31; B-F4, B-F6, B-F8, B-F9, B-F11, B-F12, B-F13, B-F14, B-F18, B-F21, B-F23, B-F33.
- Gaps: A-G2, A-G6, A-G14, A-G15, A-G16, A-G21, A-G25, A-G26, A-G38, A-G41, A-G42, A-G47, A-G66, A-G67; B-G3, B-G7, B-G8, B-G9, B-G20, B-G24, B-G28, B-G30, B-G46 (in part: upper-case names), B-G47, B-G48, B-G54, B-G66, B-G72, B-G73 (the "Folder Name" column only).

Part 6 items (P6-1 to P6-8):

- Rules: A-R1, A-R3, A-R4, A-R5, A-R6, A-R7, A-R11, A-R12 (hint texts), A-R17 (scan key), A-R21, A-R22, A-R23, A-R24, A-R25, A-R26, A-R27, A-R28, A-R29 (status rule), A-R30, A-R32, A-R33, A-R35, A-R38, A-R41 (type → folder map), A-R46, A-R48, A-R54 (cells), A-R55 (cells), A-R56, A-R57, A-R58 (Excel path), A-R59 (report names), A-R60 (both reports for every client), A-R61 (name), A-R62, A-R66, A-R67, A-R71, A-R72 (waits), A-R73, A-R74, A-R75, A-R76, A-R77, A-R78, A-R79, A-R80, A-R81, A-R82, A-R83, A-R85, A-R86 (name), A-R88, A-R89, A-R90, A-R104, A-R105, A-R112; B-R5, B-R6, B-R8, B-R9, B-R10, B-R11, B-R12 (hint texts), B-R17 (scan key), B-R21, B-R22, B-R23, B-R24, B-R25, B-R26, B-R27 (status rule), B-R28, B-R30, B-R31, B-R32, B-R34, B-R37 (type → folder map), B-R40, B-R43 (backup name), B-R50, B-R51 (names), B-R53 (Excel path), B-R54 (cells), B-R55 (cells), B-R56, B-R57, B-R61, B-R62, B-R66, B-R67, B-R68, B-R69, B-R70, B-R71, B-R73, B-R74, B-R75, B-R76, B-R77, B-R78, B-R79, B-R80, B-R81, B-R82, B-R83, B-R84, B-R86 (name), B-R88, B-R89, B-R94 (name), B-R100, B-R102, B-R103, B-R114.
- Features (shared GST parts): A-F7, A-F8, A-F13, A-F14, A-F15, A-F16, A-F21, A-F22, A-F23, A-F24, A-F26; B-F6, B-F8, B-F12, B-F13, B-F14, B-F15, B-F19, B-F20, B-F21, B-F22, B-F23, B-F26.
- Gaps: A-G1, A-G3, A-G4, A-G5, A-G6, A-G8, A-G9, A-G15, A-G17, A-G18, A-G33, A-G34, A-G35, A-G36, A-G38, A-G39, A-G41, A-G44, A-G50, A-G55, A-G76, A-G77, A-G78, A-G79; B-G1, B-G2, B-G3, B-G4, B-G6, B-G8, B-G10, B-G11, B-G15, B-G16, B-G17, B-G23, B-G27, B-G47, B-G49, B-G51, B-G52, B-G54, B-G57, B-G61, B-G62, B-G78, B-G81.

Closing list: A-R47, B-R42 (gap A-G11 / B-G26 named, not covered).
