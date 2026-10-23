## Part 11 - About you and how the tool is used

This part asks what the code cannot show: what the tool is for in your own words, who runs it, on which PCs, with what data, what has gone wrong in real use, and what matters most to you. These are your answers, so **no item carries a recommendation**: every item says "No recommendation - your call". Each item keeps the shape of the rest of the list: "What this is about" ends with the lettered questions to answer; "A's way" and "B's way" say what each old program does that bears on them; "Evidence" gives what the audits found. Answer each lettered question with a word ("yes", "no", "A", "B") or a short sentence; "Not sure yet" is a valid answer and becomes an open item in Session 0 of the Decisions Log. Where a question touches a GST rule it is marked **[GST RULE]** and asks only what happens in your office - the rule itself is decided in Parts 2 and 6. **⚠** marks a value that changes over time (Windows, Excel and Python versions, template cell positions, file-name conventions, path limits). No client name appears here; `<Client A>`, "Alpha Test Traders" and "Beta Demo Pvt Ltd" are made up.

**How the references work.** A-F / B-F = a feature, A-R / B-R = a rule, A-G / B-G = a gap (problem) in the A or B audit. "A X6-Q46" = question 46 of Appendix X6 of A's audit (B's questions are numbered "B X6-Q37" the same way). "A 11.7 #20" = row 20 of Section 11.7 of A's audit. "Section 9, H1" = row H1 of section 9 of MERGE_COMPARISON.md (the comparison document). `A:path:line` / `B:path:line` = a place in A at `bf2f1dd` or in B at `b0e7d0e`. "RAN" means the old code was actually run on made-up files; "READ" means the fact comes from reading the code. HIGH / MEDIUM / LOW is the audits' confidence in an inference (a conclusion drawn from clues rather than seen directly).

**Words used in this part**

| Word | Meaning |
|---|---|
| Steps 1-4 | The four tabs of both programs: Step 1 Setup (choose the folders and templates), Step 2 Validation (the client list), Step 3 Processing (copy the files, make the reports), Step 4 Extract (refresh the reports and collect the figures). |
| Template | Your two master workbooks, ITC and Sales, which the tool copies for every client. |
| `.xlsx`, `.xltx`, `.xlsm`, `.xls`; macro | `.xlsx` is a normal workbook, `.xltx` an Excel template file, `.xlsm` a workbook that may hold macros, `.xls` the old pre-2007 format. A macro is a small program stored inside a workbook. |
| Links sheet | The sheet in each template where the tool writes the folder paths and file names of a client's copied files, for the template's queries to read. |
| Power Query, refresh | Excel's data-import tool inside your templates. A *refresh* re-runs its queries so the report reads the client's files. |
| Report, refreshed copy | The per-client copy of a template (`ITC_Report_…`, `Sales_Report_…`). Before refreshing, Step 4 copies it to `…_Refreshed_<date time>.xlsx` and refreshes that copy. The added text (`_Refreshed_<date time>` by default) is the *suffix*, set in Step 4's "Refreshed File Suffix" box. |
| Skip refresh | A Step 4 tick box: reuse an existing refreshed copy instead of refreshing again. |
| Info!BB2 | Cell BB2 of the sheet "Info" in a report. Both projects treat a value there as proof that a refresh worked. |
| Consolidated workbook | `PQ_Extraction_Report_<date time>.xlsx`, Step 4's single workbook that collects 22 figures from every client's reports. |
| Run folder (batch), Version folder, category folders | `Annual Statement-<date time>` (one per run, holding that batch of clients); inside it `<client>-<state code>`; inside that `Version-<date time>`; inside that three category folders (A: `GSTR3B`, `ITC`, `Sales`; B: `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`). |
| Fresh, Re-Run, Resume | The three processing modes on Step 1. |
| Dry Run | The Step 2 button that previews what Step 3 would do, without copying anything. |
| Excel through COM, pywin32 | COM is Windows' way for one program to drive another. pywin32 is the Python add-on that lets the tool drive Excel. |
| openpyxl | A Python add-on that writes Excel files without Excel. It drops Power Query, and shapes drawn on a sheet (text boxes, buttons), when it saves a workbook. |
| Key tips (Alt+A, R, A) | The letters English Excel shows on the ribbon when Alt is pressed (A = Data tab, R then A = Refresh All). Both projects refresh by "typing" these keys into Excel. |
| Windows Script Host | A built-in Windows component that both projects use to bring Excel to the front and type the keys. Some offices switch it off. |
| Protected View, add-in | Excel's read-only safety mode for files from the internet or the network; an add-in is an extra tool installed into Excel. Both can show their own pop-ups. |
| Console | The black text window a program can print to. |
| Launcher, desktop shortcut | Standard 13's small `Run <App Name>.bat` file that starts the app without a console; users double-click a desktop shortcut to it. |
| Shared apps folder, `_runtime` | Standard 13's network folder that holds every office app plus one shared portable Python (`_runtime`), so nothing is installed on users' PCs. |
| Tkinter, PyQt5, PySide6 | Window kits: the add-ons that draw windows, buttons and tables. A uses Tkinter, B uses PyQt5, your Standard 12 names PySide6. |
| Network share; drive letter or UNC path | A folder on the office server that PCs open over the network. It can be reached as a drive letter (for example `P:`) or by its full network name `\\server\share\…` (a "UNC path"). |
| NAS | A network storage box used instead of a Windows server. |
| OneDrive | Microsoft's folder that copies itself to the cloud and back. |
| VPN, Tailscale | A secure connection to the office network from outside; Tailscale is one such product. |
| Path length limit ⚠ | The longest folder-plus-file name Excel will open. Both projects treat 218 characters as Excel's limit. *Long-path support* is a Windows setting that lifts Windows' own 260-character limit. |
| Display scaling | Windows' setting that enlarges everything on screen (100 %, 125 %, 150 %). |
| Dark mode | A setting that shows the window in light text on a dark background. |
| Repository, commit, branch, public | A repository is a project's folder with its full change history on GitHub (the history is kept by git, the program that records every saved change); a commit is one saved change, named by a short code such as `4f84ba9`; a branch is a separate line of changes inside a repository; a public repository can be read by anyone. |
| Gap | A problem the audits found: something the program does wrong, or a real case it does not handle. |
| Variation | The code's word for a file in the source folder that it could not recognise or use. |
| Test, test harness | A small program that checks the tool automatically; a harness is a helper set-up used to run such checks. |
| Contract, Zero-Defect Requirement | A contract is a promise the new tool must keep on every run (Part 7). The Zero-Defect Requirement is the one thing that must never fail. |

---

### P11-1. What the tool does, in your own words                                  Part 11 · refs A-F7, A-F13, A-F14, A-F15, A-F21, A-F23, A-F24, B-F6, B-F12, B-F13, B-F14, B-F19, B-F22, B-F23
What this is about: Section 1 of your Rebuild Brief must be in your words, not the audit's. Please write 3-4 sentences on what the tool is for in your annual-statement work, as you would explain it to a new article assistant who joins in the season; the one-sentence version is welcome too. The two lines below are what the audits found, for comparison only - please do not copy them; if your description differs, that difference is useful in itself.
A's way: recognises each client's downloads by file name only, copies them into a dated folder per client, fills an ITC and a Sales report from your two templates, refreshes their Power Query, and collects 22 figures from every client into one consolidated workbook (A audit Section 1; A-F7, A-F13, A-F14, A-F15, A-F21, A-F23, A-F24).
B's way: the same job in the same four steps, in a PyQt5 window that calls itself "v4.0" (B audit Section 1; B-F6, B-F12, B-F13, B-F14, B-F19, B-F22, B-F23).
Evidence: Section 1 of each audit; the comparison's summary is MERGE_COMPARISON.md section 1.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-2. The problem you first built it for, and how it has changed                                  Part 11 · refs A-R109, A-R110, A-F7, B-F6
What this is about: Knowing the job the tool was first built for, and how it has grown, tells the build what to make fast and safe - for example, whether it began as a way to sort one season's GSTR-2B and GSTR-3B downloads and only later took on the Power Query refresh. Please answer: (a) What problem were you solving when you first built it? (b) Is it bigger, smaller or different now - more clients, more file types, more people?
A's way: A's July 2026 commit messages describe a scan of 2,508 files on the office share, made faster for that size (`4f84ba9`, 25-07-2026; A-R109, A-R110), and a run that produced 3,030 output files (`7347383`, 26-07-2026); the scan itself is A-F7.
B's way: B's history holds no record of a real run (B X6-Q63); its scan is B-F6.
Evidence: the tool existed before git: about 2½ months of unrecorded work before 31-08-2025, and the oldest surviving compiled copy of the refresh code (a copy Python makes from the code by itself) is dated 25-06-2025 (A audit Section 1; B audit 8.10 #1).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-3. Parts of the tool you never use                                  Part 11 · refs A-F5, A-F6, A-F10, A-F11, A-F12, A-F13, A-F16, A-F22, A-F28, A-F29, B-F4, B-F10, B-F11, B-F12, B-F15, B-F17, B-F21, B-F25, B-F28, B-F30, B-F31, A-R43, A-R67, A-R68, B-R44, B-R62, B-R63, B-R65, A X6-Q9, A X6-Q33, A X6-Q58, A X6-Q60, A X6-Q61, B X6-Q6, B X6-Q28, B X6-Q48, B X6-Q54, A 11.7 #27, B 11.7 #6, B 11.7 #28, B 11.7 #36
What this is about: A feature nobody uses costs build time and can hide bugs - for example a button no article assistant has pressed in two GSTR-9 seasons. Parts 3 and 4 then decide keep or drop with your answer in mind. Please answer: (a) Which of the features listed below do you and your staff never use? (b) Which refreshed-file suffix do you actually use? (c) Which "max client folder name length" do you use?
A's way: Dry Run (A-F11); EXPORT LIST, the Step 2 button that saves the ticked clients to a text file (works, A-F12); the per-client "📁 Name" toggle (A only, A-F10); the "organise only / full pipeline" choice (A only, A-F5); the Re-Run and Resume modes (A-F13, A-R43); the max folder-name length box (A-F6); the Step 3 summary workbook and the per-client organisation report, a text file listing each client's folders and files (A-F16); Skip refresh and the "Refreshed File Suffix" box (A-F22, A-R67, A-R68); the separate extractor window started with `launch_extractor.py` (A-F29); dark mode (A-F28); the Step 2 keyboard keys ↑ / ↓ and Space (MERGE_COMPARISON.md section 6.2, row 9).
B's way: Dry Run (B-F11); EXPORT LIST, which does nothing in B (B-F10); the Re-Run and Resume modes (B-F12, B-R44); the max length box (B-F4); the summary workbook and organisation report (B-F15); Skip refresh and the suffix box (B-F21, B-R62, B-R65); `launch_extractor.py`, which in B opens the main window on Step 4 (B-F30), and the old stand-alone extractor, which can no longer start at all (B-F31); dark mode (B-F28); the "Last Run" column and run history, B's record of past runs in your user folder (B-F17); the `Extraction_Log` text file, a copy of Step 4's on-screen log in the batch folder (B-F25).
Evidence: A's `2aaa2b1` (25-07-2026) traces a "Reported" wrong refresh state to a customised suffix, so a custom suffix was in use at least once; B's Skip refresh finds only copies named `_Refreshed_` (B-R63; B X6-Q6). Which length is right is decided in Part 2; here only what you use.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-4. Who uses the tool, and how comfortable they are with computers                                  Part 11 · refs A-F4, A-F31, B-F33, A X6-Q59
What this is about: The new tool is built for the least technical person who will run it - for example an article assistant who joins for the GSTR-9 season. Standard 13 assumes users "never run code" and start apps from a desktop shortcut (rebuild_pipeline_FINAL.md:902). Please answer: (a) Who besides you runs it (roles and how many)? (b) Can they type a folder path, read an error message, rename a wrongly named file? (c) Who reported the July 2026 problems - you or a colleague?
A's way: A's July 2026 work added a live checklist (`322717f`) and hover help (text that appears when the mouse rests on a control; `bf2f1dd`) on Step 1 (A-F4, A-F31).
B's way: Step 1 has fixed help cards and hover texts (B-F33); nothing in B's repository says who runs the tool (B audit Section 1: NOT SURE).
Evidence: five A commits fix problems their messages call "Reported" (`2acffed`, `2aaa2b1`, `725fbaf`, `169c7b5`, `bf2f1dd`; A audit Section 1); two of them were reported "with a screenshot" (`169c7b5`) or "with screenshots" (`725fbaf`).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-5. How often, and in which months                                  Part 11 · refs A-F13, B-F12
What this is about: Whether the tool runs once a year in a rush or all year round changes how much speed, Resume and batch choice matter. Example: one big batch for every client in the annual-return season, then re-runs for clients whose downloads came late. Please answer: (a) In which months do you use it, and how many batches per season? (b) Do you re-run single clients later (new downloads, corrected files), and how often? (c) Is it used outside the season at all?
A's way: names each batch folder only by date and time (`Annual Statement-<DDMMYY HHMM>`; `A:utils/helpers.py:25-27`), so nothing in the code records how often you run; a Re-Run adds a new Version folder inside the latest batch (A-F13).
B's way: the same naming (`B:utils/helpers.py:23-25`) and the same Re-Run (B-F12).
Evidence: A's own commits cluster in July 2026 (32 of its 34), B's in February 2026 (all 5) (MERGE_COMPARISON.md section 2).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-6. Which project is in daily use now, and since when                                  Part 11 · refs A-R39, B-R35, B-F2, B-R118, A-F27, B-F27, B X6-Q60, B X6-Q63 · [GST RULE] (category folder names, question d)
What this is about: The names mislead: A ("oldv1") is the one you worked on in July 2026; B ("oldv2", PyQt5 "v4.0") stopped on 09-02-2026. Which one your office actually runs decides which behaviour people are used to, and which folder layout already sits on the share - for example whether last season's client folders hold a `GSTR3B` or a `GSTR-3B Exports` folder. Please answer: (a) A, B or an older version - since when, and on which PCs? (b) Was B ever used for real client work, and did Steps 3 and 4 run to the end? (c) Were both run on the same PC? (d) Do batches with the old long folder names still need Step 4 work?
A's way: category folders `GSTR3B` / `ITC` / `Sales` since `7347383` (26-07-2026; A-R39); last change 27-07-2026, 288 tests, five "Reported" fixes from July 2026 (A audit Section 1); settings in `~/.gst_organizer_cache.json`, a file in your Windows user folder (A-F27).
B's way: category folders `GSTR-3B Exports` / `Other ITC related files` / `Sales related files`, as in all code before `7347383` (B-R35); last change 09-02-2026, no test and no record of a run on Windows (B X6-Q63). On its first start B copies A's settings file once into its own (B-F2, B-R118, B-F27), so running both on one PC leaves a trace.
Evidence: MERGE_COMPARISON.md section 1.3. Which category-folder names the new tool uses is a GST rule decided in Part 2, because the names are also written into the Links cells; here only what exists on your share.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-7. Your actual workflow, click by click                                  Part 11 · refs A-F3, A-F5, A-F7, A-F8, A-F11, A-F17, A-F18, B-F3, B-F6, B-F8, B-F11, B-F16, B-F18, B-F29
What this is about: Brief Section 2 asks what you click first, second and third. The audits traced the path the code expects; please say where real use differs - for example, whether you wait for a client's late GSTR-2B download between Step 3 and Step 4. Please answer: (a) Which steps do you skip or repeat? (b) Between Step 3 and Step 4, what do you do - check the reports, wait a day, rename files? (c) Do you sometimes open the tool only for Step 4, on a later day?
A's way: Step 1 (source folder, two templates, target folder) → Scan → Step 2 (clients start unticked) → optional Dry Run → Start → Step 3 shows progress and stays there unless "full pipeline" was chosen (A-F5) → Step 4: Scan the folder, then Start Processing → "Open the consolidated report?" (MERGE_COMPARISON.md section 6.2).
B's way: Step 1 → Scan → Step 2 (clients start ticked) → optional Dry Run → Start → Step 3 → B moves to Step 4 by itself with the list already filled → Start Extraction; no end question (MERGE_COMPARISON.md section 6.3).
Evidence: neither program starts the Power Query refresh without a click (MERGE_COMPARISON.md section 6.1). On a later day A needs a new Scan before Step 2, because the client list is not remembered; B fills Step 4 at start-up from the last saved target folder (MERGE_COMPARISON.md sections 6.2 and 6.3, "Variants").
Recommendation: No recommendation - your call.
Your answer: ____

### P11-8. How you use Step 4 in practice                                  Part 11 · refs A-F19, A-F20, B-F18, A-R42, B-R38, B-R39, B-R136, A-G45, B-G12, A X6-Q48, A X6-Q57, B X6-Q7, B X6-Q32, A 11.7 #26, B 11.7 #19
What this is about: Step 4 picks the batch by itself. Example: you ran a batch on 15-07, then re-ran three late clients on 22-07. Which batch should Step 4 show, and do you ever go back to an older one? Please answer: (a) Do you ever need Step 4 on an older batch, and how do you do it today? (b) Do Re-Runs leave several Version folders per client that you then refresh? (c) Do you open files inside older batch folders before running Step 4?
A's way: takes the `Annual Statement-…` folder with the newest "date modified" (the time the computer last recorded a change in it) and, in each client, only the newest Version folder (`A:power_query_extractor/gui/extractor_window.py:622-648`; A-R42, A-F19, A-F20).
B's way: takes the folder with the highest name, which starts with the day, so a 31-12-2025 batch beats a 21-10-2026 one (`B:gui/pyqt_main_window.py:430-448`; B-R39, B-G12), and lists every Version folder of each client (B-R136, B-F18). B's Re-Run and Resume, by contrast, add to the batch with the newest "date modified" (B-R38).
Evidence: in A, opening and saving a file inside an older batch can change its "date modified" and so make it the batch Step 4 shows, and pointing Step 4 at a batch folder itself finds nothing (A-G45). Which batch Step 4 should use is decided in Part 1; here only how you work.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-9. Which Windows, and which screens                                  Part 11 · refs B-G68, A-F2, B-F1, A X6-Q67, B 11.7 #33, A 11.7 #29
What this is about: Your Standard 12 pins PySide6 to version 6.12.x while any PC runs Windows 10 ⚠, because Qt 6.12 is the last Qt for Windows 10 (rebuild_pipeline_FINAL.md:833). Screen size and display scaling decide whether the window fits - for example on the laptop an article assistant takes to a client's office. Please answer: (a) Windows 10 or 11 on each PC and laptop that will run it ⚠? (b) Screen sizes and display scaling (100 / 125 / 150 %)? (c) Is it ever run on a Mac or Linux?
A's way: opens at 1200×800 (MERGE_COMPARISON.md section 6.2, row 2; A-F2). Off Windows, A offers at every start to install pywin32, which can never succeed there (A X6-Q67).
B's way: opens at 1400×900 and cannot be made smaller than 1200×800 (`B:gui/pyqt_main_window.py:72-73`; B-F1), taller than a 1366×768 laptop screen (B-G68).
Evidence: neither repository records the Windows version. In both, the Scan button sits below the visible part of Step 1 at the opening size (MERGE_COMPARISON.md section 6.1).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-10. Which Excel, in which language, set up how                                  Part 11 · refs A-F21, A-F26, B-F19, B-F26, A-R71, A-R105, B-R67, B-R102, A-G44, A-G79, B-G51, A X6-Q43, A X6-Q44, B X6-Q35, A 11.7 #19, B 11.7 #27
What this is about: The refresh types Alt+A, R, A into Excel and looks for pop-ups by English words. Another Excel version or display language may not respond, and the refresh then silently does not happen ⚠ - for example on a PC whose Excel shows its menus in another language. Please answer: (a) Excel version on each PC (Microsoft 365, 2021, 2019 …) ⚠? (b) Display language - is the Data ribbon in English everywhere? (c) Any Excel add-ins, or Protected View pop-ups, on opening the reports? (d) Has IT blocked Windows Script Host on any PC? (e) Does every PC that will run the tool have Excel?
A's way: brings Excel to the front and types the keys through Windows Script Host (`A:power_query_extractor/core/report_processor.py:282-292`; A-R71, A-F21); the English pop-up words are at `A:…:425-436` (A-R105, A-F26).
B's way: the same code (`B:power_query_extractor/core/report_processor.py:238-248`, pop-up words `B:…:376-387`; B-R67, B-R102, B-F19, B-F26).
Evidence: MERGE_COMPARISON.md section 9, H17 (A-G44, B-G51; READ, as no Excel can run here, A-G79). Every PC that runs Step 3 or Step 4 needs Excel; in A, without it every Step 3 report fails with a message that blames the Links sheet (A X6-Q44).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-11. Which Python, and how people start the tool today                                  Part 11 · refs A-F1, A-F30, B-F1, B-F32, A-R120, A-R124, A-R127, B-R121, B-R124, B-R125, A-G22, A-G23, B-G34, A X6-Q62, A X6-Q65, A X6-Q76, B X6-Q50, B X6-Q51, A 11.7 #38, A 11.7 #2, A 11.7 #3, B 11.7 #2, B 11.7 #3
What this is about: Standard 13 says users double-click a desktop shortcut, see no console, and nothing is installed on their PCs. How the tool starts today shows how big that change is for your staff - for example, whether an article assistant now types `python main.py` every morning of the season. Please answer: (a) How do you and your staff start it - console, double-click on `main.py`, a shortcut, `launch_extractor.py`? (b) From a local folder or a shared network folder? (c) Who installed Python and pywin32 on each PC, and which Python version ⚠? (d) Which Python do you run A's tests with - `python` or `py`?
A's way: `python main.py` in a console; without a console A stops before its window opens (`A:main.py:17`; A-R127, A-G22); it writes its log into whatever folder it was started from (`A:main.py:20-27`; A-R120); a start folder it cannot write to, or a read-only program folder on the first start, stops it at start-up (A-G23); if an add-on is missing it offers to install it (A-F1, A-R124).
B's way: `python main.py` in a console; without a console B shows nothing when start-up fails and leaves only a line in its log file (`B:main.py:49-59`; B-R125, B-G34); it writes the same log into the start folder (`B:main.py:15-23`; B-R121); it declares no exact versions and leaves pywin32 out of its package list (B-R124).
Evidence: neither has a `.bat` launcher or a shortcut (MERGE_COMPARISON.md section 7.1). On your PC, `py` starts a "free-threaded" Python 3.14t (a special build of Python) without pywin32, while `python` starts the one with pywin32 (`c463256`, 26-07-2026) ⚠; which version `python` is, is not recorded (MEDIUM that it is 3.14; MERGE_COMPARISON.md section 7.1).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-12. Office PCs only, or also from outside over a VPN                                  Part 11 · refs A-R109, A-R110, A X6-Q68, A 11.7 #30, B 11.7 #32
What this is about: Standard 13 measured that loading Python and the window kit from the shared folder over a VPN took 50-100 seconds instead of 1-2, and asks to be consulted before an app is given to someone who works outside the office (rebuild_pipeline_FINAL.md:927-930) - for example a partner who runs a batch from home in the last week of the annual-return season. Please answer: (a) Who runs it from outside the office, and how often? (b) Only scans, or also Step 3 and the Step 4 refresh (Excel) over the VPN?
A's way: A's speed work was measured on the live office share over Tailscale: 2,508 files scanned in 41 seconds instead of a projected 11 minutes once the first bytes of the files were read in parallel (`4f84ba9`, 25-07-2026; A-R110), on top of reading the folder listing only once (A-R109; `a603c02`, `56aaad9`).
B's way: none of this speed work; B's history records no run at all (B X6-Q63).
Evidence: a measurement over Tailscale suggests that at least one scan was run from outside the office (an inference from the commit message; A X6-Q68). Nothing shows whether Step 3 or Step 4 was ever run that way.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-13. Your shared apps folder                                  Part 11 · refs Standard 13 (rebuild_pipeline_FINAL.md:908-949), Tech Stack (rebuild_pipeline_FINAL.md:213-216)
What this is about: The Tech Stack in CLAUDE.md needs the path of the office shared apps folder (example name `\\server\share\Apps`), where `_runtime` and every released app live. The new tool would sit there in its own folder with its `Run <App Name>.bat`, like any other office app built to Standard 13 - for example a GSTR-2B reconciler. Please answer: (a) Does the shared apps folder with `_runtime` already exist (other apps built under Standard 13)? (b) Its path - or write "in chat" if you regard it as confidential; it will be written into CLAUDE.md in NEW, so your Part 10 answer on making NEW private matters. (c) Which Python version is in `_runtime` today ⚠?
A's way: no shared apps folder and no shared runtime; it runs from wherever its folder is copied, with the Python installed on that PC (MERGE_COMPARISON.md section 7.1).
B's way: the same (MERGE_COMPARISON.md section 7.1).
Evidence: the office share path appears in two of A's commit messages (`4f84ba9`, `7347383`; MERGE_COMPARISON.md section 10, D5); it is not repeated here. Standard 13 says all its testing used Python 3.14 ⚠ (rebuild_pipeline_FINAL.md:931).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-14. Where client files sit while the tool works                                  Part 11 · refs A-R39, A-R56, A-R100, A-R104, B-R35, B-R56, B-R100, A-G15, A-G41, A-G46, B-G53, B-G54, B-G55, A X6-Q29, A X6-Q39, A X6-Q47, B X6-Q39, B 11.7 #31, A 11.7 #25 · [GST RULE] (the path form written into the Links cells, question b) ⚠
What this is about: The source folder (the downloads) and the target folder (organised copies and reports) may be on a local disk, a network share or OneDrive. Each behaves differently for speed, long paths and files locked by someone else ⚠ - for example a colleague who has a client's GSTR-2B reco open while the batch runs. Please answer: (a) Source and target: local disk, Windows server share, NAS or OneDrive? (b) Drive letter (e.g. `P:`) or `\\server\share`, and the same on every PC? (c) Did the move to the shorter share root happen? (d) Is Windows long-path support switched on, if you know ⚠?
A's way: writes each path exactly as chosen - drive letter or `\\server\share` - into the reports' Links cells (A-R56); a workbook already open under the other form of the same path is not recognised (A-R100); since `7347383` it uses short category folder names to fit Excel's path limit (A-R39); it uses 218 characters ⚠ only to choose how to open a file and to explain errors, and measures no path before folders and files are created (A-R104, A-G41).
B's way: the same Links values (B-R56) with the long category folder names (B-R35) and the same 218-character handling (B-R100); files already open elsewhere, and a share that drops out, are not handled (B-G53, B-G54, B-G55).
Evidence: a drive letter on one PC and a `\\server\share` path on another can differ (A-G46); which form goes into the reports is a GST rule decided in Part 2 or 6 (A audit 11.7 #25). On a NAS the order in which the folder lists files can change which report Step 4 treats as the original (A-G15, A X6-Q47). A's `7347383` measured a real run: worst path 270 characters, 1,094 of 3,030 files over Excel's ~218 limit ⚠, "which is why long client names have been failing"; its fix assumed a planned move to a shorter share root (A X6-Q39).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-15. How big a typical run is                                  Part 11 · refs A-R72, B-R68, A-G48, A-G50, B-G56, B-G57, B-G58, A X6-Q69, B X6-Q22, A 11.7 #39
What this is about: Size decides how fast the tool must be, how it shows progress and how it survives interruptions. Example: 120 clients × 8 downloads = about 1,000 files in one source folder. Please answer: (a) Clients per batch, and files per client? (b) How long does a full batch take today, Step 3 and Step 4 separately? (c) One batch for all clients, or several smaller ones?
A's way: checks the first bytes of each file in parallel (A-R110), but reads the source folder on the window's own thread (its single line of work), so the window freezes on a slow share (A-G48); Step 4 waits fixed times per report (A-R72, A-G50).
B's way: scans in the background (B-R111), but its scan log writes one line per file and slows sharply as it grows (B-G56, B X6-Q22); it walks the target folder on the window's own thread (B-G58); the same fixed Step 4 waits (B-R68, B-G57).
Evidence: A's July runs had 2,508 source files and 3,030 output files (`4f84ba9`, `7347383`). Step 4 waits about 36-39 seconds per report at the default 10-second wait - about 2 hours for 100 clients (MERGE_COMPARISON.md section 9, H18).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-16. Two people, or two jobs, at the same time                                  Part 11 · refs A-F17, A-F21, B-F34, B-R111, A-G71, B-G82, A X6-Q45, A X6-Q50, B X6-Q38, B X6-Q39, A 11.7 #28, B 11.7 #30 · [GST RULE] (which template a report is built from, question d)
What this is about: In the season two staff may work on one batch, or one person may run Step 3 for new clients while Step 4 refreshes earlier ones. Both steps drive Excel. Please answer: (a) Do two people ever process the same target folder or batch at the same time? (b) Has Step 3 ever run while Step 4 was refreshing? (c) Is the program ever open twice on one PC? (d) While Step 3 runs, does anyone change a template or a folder on Step 1, or press Scan again?
A's way: nothing stops two runs from working on the same Excel, folders or files at once (A-G71; A-F17, A-F21); A re-reads the Step 1 template boxes for every client, so a template changed mid-run applies to the later clients (MERGE_COMPARISON.md section 3.1, row 15; A X6-Q45, A X6-Q50).
B's way: refuses only a second job of the same kind, silently (B-F34, B-R111); two copies of the program, or two users, overwrite each other's settings, history and output (B-G82); B uses the template paths fixed when Start was pressed (MERGE_COMPARISON.md section 3.1, row 15).
Evidence: RAN (MERGE_COMPARISON.md section 5, Category 8, 8.6): with the ITC template changed after the first client, A built one batch from two templates without a warning, and B built all reports from the first. Which template a client's report is built from is a GST rule decided in Part 2; here only what happens in your office.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-17. What you still do by hand after the tool finishes                                  Part 11 · refs A-F16, A-F24, B-F15, B-F23
What this is about: Work done by hand after every run is the first place a new tool can save time (Brief Section 4). Please answer: what do you or your staff do after the consolidated workbook appears? For example: open each refreshed report to check it, copy figures into the GSTR-9 / 9C working papers, chase clients for missing files, rename files and re-run. Name each step and roughly how long it takes.
A's way: ends with the consolidated workbook - a Summary sheet and one sheet per topic, one row per client, 22 figures in all - and asks "Open the consolidated report?" (A-F24); Step 3 has already left a processing summary workbook in the batch folder (A-F16).
B's way: the same consolidated workbook, with one row per Version folder and no end question (B-F23); the same summary workbook (B-F15).
Evidence: neither program does anything with the consolidated workbook after writing it (A only offers to open it), so the code cannot show what happens next (READ; MERGE_COMPARISON.md sections 6.2 and 6.3).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-18. Anything you wish it also did                                  Part 11 · refs A-R94, A-R95, B-F24, B-G71, A-G39, A-G45, B-G12, B-G49, A X6-Q7, B X6-Q9 · [GST RULE] (period check, in the evidence) ⚠
What this is about: Ideas from you, your staff or clients ("I wish this tool also did X") go into Brief Section 4, or into the Parking Lot if not for now - for example "print a one-page list of the clients whose files are still missing". Please answer: list your wishes, and say whether any client or colleague has asked for something the tool does not do.
A's way: shows a failure reason per report on screen, but the reason is in no sheet of the consolidated workbook; A's never-used code for an "Extraction Details" sheet suggests this was once intended (`A:power_query_extractor/core/data_consolidator.py:250-414`; A-R94, A-R95).
B's way: Step 4 shows only "ITC: FAIL", with the reason in the log file (B-G71; B-F24).
Evidence (prompts only, not suggestions): neither lets you pick which batch Step 4 works on (A-G45, B-G12); neither checks the period or year in file names [GST RULE] ⚠ (A-G39, B-G49).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-19. Gaps that hurt you: the Step 4 figures                                  Part 11 · refs A-F24, B-F23, A-R69, A-R80, A-R81, A-R82, A-R88, A-R89, B-R63, B-R79, B-R80, B-R81, B-R83, B-R88, A-G1, A-G2, A-G17, B-G1, B-G7, B-G10, B-G52, A X6-Q1, B X6-Q1, Section 9 H1, H2, H9, A 11.7 #9, B 11.7 #8 · [GST RULE] ⚠
What this is about: Brief Section 6 asks which gaps actually caused you problems. These three sit in code both projects share, and none of them shows a message - for example a client's IMS Reco columns quietly repeating its GSTR-3B difference. Which cell should feed each column is asked in Parts 2 and 6; here only what you have seen. Please answer: (a) Have the IMS Reco / Extra In Purchase / Extra In GSTR2B columns ever shown the GSTR-3B difference figures - was it noticed, and was any delivered work based on them? (b) Have you seen a client's row full of "Missing" although its report looked fine? (c) Has Skip refresh ever given you old figures?
A's way: fills each column with the first extracted figure whose name contains the column title (`A:power_query_extractor/core/data_consolidator.py:141-150`; A-R88); a failed refresh check turns every figure of that report into "Missing" (A-R89); Skip refresh reuses any earlier refreshed copy as a success (A-R69).
B's way: the same code (`B:power_query_extractor/core/data_consolidator.py:129-138`; B-R83, B-R88, B-R63).
Evidence (RAN in both, with made-up figures; MERGE_COMPARISON.md section 9): (1) in the consolidated workbook, the IGST / CGST / SGST columns of the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets show the `Diff GSTR-3B` J15 / K15 / L15 figures ⚠ whenever those were read, not their own configured cells (H1). (2) A report whose refresh check failed shows "Missing" in every column although its figures were read (H9). (3) With Skip refresh, a copy left by a failed refresh is reused next time as a success, with ✓ (H2).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-20. Gaps that hurt you: the refresh and your own Excel                                  Part 11 · refs A-R71, A-R73, A-R97, A-R103, A-R105, B-R66, B-R67, B-R70, B-R96, B-R102, B-R103, A-G3, A-G7, A-G44, A-G50, A-G54, A-G55, B-G6, B-G32, B-G39, B-G51, B-G57, B-G61, B-G62, B-G63, A X6-Q41, A X6-Q42, B X6-Q34, B X6-Q36, A 11.7 #24, Section 9 H3, H17, H18, H19, H20 ⚠, A 11.7 #18, B 11.7 #21 · [GST RULE] (the `Info!BB2` test, in the evidence) ⚠
What this is about: The refresh happens in a visible Excel that the program drives with key presses, and Step 3 borrows any Excel you already have open - for example the GSTR-2B reco you were checking when you pressed Start. Please answer: (a) Have your Excel windows vanished, or unsaved work been lost, during a run? (b) Letters typed into the wrong window, or another program's pop-up answered by itself? (c) A run that hung until Excel was closed by hand? (d) A refresh reported fine but with old figures? (e) A report whose Power Query - or a text box or button drawn on a sheet - had disappeared? (f) Do staff use the PC, or keep other workbooks open, during Step 3 or Step 4, and which Excel pop-ups appear during a refresh?
A's way: Step 3 attaches to your open Excel, hides it and quits it (`A:core/excel_handler.py:123-133, 238-246`; A-R97); Step 4 types the keys into whatever window is in front and clicks OK / Yes on any pop-up whose title looks like an error (A-R71, A-R105); a damaged pywin32 can let Step 3 save reports without their Power Query, and without any shapes drawn on a sheet, and still say "Success" (A-G7).
B's way: the same Step 3 and Step 4 code (`B:core/excel_handler.py:110-120, 203-211`; B-R96, B-R66, B-R67, B-R102, B-R103); without a working pywin32 B's window does not open at all (B-G32).
Evidence (MERGE_COMPARISON.md section 9): your Excel hidden and quit (H19); a made-up "Delete local copy?" box of another program was answered Yes (RAN in both, H20); nothing times out (H18); apart from the pop-up check, a refresh is judged only by `Info!BB2` ⚠, so a value already there passes (H3; A-R73, B-R70) - which test proves a refresh is a GST rule decided in Part 6.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-21. Gaps that hurt you: files, folders and counts                                  Part 11 · refs A-R35, A-R43, A-R45, A-R57, A-R60, B-R32, B-R43, B-R44, B-R56, B-R60, A-G4, A-G5, A-G6, A-G8, A-G9, A-G10, A-G19, A-G41, B-G2, B-G3, B-G4, B-G5, B-G13, B-G14, B-G15, B-G19, B-G20, B-G27, B-G54, Section 9 H4, H5, H6, H7, H8, H10, M1, M2, A 11.7 #15, B 11.7 #15 · [GST RULE] (H5, H6 and M1, in the evidence) ⚠
What this is about: These gaps can lose or mix client files while the counts say "Success" - for example one month's GSTR-3B export silently missing from a client's folder. Please answer: (a) Has any of these happened to you - which, and roughly when? (b) Have long paths stopped reports from opening (`7347383` says long client names "have been failing")? (c) Have you ever relied on Resume after an interruption?
A's way: the client folder is the cleaned name plus the state code, cut to "Max length" (A-R35); a file already there is backed up, overwritten or skipped according to the mode (A-R45); Resume makes a new Version folder (A-R43, A-G10); a client counts as "Success" unless an error escapes (A-R60); only the first file of a type is linked (A-R57).
B's way: the same rules (B-R32, B-R43, B-R44, B-R60, B-R56); in addition B leaves an unfilled template under the report's name when a report fails (B-G5), and never shows its list of problem files (B-G20).
Evidence (RAN in both on made-up clients; MERGE_COMPARISON.md section 9): "Beta Demo Private Limited" and "Beta Demo Pvt Ltd" became two clients but one folder, and one file replaced the other (H6) [GST RULE]; two downloads whose cleaned names coincide leave only one of them under the real name (Fresh keeps the other as a backup), both reported "Success" (H7); Resume copies the whole batch again into a new Version folder (H8); a client counts as "Success" although both its reports failed (H10); a template without a Links sheet gives no error (H4); Links cells for files a client lacks keep what the template holds (H5) [GST RULE] ⚠; only the first file of a type is linked (M1) [GST RULE]; a repeat run in the same minute overwrites earlier output (M2). Long paths: A-G41, B-G54.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-22. Gaps that hurt you: starting, closing and the screens                                  Part 11 · refs A-R127, B-R124, B-R125, A-G22, A-G27, A-G48, B-G32, B-G34, B-G35, B-G40, B-G67, B-G68, Section 9 H11, H12
What this is about: Problems you may meet before any figures appear - for example an article assistant double-clicks the program and nothing happens. Please answer: which of these have you or your staff met, and which annoyed you most?
A's way: started without a console, A dies silently (`A:main.py:17`; A-R127, A-G22); closing the window during a run kills the work in progress, and A asks "Stop and exit?" only during Step 3 (A-G27); A's window freezes while it scans a slow share (A-G48).
B's way: started without a console, B shows nothing when start-up fails (B-R125, B-G34); without pywin32 its window cannot open at all, and its message advises the wrong install command (`B:main.py:49-54`; B-R124, B-G32); closing never asks (B-G40); an error inside B's window ends the whole program (B-G35); three of B's four Browse buttons are white on white (`B:gui/widgets/colored_section.py:62`; B-G67); its smallest window is taller than a 1366×768 laptop screen (B-G68).
Evidence: MERGE_COMPARISON.md section 9, H11 and H12; section 6.4.1.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-23. Protections that have saved you                                  Part 11 · refs A-F26, B-F26, A-R9, A-R10, A-R46, A-R47, A-R98, A-R101, A-R102, A-R104, A-R105, A-R110, B-R1, B-R2, B-R3, B-R42, B-R43, B-R98, B-R99, B-R100, B-R102, B-R103
What this is about: Brief Section 6 also asks which handled cases saved you in real use, so the new tool keeps them on purpose - for example a re-run that would have overwritten a client's GSTR-2B reco but backed it up first. Please answer: which of these has visibly saved you, and from what?
A's way: client files are only copied, never moved or changed (A-R47); a file about to be overwritten in Fresh mode is first backed up as `…_backup_<date time>` (A-R46); tiny or damaged files and Excel's own `~$` lock files (the temporary files Excel makes while a workbook is open) are skipped (A-R9, A-R10); a path over 218 characters ⚠ is tried again in Windows' old shortened form, or with a prefix that lifts Windows' length limit (A-R101, A-R102, A-R104); Excel error pop-ups during a refresh are found and closed and their text kept (A-F26, A-R105). A only: reports are built on the local disk and copied once (A-R98, `12d505f`), and the scan reads files in parallel (A-R110, `4f84ba9`).
B's way: the same protections (B-R42; B-R43; B-R1, B-R2, B-R3; B-R98, B-R99, B-R100; B-F26, B-R102, B-R103), without A's two additions.
Evidence: MERGE_COMPARISON.md section 4 (fixes only one side got) and section 5, Categories 7 and 12.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-24. How client and state names are really written in your file names                                  Part 11 · refs A-R1, A-R2, A-R3, A-R4, A-R5, A-R6, A-R7, A-R17, A-R19, A-R21, A-R22, A-R35, B-R5, B-R6, B-R7, B-R8, B-R9, B-R10, B-R11, B-R17, B-R19, B-R21, B-R22, B-R32, A-G8, A-G25, A-G33, A-G35, A-G36, A-G40, A-G77, B-G15, B-G16, B-G17, B-G18, B-G80, A X6-Q20, A X6-Q21, A X6-Q22, A X6-Q23, A X6-Q24, B X6-Q18, B X6-Q19, B X6-Q29, Section 9 H6, H13, H15, H16, A 11.7 #11, B 11.7 #9, B 11.7 #10 · [GST RULE] ⚠
What this is about: Both projects recognise a file only by its name, for example `GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx`. What your real downloads look like decides which mistakes matter. Asking only: the naming and state-code rules are decided in Parts 2 and 6. Please answer: (a) Do client names contain hyphens or dots ("<X.Y.> <Client A>", "Pvt. Ltd.")? (b) How is the state written - full name, "MH", "Tamilnadu", the two-digit code "27"? (c) Any client names that are only digits? (d) Is one client ever spelt two ways? (e) Do two clients share their first 30-odd characters?
A's way: the six file-name patterns accept "anything but a hyphen" for client and state (`A:utils/constants.py:15-55`; A-R1 to A-R7); state text becomes a letter code from a table, with an invented fallback (`A:utils/helpers.py:592-618`, table `A:utils/constants.py:251-292`; A-R21, A-R22); a client written only in digits, like "0042", stops Dry Run and Start (A-G25).
B's way: the same patterns (`B:utils/constants.py:15-55`, identical except the IMS line; B-R5 to B-R11) and the same state codes (`B:utils/helpers.py:485-511`, table `B:utils/constants.py:234-275`; B-R21, B-R22).
Evidence (RAN in both; MERGE_COMPARISON.md section 9): a hyphen inside a client name shifts the parts (`GSTR-2B-Reco-Alpha-Test-Traders-Maharashtra-Q1.xlsx` became client "Alpha", state "Test") or the file is not recognised (H13); "Tamil Nadu" gives TN but "Tamilnadu" gives TAM, and "Unknown Place" gives UP, Uttar Pradesh's letters (H16); "Private Limited" and "Pvt Ltd" spellings become two clients (H15); two clients whose names agree in their first 32 characters share one folder (H6). These are letter codes in folder and file names; neither project has GSTIN or PAN logic (H16).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-25. File types, duplicates and periods in a real source folder                                  Part 11 · refs A-R2, A-R4, A-R5, A-R8, A-R25, A-R27, A-R57, B-R1, B-R7, B-R8, B-R10, B-R11, B-R14, B-R23, B-R25, B-R56, A-G6, A-G34, A-G37, A-G38, A-G39, B-G3, B-G23, B-G46, B-G47, B-G48, B-G49, A X6-Q11, A X6-Q16, A X6-Q17, A X6-Q18, A X6-Q19, B X6-Q13, B X6-Q16, B X6-Q17, B X6-Q20, Section 9 H14, M1, A 11.7 #10, A 11.7 #12, A 11.7 #17, B 11.7 #11, B 11.7 #12 · [GST RULE] ⚠
What this is about: The same for the rest of the name and the folder. Example: a browser saves a second download of an IMS file as `ImsReco-<Client A>-<State>-31032025 (1).xlsx`. Asking only; which file should feed a report is decided in Parts 2 and 6. Please answer: (a) Do "(1)" downloads occur, and for which types? (b) Does a client ever have two files of one type (two half-year 2B recos, two Sales files)? (c) Do files of different periods or financial years ever sit in one source folder? (d) Do spellings like `Sales-Reco-`, `GSTR-3B-`, `Annual Report-` occur? (e) `.xlsm` downloads, or files in sub-folders? (f) Is the 8-digit date in IMS names the download date or a period date?
A's way: accepts such an IMS "(1)" copy as an IMS Reco file and, because it sorts first, links it before the plain file (`A:utils/constants.py:23`; A-R2, A-R57).
B's way: rejects it as an unrecognised file (`B:utils/constants.py:23`; B-R7).
Evidence (MERGE_COMPARISON.md section 5, Category 1; section 9): both link only the first file of a type (M1; A-G6, B-G3), never check the period or year text (A-R25, B-R23; A-G39, B-G49), read `Sales-Reco-…` as a Sales file of a client called "Reco" (H14), ignore sub-folders, and do not recognise `.xlsm` downloads (A-R8, B-R1; A-G37, B-G46).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-26. The worst data you have received                                  Part 11 · refs A-F9, B-F9, A-G37, B-G20, B X6-Q41
What this is about: Brief Section 6 asks for the messiest real case, so the new tool is built and tested against it - for example a season when a client sent its GSTR-2B as a PDF. Please describe it without client names (use `<Client A>`). Please answer: (a) The worst batch you remember, and what went wrong? (b) Any messy situation the audits do not mention - for example a client's download with renamed sheets, a password-protected file, files sent as a .zip, CSV instead of Excel? (c) Do you look at the "Variations" sheet of the Step 3 summary workbook to find mis-named files?
A's way: counts unrecognised names as variations on screen and lists them in the "Variations" sheet of the Step 3 summary workbook (A-F9); tiny or damaged files are dropped without a word (A-G37).
B's way: counts them too, but never shows the list, and its "Variations" sheet is always empty (B-F9, B-G20; B X6-Q41).
Evidence: neither repository holds a sample file (MERGE_ANALYSIS_PROMPT.md section 3, fact 7), so the audits know only made-up cases.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-27. Your templates: how their queries find the client files                                  Part 11 · refs A-R39, A-R54, A-R55, A-R56, B-R35, B-R54, B-R55, B-R56, A-G5, A-G9, B-G2, B-G27, B-G29, B-G52, A X6-Q26, A X6-Q27, A X6-Q28, B X6-Q24, B X6-Q25, Section 9 H5, M2, A 11.7 #8, B 11.7 #14 · [GST RULE] ⚠
What this is about: The templates are in neither repository, so only you can say how their Power Query uses the Links cells. This decides, for example, whether A's short folder names (`GSTR3B` / `ITC` / `Sales`) work with your templates, and whether backups sitting in a folder get read. Please answer: (a) Do the queries read only the named file, or every file in the linked folder? (b) Does any query name a folder itself, such as "Other ITC related files" or "GSTR-3B Exports"? (To check: open the template, Data > Get Data > Launch Power Query Editor, open the Advanced Editor of each query and search for those names.) (c) What do the Links cells hold in the blank template - nothing, sample paths, or the last client's? (d) Which extension do the queries add (and what happens with `.xls` inputs)?
A's way: writes a folder path and, for most file types, a file name without its extension into the Links cells (`A:utils/constants.py:102-126`; A-R54, A-R55, A-R56) ⚠; the category folders are named `GSTR3B` / `ITC` / `Sales` (A-R39).
B's way: the same cells (`B:utils/constants.py:96-120`; B-R54, B-R55, B-R56) ⚠, with the folders named `GSTR-3B Exports` / `Other ITC related files` / `Sales related files` (B-R35).
Evidence (MERGE_COMPARISON.md section 9): cells for a file type the client lacks keep what the template holds (H5; A-G5, B-G2); ITC B3, B6, B9 and Sales B4, B7 are never written ⚠ (B-G52; B X6-Q25); backups and "(1)" copies can land in the linked folders (M2; A-G9, B-G27, B-G29).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-28. Your templates: format, refresh settings and the success cell                                  Part 11 · refs B-F20, A-R52, A-R73, A-R74, A-R75, A-R76, A-R77, A-R78, A-R79, A-R80, A-R81, A-R82, B-R48, B-R71, B-R73, B-R74, B-R75, B-R76, B-R77, B-R78, B-R79, B-R80, B-R81, B-R82, A-G3, A-G42, A-G43, B-G6, B-G30, B-G52, B-G66, A X6-Q3, A X6-Q4, A X6-Q5, A X6-Q25, A X6-Q31, A X6-Q32, B X6-Q2, B X6-Q3, B X6-Q26, B X6-Q27, A 11.7 #21, B 11.7 #20, B 11.7 #24 · [GST RULE] ⚠
What this is about: Facts about the two template files that change how the new tool must open, refresh and check them - for example whether a template opens with a "read-only" prompt. Asking only; whether `Info!BB2` is the right proof of a refresh, and which cells are read, are decided in Parts 2 and 6. Please answer: (a) Template format, and do they contain macros? (b) Are they marked read-only on the share? (c) Is any query set to "Refresh data when opening the file" or "Enable background refresh"? (d) What fills `Info!BB2`, and can it hold a value before a refresh? (e) Have cell positions or sheet names changed between template versions, and who updates the code when they do? (f) Does K29 hold a number or a text such as "Match"?
A's way: the template picker offers `.xltx` / `.xlsx` (A-R52); a refresh counts as good if no error pop-up was seen and `Info!BB2` ⚠ is not empty, not 0 and not an error (`A:power_query_extractor/core/report_processor.py:638-685`; A-R73).
B's way: the template picker offers `.xlsx` / `.xlsm` (B-R48); the same refresh test (`B:power_query_extractor/core/report_processor.py:589-636`; B-F20, B-R71).
Evidence: reports are always saved under an `.xlsx` name, whatever the template's type (A-G42, B-G30). The 22 cells read in Step 4 ⚠ (for example `T6 of R9!K29`, "Table 6J Match") are fixed in a code file, `power_query_extractor/config/cell_mappings.py`, identical in both (MERGE_COMPARISON.md section 5, Category 10; B audit 8.10 #10 asks who edits it).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-29. A's screens: what you like and dislike                                  Part 11 · refs A-F2, A-F4, A-F8, A-F10, A-F12, A-F17, A-F18, A-F20, A-F25, A-F28, A-F31, A-R140, A-R148, A-G64
What this is about: Brief Section 7. Name what you would keep and what annoys you every time, from your own use of A (the Tkinter version) - for example the pale progress bar during a 100-client batch. Please answer: what you like, and what annoys you (a line each).
A's way (MERGE_COMPARISON.md sections 6.1 and 6.2): a blue title strip with a Dark Mode tick; one long scrolling Step 1 with a live ✓ / ○ / ✕ checklist and hover help; Step 2 starts unticked, names the missing files ("IMS Reco, Sales Reco") and needs sideways scrolling at 1200×800; the Step 3 log is time-stamped and very detailed, with "[DEBUG]" lines (technical messages meant for the programmer), and the progress bar is pale grey; Step 4 is the old extractor with a folder box, a Scan button and two sub-tabs (the second holds its progress log, A-F25), no Stop button, a failure reason per report, and "Open the consolidated report?" at the end.
B's way: see the next item.
Evidence: dark mode leaves screen parts created later (boxes, buttons) light and makes some text unreadable (A-G64; A-R148); the Step 3 log colours are in A-R140. How Step 2 shows the missing files is a GST rule decided in Part 2 [GST RULE]; here only what you like.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-30. B's screens: what you like and dislike                                  Part 11 · refs B-F1, B-F3, B-F7, B-F8, B-F10, B-F16, B-F18, B-F24, B-F28, B-F29, B-F33, B-R134, B-R136, B-G7, B-G67, B-G68, B-G69, B-G71, B-G73, B X6-Q42, B X6-Q47
What this is about: The same for B (the PyQt5 "v4.0" version), if you have used it; if not, write "not used" or answer from a quick look. Please answer: (a) What you like and dislike. (b) Can you see the three Browse buttons on your PC? (c) Were the "Stop and exit?" question and the "Open the consolidated report?" prompt dropped on purpose in the rewrite?
A's way: see the previous item; A asks "Stop and exit?" when closed during Step 3 and "Open the consolidated report?" at the end of Step 4 (MERGE_COMPARISON.md section 6.1).
B's way (MERGE_COMPARISON.md sections 6.1 and 6.3): File and View menus and coloured dots on the tabs (B-F29); two columns of coloured cards and a dark SCAN LOG on Step 1, with three Browse buttons invisible; Step 2 starts all ticked, counts the missing files ("2"), has Folder Name and Last Run columns and sortable headers (B-R134), and an EXPORT LIST button that does nothing; a green progress bar with a percentage; Step 4 is one page with a client → version tree (a list that folds out from each client to its Version folders) and a Stop button (B-F24), no end question, shows "Refreshed" in green even when every refresh failed (B-R136, B-G7), and keeps failure reasons in the log file only (B-G71).
Evidence: B's rewrite dropped both of those questions, which the shared Tkinter code had (B X6-Q47); the Browse buttons are white on white (B-G67); dark mode darkens only part of the window (B-G69); the client tables lose ticks and sort numbers as text (B-G73). How Step 2 shows the missing files is a GST rule decided in Part 2 [GST RULE]; here only what you like.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-31. How you want the new tool to look and behave                                  Part 11 · refs A-F2, A-F28, A-F31, B-F1, B-F28, B-F33
What this is about: The general preferences of Brief Section 7, beyond A against B (the window kit itself is decided in Part 8) - for example whether an article assistant would rather drag a client's downloads onto the window than browse for a folder. Please answer: (a) One window with step tabs, or something else? (b) Browse buttons, drag-and-drop, or both? (c) Errors shown in the main window, or as pop-ups? (d) If you could change one thing, what? (e) Any complaints from your team? (f) Any other app whose look you find pleasant?
A's way: one Tkinter window with four step tabs (A-F2), 25 message boxes and six file or folder pickers (MERGE_COMPARISON.md section 6.4.7).
B's way: one PyQt5 window with four step tabs (B-F1), 14 message boxes, all with a single OK button, and four pickers (MERGE_COMPARISON.md section 6.4.7).
Evidence: neither accepts files dragged onto the window (READ: no drag-and-drop code in either, by a search of both); errors appear as pop-up boxes and in the on-screen log, and in B's Step 4 only in the log file (MERGE_COMPARISON.md sections 6.1 and 6.4.7). Help texts: A-F31, B-F33; dark mode: A-F28, B-F28.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-32. How the refresh came to be done this way                                  Part 11 · refs A-F21, A-F23, B-F19, B-F22, A-R71, A-R98, A-R103, B-R66, B-R67, B-R101, A X6-Q46, B X6-Q37, A 11.7 #20, B 11.7 #22, B 11.7 #23, B 11.7 #26
What this is about: Your Standard 12 refreshes with Excel's own Refresh All command, called from code, in a hidden Excel, waiting until every query has finished before saving (rebuild_pipeline_FINAL.md:844-878). Both projects instead show Excel and type Alt+A, R, A. If you know why, the build can test the right thing - for example whether a Power Query privacy prompt (Excel's question about how data sources may be combined) appeared when Excel was hidden. Please answer: (a) Did you try Refresh All from code, and what went wrong? (b) What does "Keep it visible to avoid issues" refer to - a pop-up, a hang, old figures? (c) Why are the figures read through a second Excel? (d) Was `12d505f`'s Windows check done, and do locally built reports refresh correctly? (e) Did an earlier version build reports in a temporary folder, and why was that dropped?
A's way: shows Excel ("Keep it visible to avoid issues", `A:power_query_extractor/core/report_processor.py:225`) and types the keys (A-R71, A-F21, A-F23); reads the figures through a second, hidden Excel (A-R103); since `12d505f` (23-07-2026) builds each Step 3 report on the local disk and copies it once, with the note "validate one report on a Windows machine before bulk processing" (A-R98).
B's way: the same refresh (`B:power_query_extractor/core/report_processor.py:181`; B-R66, B-R67, B-F19, B-F22) and the same second Excel (B-R101); builds each report in place, yet creates an unused temporary folder at every run (B audit 8.4.3).
Evidence: the oldest surviving copy (25-06-2025) already used key presses and never Refresh All (B audit 8.4.1). B's unused folder may mean that an earlier version worked through temporary files, or a plan that was never coded (B audit 8.4.3, MEDIUM).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-33. Why there are two versions                                  Part 11 · refs A-F10, A-F12, B-F1, B-F5, B-F10, A X6-Q84, B X6-Q47, B X6-Q60
What this is about: The history splits after 30-10-2025. B was rewritten in PyQt5 on 07-02-2026 (`c258047`, made with an AI assistant); A's repository was created on GitHub the same morning and continued in Tkinter until 27-07-2026 (MERGE_COMPARISON.md sections 1.2 and 10.1). Your reasons matter for the merge - for example, which program the office treats as "the" tool next season. Please answer: (a) Why PyQt5 for the rewrite? (b) Why did the July 2026 work go into A (Tkinter) rather than B - was B set aside on purpose? (c) Was the copilot-worktree branch the same work as B? (d) Were the features B left out dropped on purpose?
A's way: kept the Tkinter window; 34 commits of its own after the split (2 in March 2026, 32 in July 2026); EXPORT LIST (A-F12) and the per-client name switch (A-F10) still work.
B's way: the PyQt5 rewrite and four more commits up to 09-02-2026; the rewrite left out a working EXPORT LIST (B-F10), the per-client name switch (B-F5) and the long-name warning (B audit 8.10 #8); B-F1 is the new window's start-up.
Evidence: GitHub's creation dates make B's repository probably the original and A the later copy (B audit Section 12, before 12A, also noted in its Section 9.0; MEDIUM). A's older audit document says the PyQt5 branch "was subsequently abandoned" and names a branch `copilot-worktree-2026-03-04T11-52-56` that exists nowhere today (A X6-Q84; B X6-Q47, B X6-Q60).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-34. Records that live only with you                                  Part 11 · refs A-F27, A-G60, A X6-Q75, A X6-Q77, A X6-Q78, A X6-Q85, A X6-Q86, A X6-Q88, B X6-Q57, B X6-Q58, B X6-Q59, B X6-Q61
What this is about: Some history is in neither repository. Anything that still exists can help the build (Part 9 asks how to supply files) - for example an older copy of the tool that a colleague still has on a laptop. Please answer: (a) Do the pre-git versions or the uncommitted test files still exist? (b) Were two computers used in 2025? (c) Which AI tools wrote the October 2025 review and README (`6237597`, `d674386`) and B's last two commit messages (`7e09e5c`, `b0e7d0e`)? (d) Did running A's tests ever change your saved settings? (e) Did the old extractor ever fill in its target folder by itself?
A's way: commit `d64d3a9` says "140 tests pass" but holds 132, and the test harnesses described in `0890022`, `275d9a4`, `a603c02`, `12d505f`, `4219ffb`, `56aaad9` and `4f84ba9` were never committed (A X6-Q78); one of A's tests can rewrite your saved settings file (A-G60; A X6-Q77); A's settings file is described in A-F27.
B's way: B's first three own commits end with a line naming an AI assistant as co-author; its last two, `7e09e5c` and `b0e7d0e`, do not (B audit Section 12, before 12A; B X6-Q58).
Evidence (shared history): before git there was a single-file `gui/app.py`, of which only compiled copies (the `.pyc` files Python makes from the code by itself) survive at `edac6a3`; the 2025 compiled files come from Python 3.12 and 3.13, and `47a1492` is under the name "Sidhant" (A X6-Q75, A X6-Q85, A X6-Q88; B X6-Q57, B X6-Q59); `.gitignore` names a web AI tool's working folder, `/mnt/user-data/` (A X6-Q86); the old extractor read its target folder from a settings file in its start folder that the organiser never wrote (B X6-Q61; A changed that path in `c1583d6`).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-35. Two confidentiality facts only you know                                  Part 11 · refs A X6-Q81, MERGE_COMPARISON.md section 10 D3, D5
What this is about: Both old repositories were public on 21-10-2026 (re-checked 23-10-2026). Until you say otherwise, these documents treat both items below as confidential and never copy them - as you would never put a client's GSTIN in a public post. What to do about the old public history is Part 10. Please answer: (a) Is that business name one of your clients, or an example the AI picked? (b) Do you regard the share path - and the run figures quoted from the same commit messages (file counts, scan times, path lengths) - as confidential?
A's way: (1) a company-like business name is used in the worked examples of A's older audit document `PROJECT_AUDIT_ANNUALSTATEMENT.md` (lines 130-232), in A's history since `f7eb8f1` (05-03-2026); (2) the office network share path, with figures from real runs, is in the messages of A's commits `4f84ba9` and `7347383`.
B's way: neither item is in B's files or history (MERGE_COMPARISON.md section 10, D3 and D5).
Evidence: MERGE_COMPARISON.md section 10, D3 and D5 (RAN: searches that printed only hashes - scrambled fingerprints of the text - line numbers and commit ids); A X6-Q81.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-36. What surprised you, and what you had forgotten                                  Part 11 · refs Brief Section 8
What this is about: Brief Section 8 ("What I learned") must be in your words; this item and the next collect it. Please write freely, two or three sentences each. Please answer: (a) What surprised you most in the audits or the comparison? (b) Did they show a feature or rule you had forgotten existed?
A's way: not applicable - this item asks for your own view.
B's way: not applicable.
Evidence: this part does not suggest answers. The comparison's summary is MERGE_COMPARISON.md section 1, and the gaps both projects share are in section 9.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-37. What you would do differently, and the principle above all                                  Part 11 · refs Brief Section 8
What this is about: The second half of Brief Section 8, in your words. Please answer: (a) What would you do differently if you could go back to day one (mid-2025)? (b) Is there a mistake in the old tools you never want to repeat? (c) What is the one principle the rebuild must follow above all?
A's way: not applicable - this item asks for your own view.
B's way: not applicable.
Evidence: not applicable - your own view.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-38. Office-PC time for the checks the build needs                                  Part 11 · refs A-F21, A-F26, B-F19, B-F26, A 11.7 #20, A 11.7 #24, B 11.7 #40
What this is about: Some things cannot be tested on the build computer, because Excel and Power Query run only on Windows with your templates: the side-by-side check of the key-press refresh against Excel's own Refresh All, the list of pop-ups that really appear and which button to press on each, and acceptance sessions (sessions where you try the new tool on real work and accept it or not) - ideally not in the last weeks of the annual-return season. Please answer: (a) Can you give office-PC time for these - roughly how many hours, and when? (b) On which PC (Windows and Excel version as in P11-9 and P11-10)? (c) Should they happen outside the season?
A's way: A's audit plans the side-by-side refresh check and the pop-up list on an office PC (A audit 11.7 #20 and #24; A-F21, A-F26).
B's way: B's audit plans the refresh check, a check of the two ways of reading the figures back (through a second Excel, or straight from the saved file) and three acceptance sessions (B audit 11.7 #40; B-F19, B-F26).
Evidence: A's 288 tests never exercise the Power Query refresh, and B has no tests at all (A audit Section 1; MERGE_COMPARISON.md section 8.1).
Recommendation: No recommendation - your call.
Your answer: ____

### P11-39. Your top three priorities                                  Part 11 · refs Brief "Top 3 Priorities"
What this is about: "If you could only ship THREE things in this rebuild and had to cut everything else, which three would you keep?" (REBUILD_BRIEF_ASSISTANT.md). They go at the top of the Brief and are checked against your other answers (the consistency check made after all your answers, point 3). Please answer: 1. ____ 2. ____ 3. ____
A's way (what the tool does today - not suggestions): sorts each client's downloads into its folder; fills the two reports; refreshes them; collects the 22 figures into one workbook; shows which client needs attention.
B's way: the same five jobs.
Evidence: not applicable - your own priorities.
Recommendation: No recommendation - your call.
Your answer: ____

### P11-40. The one thing that must work perfectly (Zero-Defect Requirement)                                  Part 11 · refs Brief "Zero-Defect Requirement", Part 7
What this is about: "What is the ONE thing that MUST work perfectly - the thing where a bug would be unacceptable?" It goes at the top of the Brief and must be covered by at least one contract in Part 7 (the consistency check made after all your answers, point 4). Please answer: in one sentence.
A's way: no written promise of this kind; A's 288 tests do not exercise the refresh or the consolidated workbook (A audit Section 1).
B's way: no written promise and no tests (MERGE_COMPARISON.md section 8.1).
Evidence (for orientation only, not a suggestion): among the 20 shared gaps the comparison rates HIGH are columns that show another cell's figure than the one configured (H1 [GST RULE]), refreshes judged good without proof (H2, H3) and client files lost or mixed (H6, H7) - MERGE_COMPARISON.md section 9.
Recommendation: No recommendation - your call.
Your answer: ____

#### For the closing list

None. Part 11 holds only the author's own answers about himself, his office and his data; it records no shared, working behaviour of the code, so it adds nothing to "Same in both and working - no decision needed".

#### Refs covered

- **A features:** A-F1, A-F2, A-F3, A-F4, A-F5, A-F6, A-F7, A-F8, A-F9, A-F10, A-F11, A-F12, A-F13, A-F14, A-F15, A-F16, A-F17, A-F18, A-F19, A-F20, A-F21, A-F22, A-F23, A-F24, A-F25, A-F26, A-F27, A-F28, A-F29, A-F30, A-F31 (usage, environment and experience questions only; the keep / drop decisions are in Parts 1, 3 and 4).
- **B features:** B-F1, B-F2, B-F3, B-F4, B-F5, B-F6, B-F7, B-F8, B-F9, B-F10, B-F11, B-F12, B-F13, B-F14, B-F15, B-F16, B-F17, B-F18, B-F19, B-F20, B-F21, B-F22, B-F23, B-F24, B-F25, B-F26, B-F27, B-F28, B-F29, B-F30, B-F31, B-F32, B-F33, B-F34 (same note).
- **A rules:** A-R1, A-R2, A-R3, A-R4, A-R5, A-R6, A-R7, A-R8, A-R9, A-R10, A-R17, A-R19, A-R21, A-R22, A-R25, A-R27, A-R35, A-R39, A-R42, A-R43, A-R45, A-R46, A-R47, A-R52, A-R54, A-R55, A-R56, A-R57, A-R60, A-R67, A-R68, A-R69, A-R71, A-R72, A-R73, A-R74, A-R75, A-R76, A-R77, A-R78, A-R79, A-R80, A-R81, A-R82, A-R88, A-R89, A-R94, A-R95, A-R97, A-R98, A-R100, A-R101, A-R102, A-R103, A-R104, A-R105, A-R109, A-R110, A-R120, A-R124, A-R127, A-R140, A-R148 (facts about real use only; GST-rule versions are decided in Parts 2 and 6).
- **B rules:** B-R1, B-R2, B-R3, B-R5, B-R6, B-R7, B-R8, B-R9, B-R10, B-R11, B-R14, B-R17, B-R19, B-R21, B-R22, B-R23, B-R25, B-R32, B-R35, B-R38, B-R39, B-R42, B-R43, B-R44, B-R48, B-R54, B-R55, B-R56, B-R60, B-R62, B-R63, B-R65, B-R66, B-R67, B-R68, B-R70, B-R71, B-R73, B-R74, B-R75, B-R76, B-R77, B-R78, B-R79, B-R80, B-R81, B-R82, B-R83, B-R88, B-R96, B-R98, B-R99, B-R100, B-R101, B-R102, B-R103, B-R111, B-R118, B-R121, B-R124, B-R125, B-R134, B-R136 (same note).
- **A gaps:** A-G1, A-G2, A-G3, A-G4, A-G5, A-G6, A-G7, A-G8, A-G9, A-G10, A-G15, A-G17, A-G19, A-G22, A-G23, A-G25, A-G27, A-G33, A-G34, A-G35, A-G36, A-G37, A-G38, A-G39, A-G40, A-G41, A-G42, A-G43, A-G44, A-G45, A-G46, A-G48, A-G50, A-G54, A-G55, A-G60, A-G64, A-G71, A-G77, A-G79.
- **B gaps:** B-G1, B-G2, B-G3, B-G4, B-G5, B-G6, B-G7, B-G10, B-G12, B-G13, B-G14, B-G15, B-G16, B-G17, B-G18, B-G19, B-G20, B-G23, B-G27, B-G29, B-G30, B-G32, B-G34, B-G35, B-G39, B-G40, B-G46, B-G47, B-G48, B-G49, B-G51, B-G52, B-G53, B-G54, B-G55, B-G56, B-G57, B-G58, B-G61, B-G62, B-G63, B-G66, B-G67, B-G68, B-G69, B-G71, B-G73, B-G80, B-G82.
- **Shared-gap rows (MERGE_COMPARISON.md section 9):** H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, M1, M2. **Section 10:** D3, D5.
- **A audit Appendix X6 questions asked here (fact side):** A X6-Q1, Q3, Q4, Q5, Q7, Q9, Q11, Q16, Q17, Q18, Q19, Q20, Q21, Q22, Q23, Q24, Q25, Q26, Q27, Q28, Q29, Q31, Q32, Q33, Q39, Q41, Q42, Q43, Q44, Q45, Q46, Q47, Q48, Q50, Q57, Q58, Q59, Q60, Q61, Q62, Q65, Q67, Q68, Q69, Q75, Q76, Q77, Q78, Q81, Q84, Q85, Q86, Q88.
- **B audit Appendix X6 questions asked here (fact side):** B X6-Q1, Q2, Q3, Q6, Q7, Q9, Q13, Q16, Q17, Q18, Q19, Q20, Q22, Q24, Q25, Q26, Q27, Q28, Q29, Q32, Q34, Q35, Q36, Q37, Q38, Q39, Q41, Q42, Q47, Q48, Q50, Q51, Q54, Q57, Q58, Q59, Q60, Q61, Q63.
- **Section 11.7 rows asked here:** A 11.7 #19, #20, #24, #28, #30, #38, #39; B 11.7 #22, #23, #26, #27, #30, #31, #32, #33, #40. **Fact side only** (the decision itself belongs to another part): A 11.7 #2, #3, #8, #9, #10, #11, #12, #15, #17, #18, #21, #25, #26, #27, #29; B 11.7 #2, #3, #6, #8, #9, #10, #11, #12, #14, #15, #19, #20, #21, #24, #28, #36.
