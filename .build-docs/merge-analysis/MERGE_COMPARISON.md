# MERGE COMPARISON: AnnualStatement_oldv1 (A) and AnnualStatement_oldv2 (B)

How the two old versions of the GST File Organizer differ, written for Adv. Sidhant Goyal ("the author") so that he can decide what the new merged tool, "Annual Statement 2.0", keeps. It is the Phase 2 document of `.build-docs/merge-analysis/MERGE_ANALYSIS_PROMPT.md` (section 8); the numbered decision list (`MERGE_DECISIONS.md`, Phase 3) is built from it.

| Item | Value |
|---|---|
| Project A | `goyal-sidhant/AnnualStatement_oldv1`, branch `main`, snapshot `bf2f1dd56f088d019bed6022428aab9bf0eb423d` (27-07-2026): a Tkinter window ("GST File Organizer v3.0"), 46 commits, 288 tests |
| Project B | `goyal-sidhant/AnnualStatement_oldv2`, branch `main`, snapshot `b0e7d0e8853845ce41ba0f0d2df849feed6bc88c` (09-02-2026): a PyQt5 window ("v4.0"), 17 commits, no tests |
| The names mislead | A is not the older one and B is not the newer or better one. They share their first 12 commits (`edac6a3`, 31-08-2025 … `47a1492`, 30-10-2025) and split after `47a1492`. GitHub's records show B's repository was created on 31-08-2025 and A's on 07-02-2026. |
| Date of this comparison | 23-10-2026 |
| Sources | (1) The two final audits: `/home/user/Annual_Statement_2.0/.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv1.md` (audit A, 15,699 lines, filed 22-10-2026) and `/home/user/Annual_Statement_2.0/.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv2.md` (audit B, 13,517 lines, filed 23-10-2026), with their feature, rule and gap numbers `A-F1`-`A-F31`, `A-R1`-`A-R157`, `A-G1`-`A-G79`, `B-F1`-`B-F34`, `B-R1`-`B-R144`, `B-G1`-`B-G82`. (2) The cross-project map `/root/merge-scratch/compare/MAP.md` ("MAP"), which pairs every feature, rule and gap of the two audits and lists every commit made after the split. (3) The code itself, which is the truth wherever a source disagrees: `/root/merge-scratch/A-copy` (= A at `bf2f1dd`) and `/root/merge-scratch/B-copy` (= B at `b0e7d0e`), byte-identical to the repositories, and both histories in `/root/merge-scratch/combined` (refs `A/main`, `B/main`). (4) The author's standards in `/home/user/Annual_Statement_2.0/.build-docs/toolkit/rebuild_pipeline_FINAL.md`: Standard 12 (lines 817-898; "Excel through COM", lines 844-878) and Standard 13 (lines 900-1061). Paths under `/root/merge-scratch/` are scratch copies on the analysis computer, outside every repository; every run can be repeated from the scripts each section names. |
| How it was produced | The map was written first. Sections 3 to 10 were then written by separate writers, each of whom confirmed every claim in the code of both projects (READ), ran the pure-Python parts of both on the same made-up inputs (RAN, on copies, with `python -B`, so nothing was written into the projects), and recorded the places where an audit, a note or the map was wrong. Each section was then re-checked by an independent checker who re-ran its scripts on fresh copies (marked "VERIFY-…"). Sections 1 and 2, the assembly and the final consistency and coverage review (the "Review record" at the end) came last; the review re-read every section against the two final audits, including corrections that audit B received after the sections were written. This computer runs Linux without Excel and without pywin32 (the Windows add-on that lets Python drive Excel), so everything Excel itself would do is READ only; where Excel was needed, stand-in "fake COM" objects recorded the program's own calls, which proves the order of those calls, never what a real Excel does. No client data was used: all clients, names, figures and paths in the examples are invented. |

### Legend

| Mark | Meaning |
|---|---|
| **READ** | Established by reading the code of both projects (or a commit's changes, `git show`). |
| **RAN** | The code was run on made-up inputs on this computer. "RAN (C3)", "RAN (C4)" … name the section whose writer ran it, in `/root/merge-scratch/run/<section>/`; "RAN (VERIFY-C3)" … the independent re-run of that section's checker; "RAN (A audit)", "RAN (B audit)", "RAN (MAP)" a run made by that earlier work and re-read here. |
| `A:path:line`, `B:path:line` | A line of a file in A at `bf2f1dd` or in B at `b0e7d0e`; `path:line@commit` is the same in an older commit. `Std:NNN` (Sections 7 and 8) is line NNN of `rebuild_pipeline_FINAL.md`. |
| **HIGH / MEDIUM / LOW** | How sure an inference is that could not be run. **NOT SURE** always comes with a way to check. |
| **[GST RULE]** | Anything that decides a tax figure; which return, table, sheet or cell a figure belongs to (for example GSTR-9 Table 6J or 8D); a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. When in doubt a point is treated as one. For these the document says what each project does, runs one worked example through both, and says when each version came in; it **never** says which version is right, never recommends one and gives no legal opinion. |
| **⚠** | A value that changes over time and should be checked by the author: return and table names, template sheet names and cell positions, file-naming conventions, the state list, Excel's key tips and path limit, package and Python versions. |
| **same / different / only A / only B** | The relation of a pair of features or rules: identical behaviour and values; something the author would see or get differently; present in one project only. For gaps: "same defect in both", "only A", "only B"; "(in part)" when only some parts are shared. |
| **"works better", "healthier", "closer"** | Technical judgements only: crashes, the safety of Excel and of files, speed, tests and the clarity of messages. Never a judgement on a GST rule; where a row mixes both, the GST part says "GST rule - no judgement" first. |
| `<Client A>`, `ClientA`, `Alpha Test Traders` … | Invented placeholders. No client name, GSTIN, PAN, real file name or figure appears in this document; the run log `gst_organizer.log` that exists in two old commits is described by kind only. |
| **[Corrected during verification - …]**, **[Corrected during review - …]** | A correction made by a section's independent checker, or during the final review of the whole document (all review corrections are listed in the "Review record" at the end). |

**Local labels.** Some labels are numbered only inside one section, and a few letters repeat: Section 3 rows 1-34; Section 4 rows `4A-1`-`4A-42` and `4B-1`-`4B-34`; Section 5 uses the registries' category numbers for its sub-sections (so "Section 5, 9.4" is Category 9, item 9.4); Section 6 rows `D1`-`D66` and screen GST items `G1`-`G9`; Section 7.2 sessions `S1`-`S3` and `F`; Section 9 rows `H1`-`H20`, `M1`-`M40`, `L1`-`L15`; **Section 10 locations `D1`-`D10` (not Section 6's `D1`-`D66`)**; "MAP 2.1 item n" is a worked GST-rule example in the map. The audits' own local labels (for example `U1`-`U30` in audit B's Section 10) are always cited with the audit's name.

### Contents

| Section | What it holds |
|---|---|
| 1. Summary | What each project is today, how they are related, which is more complete, better tested and more recently worked on, and the differences that matter most, with pointers. |
| 2. Family tree | The shared start, the split, and what A and B each did afterwards, phase by phase, with dates and commits. |
| 3. Feature by feature | One row per feature pair (34 rows covering `A-F1`-`A-F31` and `B-F1`-`B-F34`): the difference as the author would notice it, and which works better and why (technical only). |
| 4. Fixes only one side got | Every fix or improvement made after the split in one project (76 rows), checked in the other project's code; the GST-rule rows in detail (4.4). |
| 5. Rules side by side | Every business rule of both projects (`A-R1`-`A-R157`, `B-R1`-`B-R144`) in 16 categories, with a worked example through both for every difference. |
| 6. Screens and workflow | What the author clicks, in order, in each project; what each screen shows; every visible difference (`D1`-`D66`); the same mistakes through both; screen items that are GST rules (`G1`-`G9`). |
| 7. Tech stack against my standards | Language, window toolkit, Excel libraries, Excel through COM rule by rule, Power Query protection, start-up, where files live, tests and packages, against Standards 12 and 13. |
| 8. Code health | Tests, error handling, dead code, silent bugs and file safety, with which project is healthier on each point. |
| 9. Gaps both share | The 75 defects present in both projects (`H1`-`H20`, `M1`-`M40`, `L1`-`L15`), each re-checked in both codes, with when it came in. |
| 10. Data and confidentiality | Client data, personal data and anything secret-looking in either repository or its history, by location only. |
| Review record | What the final review checked, and every correction it made. |

---

## 1. Summary

This section says in plain words what the two old versions of your GST File Organizer are today, how they are related, which one is more complete, better tested and more recently worked on, and which differences matter most for "Annual Statement 2.0". Each point names the section that holds the evidence. Nothing here says which version of a GST rule is right: where a GST rule differs, both versions are only named, and the choice is yours.

**Words used in this section** (each is explained again where the later sections first need it).

| Word | Meaning |
|---|---|
| Commit | One saved change in a project's history, with a date, a message and a short code such as `47a1492`. |
| Snapshot | The last commit of each project: the state that was audited and compared. |
| Split | The point after which the two projects stopped sharing commits (after `47a1492`, 30-10-2025). |
| Tkinter, PyQt5 | Two *window toolkits* (libraries that draw windows, buttons and lists). Tkinter comes with Python; PyQt5 is a separate add-on built on Qt 5. Your Standard 12 names a third one, PySide6 (Qt 6). |
| pywin32, COM | pywin32 is the Windows add-on that lets Python drive Microsoft Excel through COM (Windows' way for one program to remote-control another). |
| Power Query, refresh | Excel's data-import tool inside your ITC and Sales templates; a *refresh* re-runs its queries so the report reads the client's newly copied files. |
| openpyxl | A Python library that writes `.xlsx` files without Excel. When it saves an existing workbook it drops Power Query. |
| Links sheet | The sheet in each template into which Step 3 writes folder paths and file names for the template's queries to read. |
| Step 1 … Step 4 | The four tabs of both windows: Setup (choose the folders and templates, scan), Validation (the client list), Processing (sort the files and build each client's ITC and Sales reports), Extract (refresh the reports and collect their figures). |
| Run folder, Version folder, category folders | `Annual Statement-<DDMMYY HHMM>` is one batch (run folder); inside it each client and state has a folder holding one `Version-<DDMMYY HHMM>` folder per processing run, which holds three category folders for the GSTR-3B, ITC and Sales source files. |
| Consolidated workbook | Step 4's single workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` with the figures read from every client's refreshed reports. |
| State code | The two-letter (sometimes made-up three-letter) code the program makes from the state part of a file name, for example Delhi → DL; it goes into the client folder and report names. |
| Dry Run, Skip refresh | Dry Run = a preview of what Step 3 would do; Skip refresh = a Step 4 tick that reuses an earlier refreshed copy instead of refreshing again. |
| Background thread | A second line of work beside the window, so that the window keeps responding while files are processed. |
| Launcher, error hook, `%LOCALAPPDATA%` | Terms of your Standard 13: the `Run [App Name].bat` file that starts the app without a console window; Python's last-chance handler that records an error nothing else caught; Windows' per-user folder for program settings and logs. |
| Feature, rule, gap | The audits' numbered lists: features `A-F#` / `B-F#` (what the program does), business rules `A-R#` / `B-R#` (fixed decisions in the code), gaps `A-G#` / `B-G#` (defects and situations the code does not handle). |
| Test | A small program that runs part of the real code on known inputs and checks the result automatically. |
| [GST RULE], ⚠ | A GST rule decides a tax figure, which return, table, sheet or cell a figure belongs to, a rate, limit or threshold, rounding, the financial year or tax period, GSTIN, PAN or state-code logic, ITC eligibility or reversal, or a due date. ⚠ marks a value that changes over time (table numbers, template cell positions, file-naming conventions, state lists, versions). |
| READ, RAN, HIGH / MEDIUM / LOW | READ = established by reading the code of both projects; RAN = the code was run on made-up inputs on this computer; HIGH / MEDIUM / LOW = how sure an inference is that could not be run. |

### 1.1 The two versions today

| | A = `goyal-sidhant/AnnualStatement_oldv1` | B = `goyal-sidhant/AnnualStatement_oldv2` |
|---|---|---|
| Snapshot | `bf2f1dd`, 27-07-2026 07:48 IST, "Fix scrolling dying after visiting other tabs; add hover help" | `b0e7d0e`, 09-02-2026 08:42 IST, "feat: Enhance extraction process with logging and statistics tracking" |
| History | 46 commits = 12 shared + 34 of its own: 2 in March 2026 (documents only) and 32 in July 2026, every one of the 32 made with an AI coding assistant (its "Co-Authored-By" line) | 17 commits = 12 shared + 5 of its own, 07-02-2026 to 09-02-2026, the first three made with an AI coding assistant |
| What it is | "GST File Organizer v3.0": one Tkinter window with four tabs, Setup, Validation, Processing and "Step 4: Extract" | "GST File Organizer v4.0": one PyQt5 window with four tabs, the last "Step 4: Extract & Refresh" |
| Size | 70 tracked files: 43 program files (8,930 lines), 19 test files (2,163 lines) and 8 others | 41 tracked files: 38 program files of 7,821 lines; no test file |
| Features (each audit's Section 3) | 31 (`A-F1` … `A-F31`); every one works, every one "with defects" | 34 (`B-F1` … `B-F34`); 30 work with defects, 2 are dead (`B-F5` per-client folder-name option, `B-F10` EXPORT LIST), 2 are broken (`B-F9` the problem-file list, `B-F31` the old extractor entry point) |
| Business rules (Section 5) | 157 (`A-R1` … `A-R157`), 44 tagged [GST RULE] | 144 (`B-R1` … `B-R144`), 45 tagged [GST RULE] |
| Gaps (Section 7) | 79 (`A-G1` … `A-G79`), 13 rated HIGH | 82 (`B-G1` … `B-G82`), 20 rated HIGH |
| Tests | 288 tests; on this computer 287 pass and 1 fails (a test of Windows path splitting that can pass only on Windows) | none, at any of its 17 commits |
| On a PC without pywin32 | starts; on Windows it refuses to start, says why and offers to install pywin32 | cannot open its window at all, on any PC; its `requirements.txt` (the list of packages to install) lists pywin32 only as a comment |
| State, as the audits judge it | complete and in real use, carrying many defects (MEDIUM-HIGH; A audit Section 1) | an unfinished rewrite, dormant since 09-02-2026 (HIGH for the facts, MEDIUM that it was set aside rather than paused; B audit Section 1) |

Evidence: RAN `git rev-list --count`, `git ls-tree` and `git log` on `/root/merge-scratch/combined` (46 / 34 and 17 / 5 commits; 70 and 41 files; 32 and 3 AI "Co-Authored-By" lines); line counts from Section 8.7 item 3 (RAN C78); A audit Section 1 and Appendix X3; B audit Section 1 and Appendix X3. The pywin32 difference: `B:power_query_extractor/core/report_processor.py:11-15` imports pywin32 with no guard, `A:power_query_extractor/core/report_processor.py:20-36` inside a guard (READ here; RAN in Section 3 check C3-1, Section 4 row 4A-32 and Section 5, 9.2). `A:requirements.txt:9` `pywin32>=305; sys_platform == 'win32'` against `B:requirements.txt:13-14` `# pywin32>=306` (READ).

### 1.2 How they are related

**One start, two paths.** Both projects hold the same first 12 commits, from the first upload `edac6a3` (31-08-2025) to `47a1492` (30-10-2025), and split after `47a1492` (RAN `git merge-base A/main B/main`, git's command for the last commit two lines of work share → `47a1492`). Three months later B was rewritten in PyQt5 in three days (07-02-2026 to 09-02-2026) and then left alone. A was left alone until July 2026 and then reworked in five days (23-07-2026 to 27-07-2026). Since the split A changed 47 files (6,028 lines added, 852 removed) and B 49 files (4,056 added, 3,535 removed); between the two snapshots 82 files differ (RAN `git diff --shortstat`, git's count of changed files and lines). Section 2 draws the family tree.

**The business core is still largely the same code.** The parts that recognise files, organise them, fill the Links cells, refresh the reports and read the figures were barely touched by B and were tuned, not rewritten, by A: 52 functions are word for word the same in both, 14 differ only in comments, log lines or the kind of error caught, 16 really differ and 4 exist only in A; the extraction cell list `power_query_extractor/config/cell_mappings.py` is byte for byte the same file (RAN, cross-project map `/root/merge-scratch/compare/MAP.md`). B rewrote the windows; A made the scan fast on the network share, added the tests, rebuilt Step 1 and brought the Power Query Extractor into the main window as Step 4. So most defects are shared as well: 73 of the 98 gap pairs are the same defect in both (Section 9).

**Which repository is the original.** GitHub's records say B's repository was created on 31-08-2025, three minutes before the first commit, and A's on 07-02-2026, six minutes before B's PyQt5 rewrite `c258047` (RAN in the B audit and in Section 10.1). So B is probably the original repository and A a second repository made that morning to keep the Tkinter version (MEDIUM; B audit Section 12A, "Which repository is the original"). The A audit's Section 1 and one row of its 12A say the opposite, from the old name `goyal-sidhant/AnnualStatement` alone; its own Section 12 has the creation dates (see the review record at the end). Either way there is no third project: both share the same 12 commits.

**The names mislead.** "v3.0" (A) has 34 commits of its own after the split and the later dates; "v4.0" (B) has 5. Neither label tells which version is newer (Section 6, D3).

### 1.3 More complete, better tested, more recently worked on

These are technical judgements only (crashes, safety of Excel and files, speed, tests, clear messages); none of them is a judgement on a GST rule.

| Question | Answer | Why | Confidence |
|---|---|---|---|
| More recently worked on | **A** | Last commit 27-07-2026 against 09-02-2026; 34 commits of its own against 5. Five of A's July commits fix problems their messages call "Reported" from use (`2acffed`, `2aaa2b1`, `725fbaf`, `169c7b5`, `bf2f1dd`; RAN `git log`). Nothing was committed or pushed to B after 09-02-2026 (B audit Section 1, RAN). | HIGH (facts) |
| Better tested | **A**, clearly | 288 tests against none. A's tests are narrow: they do not touch organising the files, building the reports, the refresh or the consolidated workbook (A audit X3.5), and 16 of 30 old bugs put back on purpose went unnoticed (Section 8.1). A's pure-logic tests run against B's code give 66 passes and 7 failures, and each failure marks a real difference between the two (Section 8.1; Section 4, row 4A-9). | HIGH |
| More complete | **A** | Every A feature works, with defects. B lost three features in its rewrite (the client-list export, the per-client folder-name switch, the problem-file list in the summary workbook; Section 3.3), cannot start without pywin32 (Section 3, row 1), and leaves an unfilled template under a report's name when a report fails (Section 4, row 4A-2). B has two features A lacks: the one-time import of old settings and the run history (Section 3, rows 32-33). | HIGH for the counts; MEDIUM as an overall verdict |
| Healthier code | **Mixed** | Tests and dead code: A. Error handling and background work: B (no catch-everything error handlers; workers that never touch the window). Silent bugs and file safety: about even, A by a narrow margin. Your Standards 12 and 13: neither complies; B is nearer on the window toolkit and start-up, A on tests and on stating its packages (Sections 7 and 8.6). | MEDIUM |

### 1.4 The differences that matter most to your work

**(a) GST rules that the two versions apply differently - described only; the choice is yours.** The cross-project map lists eleven such items (MAP 2.1). Sections 4 and 5 tag a few more "when in doubt"; they are marked * below. Each one has a worked example through both versions and the commit that brought it in, at the place named.

1. **IMS Reco files with a browser "(1)" after the date** ⚠ (`A-R2` / `B-R7`): A counts them as IMS Reco files since `5c8ce95` (23-07-2026) and, when both copies exist, links the "(1)" copy in Links `B11`; B rejects them (Section 4, 4.4.1; Section 5, 1.1).
2. **The three category folder names** ⚠, which are also the folder text written into the Links cells (`A-R39`, `A-R41`, `A-R54`, `A-R55` / `B-R35`, `B-R37`, `B-R54`, `B-R55`): A `GSTR3B`, `ITC`, `Sales` since `7347383` (26-07-2026), B the original longer names. Whether your templates' queries depend on the names is unknown, because the templates are not in either repository (Section 4, 4.4.2; Section 5, 6.2 and 8.3).
3. **The shortest client-folder name, which decides whether the state code stays in the folder name*** (`A-R36` / `B-R32`): A's box takes any typed value since `322717f` (25-07-2026); B's box stays within 15-100 (Section 4, 4.4.3; Section 5, 6.1).
4. **What must be chosen before a Dry Run** (`A-R51` / `B-R48`, `B-R133`): A needs both templates and a target; B only a ticked client (Section 5, 8.1; Section 6, G5).
5. **When the template paths are read** (`A-R60` / `B-R50`): A reads them again for every client during a run; B fixes them when Start is pressed (Section 5, 8.6; Section 6, G6).
6. **Which earlier refreshed copy "Skip refresh" may reuse*** (`A-R68`, `A-R69` / `B-R63`): A follows the suffix you set; B looks only for "_Refreshed_" (Section 5, 9.4).
7. **Smaller items:** whether the hint for an unrecognised file name ever reaches the user (Section 5, 1.6); the example file names shown on Step 1 ⚠ (Section 5, 1.7); how a client's status is shown on Step 2 (Section 5, 5.1); the template types offered* (Section 5, 8.2); a merged Links cell on the openpyxl path ⚠ (Section 4, 4.4.7; Section 5, 8.4); how Step 4 recognises report files (Section 5, 8.5); the per-client folder-name setting keyed by state code* (Section 4, 4.4.4); and the client key handed back by Step 2, where the formula is the same and only A's crash on a client written "0042" differs (Section 4, 4.4.5; Section 5, 2.1).

**(b) GST-rule items that are the same in both but still need your confirmation.** Neither side changed these after the split, so neither version offers an alternative:

- Nine figure columns of the consolidated workbook (IGST, CGST and SGST of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B") show the `Diff GSTR-3B` cells J15, K15 and L15 ⚠ instead of their own configured cells, in both, since `47a1492` (Section 5, 10.3; Section 9, H1). Which cell each column should show is your call.
- A refresh counts as successful when `Info!BB2` ⚠ holds something other than nothing, 0 or an Excel error text (Section 5, 9.7; Section 9, H3).
- Links cells for file types a client lacks keep the template's own text, and only the first file of a type in name order is linked (Section 9, H5 and M1).
- The six file-name patterns ⚠ (a hyphen inside a client name breaks them), the state-code table ⚠, and the absence of any financial-year, GSTIN or PAN check (Section 5, categories 1 to 4; Section 9, H13 to H16 and M18).

**(c) Technical differences that matter most** (here a "works better" is allowed):

1. **Starting up.** A checks its packages and explains what is missing; B cannot open without pywin32 and advises `pip install PyQt5 openpyxl` (pip = Python's package installer), which does not fix it (Section 3, row 1; Section 4, rows 4A-32 to 4A-36; Section 7.1).
2. **Excel and Power Query safety.** The Excel layer is the same code in both and misses the same Standard 12 rules: it takes over and finally quits your own open Excel, presses Alt+A, R, A ⚠ instead of asking Excel to refresh, and cannot confirm that a refresh worked (Section 7.2; Section 9, H2, H3, H19, H20). Both also keep an openpyxl fallback that saves a report without its Power Query, and without any shapes drawn on the template (text boxes, buttons), while reporting success: A reaches it only when pywin32 cannot be loaded, B's window cannot reach it at all (Section 5, 12.2; Section 7.1 and 7.3). Neither version gives you a compliant Excel layer to keep.
3. **A failed report.** A builds each report in a local temporary copy and copies it to the share once (since `12d505f`, 23-07-2026); B leaves an unfilled template under the report's name, which its Step 4 later lists and refreshes (Section 4, row 4A-2; Section 5, 8.8 and 12.3; Section 8.5).
4. **Which batch Step 4 offers.** A takes the newest `Annual Statement-…` folder by date modified; B the highest name, which puts a batch of 31-12-2025 ahead of one of 21-10-2026 (Section 3, row 19; Section 5, 6.5).
5. **Speed on the office share.** A's scan reads the folder once and checks 16 files at a time (its commit message: about 11 minutes projected, 41 seconds measured, for 2,508 files); B checks one file after another, but in a background thread, so its window stays usable (Section 4, rows 4A-18, 4A-19 and 4B-1; Section 5, 13.2).
6. **Screens.** A has a live Step 1 checklist and hover help (small pop-up hints). B has three invisible Browse buttons, an EXPORT LIST button that does nothing and a summary sheet that always says "No file variations found"; B also has a Step 4 Stop button and a per-run extraction log, which A lacks (Section 3; Section 6.5).
7. **Tech stack.** Neither uses PySide6; B's Qt structure is the nearer starting point. Neither has the Standard 13 launcher, error hooks or settings and logs under `%LOCALAPPDATA%` (Sections 7.1 and 7.4).
8. **Data and confidentiality.** Both repositories are public. Their shared history still holds a real run log (`gst_organizer.log`) and 70 compiled files with Windows paths; A's older audit document holds a company-like name, and two of A's commit messages hold the office share path (Section 10).

### 1.5 What this comparison could not settle

- Nothing that needs Windows or Excel could run on this Linux computer. Everything Excel itself does is READ; runs with stand-in "fake COM" objects prove only the order of each program's own calls.
- The ITC and Sales templates are in neither repository, so whether their queries depend on folder names or cell positions cannot be checked here.
- Which version the office actually used after February 2026 is not recorded anywhere; it is one of the questions for you (cross-project map, "Questions for the author").

**How this section was checked.** Every count and verdict above was taken from Sections 3 to 10 and the two final audits and re-read there (READ). The facts of 1.1 and 1.2 were re-derived for this section (RAN, read-only, in `/root/merge-scratch/combined`): commit counts, the split point, the AI "Co-Authored-By" lines, the "Reported" commits, the file counts and the `--shortstat` figures; the code lines cited in 1.1 and in (c) items 3 and 4 were re-read in both projects (`A:core/excel_handler.py:147-155` against `B:core/excel_handler.py:122-126`; `A:power_query_extractor/gui/extractor_window.py:631` against `B:gui/pyqt_main_window.py:432-436`; `A:utils/constants.py:23, 91-95` against `B:utils/constants.py:23, 85-89`; `A:power_query_extractor/core/data_consolidator.py:141-150` against `B:power_query_extractor/core/data_consolidator.py:129-138`; READ).

## 2. Family tree

The two projects grew from one line of work. The diagram reads from top to bottom in date order; each `●` is a commit (or a group of commits on one day) and `■` marks a snapshot. Dates are India time (IST). Phase names are those of each audit's Section 12A.

```text
 SHARED HISTORY - 12 commits, identical in A and B ("shared - before the split on 30-10-2025")

 31-08-2025 22:11  ● edac6a3  Phase 1  first upload of the finished v3.0 Tkinter program
                   │     └─────► ● cbfa4a5  31-08-2025 22:14  Phase 1  run log deleted on GitHub's web page ─┐
 29-10-2025 17:21  ● 6237597  Phase 2  import and threading fixes (made on a copy that still had the log)   │
 29-10-2025 17:21  ● 9c56a1c  Phase 2  merge that joins the two lines ◄───────────────────────────────────┘
 29-10-2025        ● d674386 → 5ea77c2 → 1cb878b → 01862a0   Phase 2  README, .gitignore, clean-up
 29-10-2025        ● 1c48722 → 4c42758                         Phase 3  fixes from real use
 30-10-2025 09:04  ● bd176bf  Phase 3  Step 4 wait time, file suffix, ITC/Sales choice
 30-10-2025 14:38  ● 47a1492  Phase 3  six more extraction sheets (Table 6J, Table 8D ...)   ◄── SPLIT POINT
                   │
       ┌───────────┴────────────────────────────────────────┐
       │                                                    │
  A = AnnualStatement_oldv1                            B = AnnualStatement_oldv2
  (GitHub repository created 07-02-2026 11:21)         (GitHub repository created 31-08-2025 22:08)
       │                                                    │
       │ (no commit for four months)                        │ (no commit for three months)
       │                                                    ● 07-02-2026 11:27  c258047           Phase 4  PyQt5 rewrite: one four-tab window
       │                                                    ● 07-02-2026 11:29  a62d3cc, 11:57 c21f788
       │                                                    │                   Phase 5  Tkinter files deleted; start-up repaired
       │                                                    ● 08-02-2026 10:52  7e09e5c           Phase 6  v3 look restored, live scan log
       │                                                    ■ 09-02-2026 08:42  b0e7d0e           Phase 7  Step 4 read from the folders on disk,
       │                                                                                                   extraction log, scan statistics, tab dots
       │                                                      = B SNAPSHOT; nothing committed or pushed since
       ● 04-03-2026        fc279af     Phase 4   first AI-written audit (document only)
       ● 05-03-2026        f7eb8f1     Phase 4   the audit rewritten (document only)
       ● 23-07-2026        6 commits   Phase 5   network speed, first correctness fixes
       ● 23-07-2026        4 commits   Phase 6   a test safety net, then fixes guarded by tests
       ● 25-07-2026        5 commits   Phase 7   scan speed on the real share, a new Setup screen
       ● 25-07-2026        5 commits   Phase 8   preparing to bring the Power Query Extractor into the main window
       ● 26-07-2026        3 commits   Phase 9   shorter folder names, the Step 4 tab, the workflow choice
       ● 26-07-2026        5 commits   Phase 10  dark mode rebuilt, the pywin32 start-up chain
       ● 26-07-2026        3 commits   Phase 11  dark-mode corrections from screenshots
       ■ 27-07-2026 07:48  bf2f1dd     Phase 12  scrolling across tabs, hover help
         = A SNAPSHOT; GitHub still showed it on 23-10-2026
```

Evidence: RAN `git log --format='%h %p'` (the parents of each commit: only `9c56a1c` has two, `6237597` and `cbfa4a5`) and `git log --date` on `/root/merge-scratch/combined`; RAN `git ls-remote` of both GitHub repositories on 23-10-2026 (read-only): `main` is still at `bf2f1dd` and `b0e7d0e`; repository creation times from GitHub's records (RAN in the B audit's summary table and in Section 10.1); phase names READ in the A audit's Section 12A (Phase headings at its lines 10702-11957) and the B audit's Section 12A (lines 9520-10513).

### 2.1 The shared start (Phases 1-3, both projects)

| Phase | Dates | Commits | What happened |
|---|---|---|---|
| 1 First upload of the finished v3.0 program | 31-08-2025 | `edac6a3` 22:11 "Initial commit"; `cbfa4a5` 22:14 "Delete gst_organizer.log" | A program already built and used outside git for about ten weeks was uploaded in one commit, with its run log and its compiled files (Section 10, D1 and D2). The log was deleted on GitHub three minutes later. |
| 2 Code review, README rewrite and clean-up | 29-10-2025 | `6237597` "Fix critical import and threading bugs in GST Organizer"; `9c56a1c` "Merge branch 'main' of https://github.com/goyal-sidhant/AnnualStatement"; `d674386` README; `5ea77c2` shebang line removed (a first line used only on Linux and macOS); `1cb878b` `.gitignore` added (the list of files git should not track); `01862a0` compiled files removed (copies of the code that Python makes by itself) | Fixes made on a copy that still held the log, then merged; the README rewritten; the repository cleaned up. |
| 3 Fixes and extensions driven by real use | 29-10-2025 to 30-10-2025 | `1c48722` row colours; `4c42758` mouse-wheel scrolling; `bd176bf` Step 4 wait time, file suffix and ITC/Sales report choice; `47a1492` "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets" [GST RULE] ⚠ | The last shared commit. Its six new extraction sheets also switched on the nine-column mismatch that both projects still have (Section 9, H1). |

### 2.2 What A did after the split (Phases 4-12, 34 commits)

| Phase | Dates | Commits | What happened |
|---|---|---|---|
| 4 Two AI-written audits of the program as it stood at the split | 04-03-2026 to 05-03-2026 | `fc279af` "Project Audit"; `f7eb8f1` "Project Audit 2.0" | Documents only (`PROJECT_AUDIT_ANNUALSTATEMENT.md`). Its list of "Known Improvements" became much of the July to-do list. |
| 5 Network speed and the first correctness fixes | 23-07-2026 00:22-01:54 | `a603c02`, `12d505f`, `275d9a4`, `5c8ce95`, `c1583d6`, `0890022` | One folder listing per scan; reports built on local disk and copied once; the IMS Reco "(N)" names [GST RULE] ⚠; the extractor's settings path and newest-version choice; log lines that could show the wrong client (a "late-binding" fault) (Section 4, rows 4A-1 to 4A-8). |
| 6 A safety net of tests, then fixes guarded by tests | 23-07-2026 07:58-08:54 | `6c1e1c6`, `33d0baf`, `064c634`, `ba2a753` | The first 81 tests; whole-word "Private"/"Limited" cleaning; Start guarded against a second press; the per-client folder setting keyed by state code [GST RULE]; "[DEBUG]" lines out of the log file; pywin32 added to `requirements.txt` (rows 4A-9 to 4A-16). |
| 7 Scan speed on the real share, and a new Setup screen | 25-07-2026 18:45-19:31 | `4219ffb`, `56aaad9`, `4f84ba9`, `322717f`, `d64d3a9` | Folder-listing details reused and Excel signatures (the first 8 bytes that show a file really is an Excel file) read 16 at a time; the live Step 1 checklist (which also dropped the old check on what is typed into the length box) (rows 4A-17 to 4A-22). |
| 8 Preparing to bring the Power Query Extractor into the main window | 25-07-2026 22:39-23:11 | `b276542`, `c1f425e`, `2acffed`, `a677670`, `2aaa2b1` | A design note; Step 4's screen updates moved onto the window's own thread; refresh times re-read after a run; the extractor split into a reusable panel; refreshed copies recognised by the configured suffix (rows 4A-23 to 4A-26). |
| 9 Shorter folder names, the Step 4 tab and the workflow choice | 26-07-2026 08:20-22:58 | `7347383`, `bfcb24a`, `510f7b4` | The category folders shortened for Excel's path limit [GST RULE] ⚠; Step 4 became a tab of the main window; "organise only" or "full pipeline" (rows 4A-27 to 4A-30). |
| 10 Dark mode rebuilt, and the pywin32 start-up chain | 26-07-2026 23:04-23:27 | `aeb83cf`, `c463256`, `f363946`, `be6418c`, `a5ad487` | Dark mode that restores the exact colours; start without pywin32, then a start-up package check, then pywin32 required on Windows "so reports cannot lose Power Query", then an offer to install (rows 4A-31 to 4A-36). |
| 11 Dark-mode corrections driven by screenshots | 26-07-2026 23:37-23:50 | `c89c1d4`, `725fbaf`, `169c7b5` | Three fixes, two of them "Reported with screenshots" (rows 4A-37 to 4A-39). |
| 12 The last commit | 27-07-2026 07:48 | `bf2f1dd` (A's snapshot) | Scrolling that died after visiting other tabs; hover help on Step 1; 64 tests guarding the scroll bindings (rows 4A-40 to 4A-42). |

### 2.3 What B did after the split (Phases 4-7, 5 commits)

| Phase | Dates | Commits | What happened |
|---|---|---|---|
| 4 The PyQt5 rewrite - one four-tab window | 07-02-2026 11:27 | `c258047` "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching" | The whole window layer rewritten in PyQt5 in one AI-assisted commit; the organiser and the separate Power Query Extractor became one window; settings moved to `~/.gst_organizer/` with a one-time import of the old file; run history; background threads; every catch-everything error handler (a bare `except:`) given a type. The business core was left almost untouched: no GST rule value (file-name pattern, state code, Links cell, extraction cell) changed, though some GST-tagged behaviour did (Section 4, rows 4B-9 and 4B-10; the lower limit of 5 it gave the client-folder length box was put back to 15 in Phase 6, row 4B-20). Its fixes and improvements are rows 4B-1 to 4B-15. The same commit made the window unable to open without pywin32 and emptied the summary's "Variations" sheet. |
| 5 The Tkinter layer deleted, and start-up repaired | 07-02-2026 11:29-11:57 | `a62d3cc` "Remove old Tkinter GUI files replaced by PyQt5 in v4.0"; `c21f788` "Fix gui/__init__.py to import PyQt5 main window instead of deleted Tkinter module" | The old screens deleted (with them the client-list export and the per-client folder-name switch); the program could not start until the one-line repair 27 minutes later (11:29:56 to 11:57:02; RAN `git log --date=iso`) (row 4B-16). |
| 6 The v3 look restored, a live scan log, and Step 4 filled from the run history | 08-02-2026 10:52 | `7e09e5c` "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections" | Coloured cards in the old Tkinter style (and, with them, three invisible Browse buttons); a live scan log; small and non-Excel files listed instead of dropped; Excel's temporary `~$` lock files skipped; the length box back to 15-100 (rows 4B-17 to 4B-24). |
| 7 Step 4 driven by the folders on disk, an extraction log, scan statistics and tab dots | 09-02-2026 08:42 | `b0e7d0e` (B's snapshot) "feat: Enhance extraction process with logging and statistics tracking" | Step 4 rebuilt for the third time in three days: every Version folder of every client, read from the disk; the consolidated workbook back in the run folder beside a new `Extraction_Log_<YYYYMMDD_HHMM>.txt`; tab dots (rows 4B-25 to 4B-34). Pushed to GitHub five seconds after its time stamp; nothing since. |

**Branches.** Each repository has one branch on GitHub, `main`. The local copies used for this comparison also carry a second name, `claude/pensive-feynman-0mz1aa`, which points at the same snapshot, holds no extra work and was made when the copies were prepared on 21-10-2026, not by you (both audits' Section 12A, "Branch awareness"). The only merge commit in either history is the shared `9c56a1c`.

## 3. Feature by feature

This section puts every feature of project A (A-F1 to A-F31) next to its counterpart in project B (B-F1 to B-F34). It uses the same pairs as the cross-project map (`/root/merge-scratch/compare/MAP.md`, §1). For each pair it says whether the two are the same, what the author would notice while using the tool, and which version works better and why, judged on technical behaviour only. Where the difference is a GST rule, the last column says **"GST rule - no judgement"**, and both sides are only described.

### How to read this section

**The two projects.** A = `goyal-sidhant/AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026). It is a Tkinter window ("Tk" for short: Python's built-in window toolkit) that calls itself "GST File Organizer v3.0" and has 288 tests (small programs that check the code automatically). B = `goyal-sidhant/AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026). It is a PyQt5 window ("Qt" for short: a separate, add-on window toolkit) that calls itself "v4.0" and has no tests. The names mislead: A is not the older project, and B is not the newer or better one. Both share their first 12 commits (a commit is one saved step in a project's history, with a date and a message) and split after `47a1492` (30-10-2025).

**Relation words.** *same* = identical behaviour and values. *different* = the author would see or get something different. *only A* / *only B* = the other project has nothing comparable.

**"Which works better".** This column judges technical behaviour only: crashes, the safety of Excel and of the files, speed, tests, and how clear the messages are. It never judges a GST rule. A **GST rule** is anything that decides:

- a tax figure;
- which return, table, sheet or cell a figure belongs to (for example GSTR-9 Table 6J or 8D ⚠);
- a rate, limit or threshold, or rounding;
- the financial year or tax period;
- GSTIN, PAN or state-code logic;
- ITC eligibility or reversal, or a due date.

When in doubt, a point is treated as a GST rule. Such parts are tagged **[GST RULE]**. **⚠** marks values that change over time (return and table names, file-name conventions, template cell positions, folder names written into templates, state lists).

**Rows that mix both kinds.** When a row has both a GST-rule part and a purely technical part, the last column first says "GST rule - no judgement" for the GST part. It then judges only the technical part, under the label "Technical part only". That judgement is never a vote for either version of the GST rule.

**Evidence marks.**

- `A:path:line` and `B:path:line` point into A at `bf2f1dd` and B at `b0e7d0e`; `path:line@commit` points into older code.
- **READ** = established by reading the code of both projects.
- **RAN** = the code was run on made-up inputs (invented clients such as `ClientA`, `<Client A>`, `ClientB-MH`; no client data). "RAN (C3)" means it was run for this section, in `/root/merge-scratch/run/C3/` with `python -B`; the list is in 3.4. "RAN (VERIFY-C3)" or "re-run in VERIFY-C3" means it was run again, independently, by the check of this section, in `/root/merge-scratch/run/VERIFY-C3/` (3.4, C3-11). "RAN (A audit)" or "RAN (B audit)" means it was run by that audit, and the trace was re-read here.
- **COM** (Windows' way for one program to remote-control another; here Python drives Microsoft Excel) cannot run on this Linux machine. So everything Excel itself would do is READ only. The audits drove the Python code with fake COM objects (stand-ins that only record which calls are made). That proves the order of the calls, never Excel's reaction.
- **HIGH / MEDIUM / LOW** = the confidence of an inference that was not run.

**Pointers.**

- "A audit §3 A-F7 (l. 1457)" = that heading in `/home/user/Annual_Statement_2.0/.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv1.md`.
- "B audit §3 B-F6" = that heading in Section 3 of the B audit. Section files: `/root/merge-scratch/audit/B/S3a.md` for B-F1 to B-F17 and `S3b.md` for B-F18 to B-F34; the assembled file is `/root/merge-scratch/audit/PROJECT_AUDIT_AnnualStatement_oldv2.md`.
- A-R#/B-R# = rules (Section 5 of each audit); A-G#/B-G# = gaps (Section 7); "MAP 2.1 item n" = the worked GST-rule examples in the map.

**Words used below (first use).**

| Word | Meaning |
|---|---|
| pywin32 | The Python add-on that lets Python drive Excel through COM. |
| Power Query | Excel's data-import tool. Its queries live inside the ITC and Sales templates, and "refresh" re-runs them. |
| Links sheet | The sheet in each template into which Step 3 writes folder paths and file names, for the queries to read. |
| openpyxl | A Python library that writes `.xlsx` files without Excel. When it saves an existing workbook, it drops Power Query. |
| Run folder (Level 1) | `Annual Statement-<DDMMYY HHMM>`. |
| Client folder (Level 2) | `<client>-<state code>`. |
| Version folder | `Version-<DDMMYY HHMM>`. |
| Category folders | The three folders inside a Version folder. |
| Variations | The code's word for files it could not recognise or rejected. |
| Date modified | The time the system last recorded a change to a file or folder. |
| Thread | A second line of work running beside the window, so the window does not freeze. A **daemon thread** is killed instantly, without clean-up, when the program ends. |
| Signal | Qt's message from one part of the program to another, used to pass progress safely from a background thread to the window. |
| SendKeys | Pressing keys by program into whatever window has the keyboard. |
| BB2 check | Step 4's only success test: the cell `Info!BB2` ⚠ must not be empty. |
| Glob | A wildcard file search such as `ITC_Report_*.xlsx`. |
| Tooltip | A small pop-up hint that appears when the mouse rests on a control. |

**At a glance.** 34 rows:

- 2 *same* (rows 23, 26);
- 28 *different*;
- 2 *only A* (rows 4, 5);
- 2 *only B* (rows 32, 33).

Nine of the 28 *different* rows contain a GST-rule difference: rows 7, 8, 9, 11, 13, 14, 15, 20 and 31. Between them they hold all eleven GST-rule pairs that MAP 2.1 lists as not the same. Two more rows hold a part that Sections 4 and 5 tag as a GST rule "when in doubt": row 6 (the shortest client-folder length, which decides whether the state code stays in the folder name; Section 4, 4.4.3; Section 5, 6.1) and row 22 (which earlier refreshed copy Skip refresh may reuse; Section 5, 9.4). [Corrected during review - originally only the nine MAP-based rows were named; rows 6 and 22 now carry "GST rule - no judgement" first, like the other mixed rows, so that no verdict in this section can be read as a vote on those two points.] The GST content of rows 6, 19, 21, 23 and 24 is otherwise identical in both projects: the state code in the client-folder name (row 6: the rule that builds the name, A-R35 / B-R32, is the same function; only the limit each box lets through differs), the report-name pattern (row 19; how Step 4 finds the reports differs, row 20; the cleaned client part inside the names differs through name cleaning, row 13, a rule neither registry tags as GST), the BB2 test (row 21), the extraction cells (row 23) and the consolidated look-up (row 24).

On the technical part alone:

- A works better in 11 rows: 1, 3, 9, 12, 13, 14, 16, 20, 22, 30, 31.
- B works better in 6 rows: 6, 15, 24, 25, 27, 34.
- 11 rows are mixed, even, or neither works properly: 2, 7, 8, 10, 11, 17, 18, 19, 21, 28, 29.

The main differences were re-confirmed in the code of both projects for this section (3.4). Several of them were also re-run.

### 3.1 The table

| # | A-F# | B-F# | Relation | The difference as the author would notice it while using the tool | Which works better, and why |
|---|---|---|---|---|---|
| 1 | A-F1 Start-up and dependency checks | B-F1 Start-up | different | **A** prints a package report at start and offers (Yes/No box) to install missing packages and restart. On Windows it refuses to start without pywin32; elsewhere it starts with a warning. Steps 1-3 then still run (off Windows, Step 3 builds the reports with openpyxl, which drops Power Query from them: A:core/excel_handler.py:18-27, A-G7), and Step 4 refuses each report with "pywin32 is not installed…". **B** checks nothing. Without pywin32 it cannot open its window at all, on any PC. The console says "Missing module: No module named 'pythoncom'" and advises `pip install PyQt5 openpyxl`, which does not fix it. B's `requirements.txt` has pywin32 only as a comment (`# pywin32>=306`). | **A.** B cannot start on a PC set up the way its README says (RAN B audit, Case A; RAN (C3), re-run in VERIFY-C3: importing B's window module (the code file that builds the window) fails with `ModuleNotFoundError: No module named 'pythoncom'`, while A's imports with `WIN32_AVAILABLE = False`). Cause: B's unguarded imports (lines that load pywin32 with no fallback if it is missing) at B:power_query_extractor/core/report_processor.py:11-15, reached through B:gui/controllers/extraction_controller.py:10; A's guard is at A:…/report_processor.py:20-36. A's checks came in `f363946`, `be6418c` and `a5ad487` (26-07-2026) and are tested (`tests/test_requirements_check.py` 15, `tests/test_dependency_installer.py` 13, `tests/test_optional_pywin32.py` 4 tests). A's own defects: started without a console it dies silently at A:main.py:17 (RAN A audit); the install box names one package but installs the whole `requirements.txt`, pytest (a test tool) included (RAN A audit). Neither project meets Standard 13 (no error hooks, no `.bat`/`pythonw` start, i.e. no double-click launcher that runs Python without a console window). |
| 2 | A-F2 Main window | B-F1, B-F29 Window, tab dots | different | **A:** "GST File Organizer v3.0", 1200×800 (minimum 1000×700), a blue title strip with the Dark Mode tick, status bar "v3.0 \| Production Ready". A asks "Stop and exit?" only while Step 3 runs. **B:** "GST File Organizer v4.0", 1400×900 (minimum 1200×800), File and View menus, and a coloured dot on each tab (orange = ready, green = done). B never asks before closing. Both switch tabs on their own after a scan and when processing starts. | **Mixed.** A asks before closing during Step 3, although the worker is still killed about 1 s later (RAN A audit), and it asks nothing during Step 4. B never asks: closing during Step 4 left an unrefreshed copy and no workbook, log or history (RAN B audit B-F34 Run 2). B's dots turn green after an empty scan, a failed or a stopped run, and are never reset (RAN B audit B-F29), so they can mislead. At A's minimum size the Dark Mode tick and the status bar disappear (RAN A audit). B's minimum size does not fit a 1366×768 laptop screen (A-G63 / B-G68). |
| 3 | A-F3 Choose paths (Setup) | B-F3 Choose paths | different | The same four paths. The template picker offers `*.xltx *.xlsx` in A (A:gui/handlers/file_handler.py:31, 42) and `*.xlsx *.xlsm` in B (B:gui/views/setup_view.py:612, 619). **A** saves the settings after each Browse. Typed paths are not saved until some other action saves. **B** saves at a scan, after processing, on the dark-mode switch, on File > Save Settings and on closing. In B, three of the four Browse buttons are white on white: they work, but cannot be seen. | **A.** All of A's buttons are visible. Its live checklist (row 4) flags a wrong entry while typing. B's invisible buttons were RAN in the B audit (pixel check), cause B:gui/widgets/colored_section.py:62. Both reject a correct folder pasted with the quotes that Explorer's "Copy as path" adds: B with "Folder not found" when SCAN is pressed (RAN B audit), A with ✕ "Folder not found" in its live checklist, which keeps Scan unavailable (RAN (VERIFY-C3), `quote_out.txt`); A at least shows this while typing. A's own defect: text in the length box makes every save fail (RAN A audit). Shared: a folder passes as a template (RAN both audits); the report is always named `.xlsx` whatever the template type (A-G42 / B-G30). |
| 4 | A-F4 Live setup checklist | — (B-F33 has only a fixed card) | only A | A re-checks the four paths 350 ms after typing stops. It shows ✓ / ○ / ✕ per path and a "Ready - all 4 items set" pill, and enables Scan only when the source folder exists. B shows a fixed 7-line INSTRUCTIONS card and checks the paths only when SCAN or START is pressed. | Only A, so no contest. What A has: 19 logic tests and 13 window tests. RAN (C3): the full A suite gives 1 failed, 287 passed; the one failure is a Linux-only path-separator case in this file (it passes on Windows, per the A tests notes). A's defects: a ✕ stays after the folder is created, until restart or a dark-mode switch (RAN A audit). A file passes as a folder (RAN). "Ready" does not stop Step 3. The checks run on the window's thread (MEDIUM: can stall on an unreachable share). |
| 5 | A-F5 Workflow choice | — | only A | A asks "WHAT SHOULD THIS RUN DO?": organise only, or full pipeline. "Full" means that after Step 3, A opens Step 4, fills its folder and lists the clients. B has no choice. After every processing run it rebuilds the Step 4 list from disk and switches to Step 4 (in effect A's "full"). Neither starts the refresh by itself. | Only A, so no contest. A's texts promise both more and less than the code does. "Also refresh Power Query and extract the data": the refresh is not started (RAN A audit). "Fast - no Excel automation": Step 3 drives Excel on Windows. The hand-over runs even after Stop, and it breaks the "Step 4 follows Step 1" link (RAN A audit). There are 7 tests (`tests/test_workflow_mode.py`). |
| 6 | A-F6 Max client-folder length | B-F4 Name option and length | different | Both have arrows from 15 to 100 and a default of 35, and both shorten only the client folder (Level 2). The shortening code is identical (`create_client_state_key`). **A**'s box is always active and takes any typed text. **B**'s box is greyed out unless "Include client name" is ticked, yet its value is applied anyway; Qt accepts only whole numbers in range. | **GST rule - no judgement** on the shortest client-folder length the tool should allow, which decides whether the state code stays in the folder name [GST RULE] (A-R36 / B-R32, B-R138; Section 4, 4.4.3; Section 5, 6.1). **Technical part only: B**, for input safety: it judges only that A's box takes text and values outside its own arrows' range, not which minimum is right. [Corrected during review - the GST-rule label was added; originally the row said only "B, for input safety", while Sections 4 and 5 treat the length limit as a GST rule through B-R32.] In A, an empty or non-number box makes every settings save fail and stops Step 3 with "Fatal error" (RAN A audit). A typed value of 8 or less drops the state code from the folder name (RAN A audit; re-run in VERIFY-C3, `key_length_out.txt`). A's own commit `322717f` (25-07-2026) removed the earlier input check (A audit history). B cannot be given bad text, but its greyed-out box still decides the length (RAN B audit), and a migrated 0 becomes 15 (RAN B audit B-F2). The Level-2 name carries the state code (A-R35 [GST RULE] ⚠). That rule is identical in both; only the input handling is compared here. |
| 7 | A-F7 Scan and recognise files | B-F6 Scan and recognise files | different | The same six name patterns ⚠, except one. An IMS Reco name with "(N)" after the date (a browser's duplicate download) is accepted by A and rejected by B as "No matching pattern found". **A** drops files under 1 KB or without an Excel signature (the first bytes of a file, which show its real type) without a word. It reads the folder once and checks signatures 16 at a time, all on the window's own thread. **B** lists such files as variations (shown only as a number). It skips Excel's `~$` lock files (small hidden files Excel creates while a workbook is open) by name and scans in a background thread. On Linux, B misses an upper-case `.XLSX`; on Windows both find it (HIGH). RAN (C3), re-run in VERIFY-C3 with identical output, same made-up folder: A listed 4 files, including the IMS "(1)", with 0 variations; the 500-byte file vanished. B listed 3, with 2 variations (the IMS "(1)" and "File too small"), and missed the `.XLSX`. | **GST rule - no judgement** on the IMS "(N)" acceptance [GST RULE] ⚠ (A-R2 / B-R7, MAP 2.1 item 1). A changed it in `5c8ce95` (23-07-2026); B keeps the original pattern from `edac6a3`. **Technical part only: mixed.** B keeps the window responsive (thread) and at least counts rejected files. A reads a network share far faster: A's commit `4f84ba9` reports 41 s measured for 2508 files on the office share, against about 11 min (677 s) projected for the earlier one-at-a-time signature reads, and `56aaad9` went from 6.2 to 1.2 round trips (separate requests to the share) per file. A's scan is tested (`tests/test_scan.py` 5, `tests/test_file_ops.py` 22, `tests/test_file_patterns.py` 21 tests). A's `test_file_ops.py` cannot even be imported against B (RAN (C3)). B's per-file lines go into an HTML log (HTML = a web-page text format) that slows as it grows (RAN B audit: 1,000 lines 5.6 s). Shared: sub-folders ignored and `.xlsm` never recognised (A-G37 / B-G46); hyphenated client names mis-split [GST RULE] (A-G33 / B-G17). |
| 8 | A-F8 Client list (Validation) | B-F8, B-F7 Client list, scan log and summary | different | The status values are identical (same function). **A:** rows start unticked; the Missing column names the missing types; a "📁 Name" column toggles the client name per row; rows are tinted. Tk hands back number-looking text as a number, so a client written "0042" breaks Dry Run and Start. **B:** every row starts ticked; extra columns "Folder Name" (the scan key, not the real folder) and "Last Run"; Missing and Extra are counts only; clicking a header sorts, with numbers sorted as text, and processing follows the screen order. B keeps the key with the row. It also writes a scan log on Step 1 and an HTML scan summary on Step 2. | **GST rule - no judgement** on the status and client-key rules [GST RULE] (A-R17 / B-R17, A-R29 / B-R27, MAP 2.1 items 4-5). The key built at the scan (`<client as written>-<state code>`) and the status texts are identical in both. The map lists the two pairs as not the same because of the way back from the table (A rebuilds the key from the screen text, B keeps it) and because of the display: A tints "Has duplicates" like "Missing", B gives it a third colour. "Complete Only" ticks the same rows in both with today's status texts (READ A:gui/handlers/client_handler.py:152 contains, B:gui/views/validation_view.py:446 equals). **Technical part only: mixed.** B returns the stored key unchanged. A rebuilds the key from the table text and stops the whole run on "0042" (RAN A audit; A-G25, only A). Only that crash is judged here; the key form itself (`0042-MH` in both at the scan) is not. A's default selection is safer: B ticks every client, incomplete ones too, so START processes all of them (RAN B audit). A shows which types are missing. |
| 9 | A-F9 Variations list | B-F9 Variations list | different | Both collect unrecognised names with an issue text and a suggested pattern (an example of the expected name; same code). **A** writes that list into the summary workbook's "Variations" sheet, though files it dropped as too small or not Excel are never listed. **B** also lists the rejected files, but shows only their number. Its "Variations" sheet always says "No file variations found". | **GST rule - no judgement** on the suggestion texts [GST RULE] ⚠ (A-R12 / B-R12, identical; MAP 2.1 item 2). **Technical part only: A.** A's sheet names the files. B's sheet states the opposite of the truth (RAN B audit: 3 problem files found, sheet "No file variations found"), because B:gui/controllers/processing_controller.py:73 passes `'variations': []` while A:gui/handlers/processing_handler.py:224 passes the real list (READ). This is a regression (something that worked before and was broken by a later change) in B's `c258047`: at the split the Tk code passed the real list (`47a1492:gui/handlers/processing_handler.py:161`). Neither shows the list on screen. |
| 10 | A-F10 Per-client name toggle | B-F4, B-F5 Global option; per-client plumbing | different | **A:** a global tick on Step 1, a per-client "📁 Name" toggle on Step 2, and a long-name warning. **B:** only the global tick. The per-client setting exists inside the organiser, but nothing on screen can switch it on (B:models/app_state.py:83; no writer, RAN B audit). | **Mixed.** A offers the feature; B cannot get out of step. In A, turning the global tick on and then off leaves every per-client flag on. After a scan the table shows "☐ No" while the global tick still adds names. The summary says "No" while per-client names were used (all RAN A audit). Before `33d0baf` (23-07-2026), A's per-client flags were silently ignored (state name against state code key). B's note still promises "individual client settings" and a warning that do not exist (B:gui/views/setup_view.py:385). |
| 11 | A-F11 Dry run | B-F11 Dry run | different | The same preview lines. **A** refuses to preview unless both templates exist and a target is set, and it also prints the mode and the number of clients. **B** checks only that a client is ticked; with empty template and target boxes it still promises two reports per client. | **GST rule - no judgement** on what must be chosen before Dry Run [GST RULE] (A-R51 / B-R48 / B-R133, MAP 2.1 item 7). The gate dates from `edac6a3` (shared); B's PyQt dry run (`c258047`) has none. **Technical part only: even.** A adds the mode and count lines. In both, after a run the preview shows the cleaned copy names instead of the source names (RAN both audits). Both promise "Would create: Sales Report" for a client without Sales files. |
| 12 | A-F12 Export client list | B-F10 EXPORT LIST button | different | **A** saves the ticked clients to a text file and opens it. Opening works on Windows only; elsewhere the user sees "Export Complete" followed by "Export Error". **B**'s "EXPORT LIST" button is connected to nothing. | **A.** B's button has 0 receivers (nothing is connected to its click; RAN B audit; READ here: `btn_export` is never connected, B:gui/views/validation_view.py:186-221). B's `a62d3cc` deleted the Tk handler, while its message says "All functionality has been reimplemented". A's own defects: the file is written in the Windows code page, and the error after the export is off Windows only (A-G31). |
| 13 | A-F13 Organise files into folders | B-F12 Organise files into folders | different | The same folder tree, modes, backups and copy method. The organiser code is identical, apart from A working out the run folder once per run, and both give the same results in all modes (RAN MAP probe). Only the three category folder names ⚠ differ: A `GSTR3B`, `ITC`, `Sales`; B `GSTR-3B Exports`, `Other ITC related files`, `Sales related files` (in both, followed by ` (<client>)` when the include-name option is on; READ A:utils/constants.py:91-95, B:utils/constants.py:85-89). Copied file names, and the client part of the report names (`ITC_Report_<cleaned client>_…`, built with the same cleaning: A:gui/handlers/processing_handler.py:292-297, B:gui/controllers/processing_controller.py:112-127), also differ when "Private"/"Limited" is written in capitals or inside a word. RAN (C3), re-run in VERIFY-C3 with identical output: "ClientA PRIVATE LIMITED" gives `ClientA_Pvt_Ltd` in A and `ClientA_PRIVATE_LIMITED` in B; "Privateer Unlimited Co" stays unchanged in A and becomes `Pvter_Unlimited_Co` in B. | **GST rule - no judgement** on the category folder names [GST RULE] ⚠ (A-R41 / B-R37; the same names are written into the Links cells, A-R54 / A-R55; MAP 2.1 items 6 and 8). A changed them in `7347383` (26-07-2026), and its message gives a path-length reason (a real run of 3030 files, worst path 270 characters, 1094 over Excel's ~218 limit). Whether the templates' queries depend on the names is unknown (the templates are not in the repositories). **Technical part only: A.** A's name cleaning (`33d0baf`, `4219ffb`) handles any letter case and whole words only, and for names written with spaces it matches its own client-folder abbreviation. For underscore-joined names it does not, in either project: `ClientA_Private_Limited` gives copies `ClientA_Pvt_Ltd…` but the client folder `ClientA_Private_Limited-DL` in both (RAN (C3) `sanitize_out.txt`, re-run in VERIFY-C3). 3 of A's tests for it fail against B's code (RAN (C3), re-run in VERIFY-C3). Shared defects: Resume does not resume, same-minute runs overwrite, and two clients can share one folder (A-G10 / B-G13, A-G9 / B-G27, A-G8 / B-G15). |
| 14 | A-F14 Create ITC report | B-F13 Create ITC report | different | The same Links cells ⚠ and the same Excel steps. **A** fills each report in a local temporary copy (a "staging" copy) and copies the finished file to the target once. **B** copies the template straight to the report's final name and fills it there. In B, a failed fill leaves an unfilled template under the report's name, which Step 4 later treats as a report. | **GST rule - no judgement** on the Links contents [GST RULE] ⚠. The cells and sources are identical; only the folder text differs, through row 13. On the openpyxl path, A skips a merged Links cell silently while B fails the report (A-R58 / B-R53, MAP 2.1 item 9); B's window cannot reach that path. **Technical part only: A** (READ A:core/excel_handler.py:135-155, 212-215 against B:core/excel_handler.py:122-129). A's `12d505f` (23-07-2026) gives two reasons: fewer slow network round trips, and "the destination is no longer left with a half-written template copy". But a final copy broken by a network drop still leaves a partial file (A-G12). B's leftover template was RAN in the B audit (fake COM, Case D). Shared and serious: both take over, hide and quit the user's own open Excel (A-G54 / B-G63). Both report success when the template has no Links sheet (A-G4 / B-G4). |
| 15 | A-F15 Create Sales report | B-F14 Create Sales report | different | As row 14. In addition, A shows 14 "[DEBUG]" lines in the Sales part of the Step 3 log (8 of them for every client in a normal run). A also re-reads the template paths from the Step 1 boxes for every client. B uses the paths fixed when Start was pressed. | **GST rule - no judgement** on which template a client's report is built from if the box is changed mid-run [GST RULE] (A-R60 / B-R50, MAP 2.1 item 11). A's live re-read is the shared Tk code; B's snapshot dates from `c258047`. **Technical part only: B,** for log clarity. B's `c258047` removed the user-visible "[DEBUG]" lines. A's `064c634` left its 14 lines on purpose ("a separate UX decision"); A has 14 (READ, grep), B none in its window code. Shared: a client with no Sales-type files gets an unfilled Sales report counted as a success [GST RULE] (A-G5 / B-G2). |
| 16 | A-F16 Summary workbook and text report | B-F15 Summary workbook and text report | different | The same five-sheet summary workbook and per-client text report (identical code and identical output from the same input, RAN MAP probe). B's "Variations" sheet is always empty. If the summary cannot be written, A logs "❌ Failed to create summary report"; B shows nothing. | **A.** B is silent when the summary fails (RAN B audit Case D; READ B:gui/controllers/processing_controller.py:189-193 logs only on success, while A:gui/handlers/processing_handler.py:469-475 logs both outcomes). B's Variations sheet is false (row 9). Shared: the text report shows whole-run running totals; Client Status lists only successful clients; a client is "Success" even if both reports failed (A-G19 / B-G19). |
| 17 | A-F17 Progress and log | B-F16 Progress and log | different | Both have a progress bar, an ETA (estimated time left) from the same tracker, a coloured log, Stop only between clients, and "Successfully processed p/t clients" at the end. **A:** the worker writes to the window directly; lines carry `[HH:MM:SS]`; the log is never cleared; the run folder opens on any system. **B:** messages go through Qt signals; there is a Clear Log button; Explorer opens on Windows only; then B switches to Step 4. B's log is HTML (a web-page text format), so text between "<" and ">" vanishes, as does indentation. | **Mixed.** B's design is sounder: its worker never touches the window (RAN B audit thread names). A's worker raises the completion box itself (RAN A audit; A-G70). A's own `c1f425e` says Tk "is not thread-safe" (its window may safely be changed only from the window's own thread) and fixed this for Step 4 only. A's log is more faithful: B lost the rest of an error message after "<" (RAN B audit) and its log slows sharply as it grows (2,000 lines ≈ 49 s, RAN B audit). A asks before closing during Step 3. Both report a stopped run as "Successfully processed p/t" (A-G69 / B-G70). |
| 18 | A-F18 Step 4 inside the main window | B-F18, B-F24 Step 4 page | different | **A** embeds the extractor panel as Step 4, with its own "Setup" and "Processing" sub-tabs. Its folder follows Step 1's target as typed, and a notice appears if the panel cannot load. **B**'s Step 4 is one native page. It reads the last SAVED target and fills itself from disk when the window opens. | **Mixed.** In A, the panel's status messages are invisible inside the main window, and after a "full" hand-over Step 4 stops following Step 1 (both RAN A audit). The notice blames pywin32 for any import failure (RAN). In B, a newly typed but unsaved target is ignored (RAN B audit). B walks the folders on the window's thread at start-up, after processing and after an extraction that ends normally or is stopped, not after one that ends in an error (READ `B:gui/pyqt_main_window.py:549-565`; MEDIUM on how long). [Corrected during review - originally "after extraction"; Section 4, 4.7 item 2 found the error path does not rebuild the list.] B has one flat page with Stop. A has tabs inside a tab, which its own `bfcb24a` calls a decision still to make. |
| 19 | A-F19 Find reports for Step 4 | B-F18 Find reports for Step 4 | different | Which batch Step 4 shows: A picks the `Annual Statement-…` folder with the newest date modified; B picks the one with the highest NAME. Names are day-first, so 31-12-2025 beats 05-01-2026. Inside a client: A shows only the newest Version folder by date modified; B shows every Version folder that holds a report, newest name first. | **Neither is reliable.** This is a technical rule (A-R42 / B-R39 are not tagged GST). The report names it looks for are identical [GST RULE] (A-R59). RAN (C3), re-run in VERIFY-C3 with identical output, with both selection expressions copied verbatim: for a 31-12-2025 batch and a 05-01-2026 batch, A picks 05-01-2026 and B picks 31-12-2025. After an Excel lock file appears in the older batch, A switches to 31-12-2025 too. B's name rule failed in the B audit as well (RAN). A's own `c1583d6` (23-07-2026) dropped name sorting for Version folders because "day-first text does not sort chronologically across month boundaries"; B still uses name sorting for batches and versions. Which rule fails more often in the office depends on habits (NOT SURE). Neither is a dependable "latest". |
| 20 | A-F20 Step 4 selection and status | B-F18 Step 4 tree and status | different | **A:** one row per client; ITC and Sales ticks are pre-ticked where the report exists; the status reads "ITC: <time> \| Sales: <time>", found with the CONFIGURED suffix. **B:** a tree of versions, all unticked; one status per version, "Refreshed (dd/mm/yy HH:MM)" or "Pending", found by any file whose name contains "_Refreshed". | **GST rule - no judgement** on how reports are recognised [GST RULE] (A-R59 / B-R51 / B-R61, MAP 2.1 item 10). The two differ only for a hand-renamed file whose letter case differs from what the program writes (for example `.XLSX`, or `itc_report_…`) on a case-sensitive disk such as Linux's; on Windows both ignore letter case (HIGH). The program itself always writes `ITC_Report_…xlsx` / `Sales_Report_…xlsx`. **Technical part only: A.** A's status follows the suffix the user set (`2aaa2b1`, with tests) and is re-read after a run (`2acffed`). With a custom suffix, B's status stays "Pending" for ever (RAN B audit; RAN (VERIFY-C3): B's status expression, copied verbatim from B:gui/pyqt_main_window.py:467-479, gives "Pending" for a version folder holding a `_PQ_` copy, `b_status_out.txt`). Shared: a copy left by a failed refresh is shown as refreshed (A-G2 / B-G7, RAN both). |
| 21 | A-F21 Refresh through Power Query | B-F19, B-F20 Refresh and BB2 check | different | The refresh code is identical: copy, visible Excel, the key presses Alt+A, R, A ⚠, fixed waits, the `Info!BB2` ⚠ check, save. The user sees two differences. The wait box runs 5-60 in A, where typed values are not checked, and 5-120 in B, where Qt enforces the range. Without pywin32, A fails each report with a clear message; B cannot start at all. | **Even.** It is the same engine with the same serious defects: the refresh cannot be confirmed (A-G3 / B-G6); keys go to whatever window is in front (A-G55 / B-G62); the user's own Excel may be taken over. B's box cannot take bad text; in A, text in the box makes Start silently do nothing (RAN A audit). The BB2 test [GST RULE] ⚠ is identical (A-R73 / B-R70, `_validate_refresh` IDENTICAL), so no judgement is needed on it. |
| 22 | A-F22 Skip refresh and naming | B-F21 Skip refresh and naming | different | The same default suffix `_Refreshed_{timestamp}` and the same copy name. **A** recognises refreshed copies by the CONFIGURED suffix (plus the old "_refreshed_"). **B** recognises them only by the fixed text "_Refreshed_" for Skip and "_Refreshed" for the status, so in B a custom suffix breaks both Skip refresh and the status. | **GST rule - no judgement** on which earlier refreshed copy Skip refresh may reuse, because that decides which file's figures reach the consolidated workbook (Section 5, 9.4 tags it [GST RULE] when in doubt). **Technical part only: A,** with one caveat: A's search follows the suffix set in its own box, B's ignores its own setting (B-G9). [Corrected during review - the GST-rule label was added to match Section 5, 9.4; the technical verdict is unchanged.] RAN (C3), re-run in VERIFY-C3 with identical output, with suffix `_PQ_{timestamp}`: A's Skip finds the earlier `…_PQ_060126_1015.xlsx`; B's search (B:power_query_extractor/core/report_processor.py:102) finds nothing. A's `2aaa2b1` has 19 + 2 tests. Caveat, RAN (C3): with suffix `_R{timestamp}`, A counts the ORIGINAL report as a refreshed copy (A-G16). Also, A never stores the suffix (A-R119), so after a restart yesterday's custom-suffix copies read "Never refreshed". B stores it (B-F27). Shared: Skip reuses a copy whose refresh failed and reports success (A-G2 / B-G7); the "original" can be an earlier copy (A-G15 / B-G8). |
| 23 | A-F23 Extract cells | B-F22 Extract cells | same | A second, hidden Excel reads the configured cells ⚠ [GST RULE] (22 mappings in an identical `cell_mappings.py`, counted in VERIFY-C3); only the kind of error caught differs. The second Excel is requested with `Dispatch` (A:power_query_extractor/core/report_processor.py:536; B:power_query_extractor/core/report_processor.py:487), which may attach to an Excel that is already running instead of starting a new one (READ; Excel's own behaviour cannot be run here, MEDIUM; B-G63). | Same, so no judgement. Shared defects: a sheet name that differs only in letter case gives no values (A-G17 / B-G52); the log says "read-only" while the file is opened read-write; there is a second Excel session per report. |
| 24 | A-F24 Consolidated workbook | B-F23 Consolidated workbook | different | The same layout and the same 9-column mismatch [GST RULE] ⚠ (the same code: A's copy differs only by a leftover header block whose two cells are overwritten at once and by two `except:` lines that B narrowed, READ `diff` of `power_query_extractor/core/data_consolidator.py`; same output from the same results, RAN MAP probe). **A** writes into the `Annual Statement-…` folder chosen by the LAST Step 4 scan, labels each row with the client folder name, and asks "Open the consolidated report?". **B** writes into the run folder that was shown when Start Extraction was pressed, labels rows "<client folder> / <version folder>", and asks nothing. | **B**, on the technical part (where the file goes and what survives an error). The column look-up [GST RULE] is identical (A-R88 / B-R83), so there is no judgement on it. B's destination is fixed at Start: the worker keeps it (READ B:gui/controllers/extraction_controller.py:29, 92). RAN B audit: the workbook went to the batch being processed even after the tree switched to a new batch. A reads the folder at the end of the run (A:power_query_extractor/gui/extractor_window.py:945). A re-scan during a run sent A's workbook into another target's batch (RAN A audit). On a save error, both lose the workbook (A-G29 / B-G43); when the error comes while the file is already being written (for example a date with a time zone), both also leave a half-written workbook that cannot be opened under the normal name (Section 9, M4, RAN both; audit B's B-R86 and B-G43 as corrected in its self-review). [Corrected during review - the half-written file was added from audit B's final text and Section 9.] B still writes its log file and run history; A shows an error line and an error box and keeps no record of its own on disk (the figures survive only in the program log, A-G56). A's "Open?" question is a convenience on Windows. |
| 25 | A-F25 Step 4 log, statistics, history | B-F24, B-F25 Step 4 progress, Stop, log file, history | different | **A:** a log sub-tab only; no Stop button, no log file and no history. **B:** a Step 4 log with Stop and Clear Log, a text file `Extraction_Log_<YYYYMMDD_HHMM>.txt` in the run folder, and extraction records in the run history. | **B.** B's run can be stopped (between versions) and leaves a record on disk. A's Step 4 run can only be ended by closing the program (A-G50 / A-G27; READ: no Stop in A:power_query_extractor/gui/extractor_window.py). B's caveats: the log file is written only at the end, so closing the window loses it (RAN B audit). It leaves out warnings that go only to the program log. The history stores every figure for ever under a key no screen reads (B-G25, B-G64). Shared: every refresh line is coloured "info", errors included (A-G65 / B-G71). |
| 26 | A-F26 Close Excel error dialogs | B-F26 Close Excel error dialogs | same | The same desktop-wide search for error dialogs and the same button clicking. | Same, so no judgement. Shared defect: another program's dialog can be answered "Yes" (A-G55 / B-G61, RAN both with fake windows). |
| 27 | A-F27 Settings | B-F27, B-F2 Settings and import | different | **A:** one file `~/.gst_organizer_cache.json`. It holds the four paths, two modes, the include-name tick, dark mode and the length, plus a recent-folders list nobody reads. **B:** `~/.gst_organizer/settings.json` (11 keys, including the three Step 4 options) plus `session_state.json`. B imports `~/.gst_organizer_cache.json` once, which is exactly the file A writes today. | **B,** narrowly. B saves on closing and keeps the Step 4 wait, suffix and Skip tick. A forgets those at each restart (A-R119), and does not save typed paths, the mode radio or on closing (RAN A audit t27). Shared: valid but wrongly typed content stops the start (RAN both: A-G24 / B-G36); a damaged file silently resets to defaults; files are rewritten in place; the last of two open copies wins. Neither file is in `%LOCALAPPDATA%\<App>\` (Windows' per-user folder for program data; Standard 13, fact only). |
| 28 | A-F28 Dark mode | B-F28 Dark mode | different | **A:** a tick in the title strip recolours every Tk widget and restores the recorded colours afterwards. **B:** View > Dark Mode applies one style sheet (QSS, Qt's CSS-like colour sheet). In both, parts stay light or become hard to read. | **Neither works fully.** A covers more widgets: it was rebuilt in `aeb83cf` with exact restore and has 11 tests. But a scan in dark mode gives white text on pale rows on Step 2 (RAN A audit), and Step 4 log lines turn black on dark grey. B leaves Steps 1-2 light, and its Step 3 and Step 4 status lines become dark-on-near-black (RAN B audit). Its disabled Stop looks enabled. B's code is far simpler (one sheet). |
| 29 | A-F29 Standalone extractor launcher | B-F30, B-F31 Launcher; old entry point | different | **A:** `launch_extractor.py` opens the extractor alone in its own window, takes the target from the settings file and scans it. **B:** `launch_extractor.py` opens the full four-tab window on Step 4. B's old `extractor_main.py` cannot even be imported. | **Mixed.** B's way keeps one window, so there is no second screen to maintain. But started that way B writes no log file at all and only prints its errors (RAN B audit). A's standalone window writes `pq_extractor.log` and works on its own (RAN A audit). It ignores the saved dark mode, and closing it during a run kills the run without a question (A-G27). Running `extractor_main.py` directly fails in both (A-G32 / B-G45). |
| 30 | A-F30 Application log | B-F32 Application log | different | The same main log set-up: `gst_organizer.log` in the start folder, level INFO. A's standalone extractor writes `pq_extractor.log`; B's launcher writes no log file at all. | **A,** narrowly. A logs under both of its start routes. Shared: the log lands in whatever folder the program was started from; on a cp1252 PC (an older Windows character set without ✓ or ❌) lines with ✓/❌ are lost (RAN both audits, simulated); client figures are logged at INFO (A-G56 / B-G64). |
| 31 | A-F31 Help, tooltips, scrolling | B-F33 Help, cards, scrolling | different | **A:** hover tooltips on Step 1; help texts; an "EXPECTED FILE NAMES" card read from the description texts stored next to the six patterns; wheel-scrolling fixes; a collapsible box on Step 2. **B:** no tooltips (0 of 205 widgets); fold-away cards with a hand-typed pattern list and mode lines that contradict the mode buttons; Qt scroll areas. At the opening window size, B's SCAN FILES button is below the visible part of Step 1. | **GST rule - no judgement** on the content of the file-name card [GST RULE] ⚠ (A-R16 / B-R13 / B-R141, MAP 2.1 item 3). Neither card mentions `.xls`, the hyphen ban or the "(N)" forms. **Technical part only: A.** A has tooltips, and its card follows the pattern table (READ A:gui/tabs/setup_tab.py:477-503; B:gui/views/setup_view.py:488-512 is typed by hand). B's SCAN FILES is out of view (RAN B audit). B's help contradicts its own buttons (Re-Run/Resume). Shared: help texts promise behaviour the code lacks (A-G66 / B-G72). B's Qt scrolling needs none of A's hand-written fixes (A-G68 is only A). |
| 32 | — | B-F2 One-time import of old settings | only B | The first time B's window opens on a PC where the old Tk program was used, it copies the old settings (paths, mode, include-name tick, length, dark mode) into its own settings file, once. | Only B, so no contest. It reads exactly the file A writes today; A's extra key `workflow_mode` is ignored (READ B:models/cache_manager.py:93-131). Defects, RAN B audit: the old recent-folders list is lost (key `recent_folders` against `recent_source_folders`); a saved length of 0 becomes 15; a damaged old file is retried at every start. |
| 33 | — | B-F17 Run history and "Last Run" | only B | B keeps a history of processing runs in the user's profile, and a "Last Run" column on Step 2. A keeps no history (A-R123). | Only B, so no contest. Defects, RAN B audit: "Last Run" appears only after the next scan; client records are never trimmed (55 kept after 55 runs); a wrongly typed history file makes B abort at the end of the next run (exit code 134); extraction records sit under a key "Last Run" never reads. |
| 34 | A-F7, A-F17, A-F21 (how the work is run) | B-F34 Background jobs and guards | different | **B** runs scanning, processing and extraction each in a background thread, and silently refuses a second job of the same kind. **A** scans on the window's own thread, so the window freezes on a slow share. It runs Step 3 and Step 4 in daemon threads, which are killed instantly when the program exits. A re-scan re-enables Step 4's Start during a run. | **B.** The window stays responsive, and workers talk to it only through signals (RAN B audit thread map). There is one job of each kind at a time (B:gui/controllers/*_controller.py guards). A's re-scan during Step 3 empties the data the running job uses and can stop the whole run (A-G13, only A). A's Step 4 can be started twice (RAN A audit). A's guard on Step 3 Start came only in `33d0baf`. B's caveats: different kinds of jobs run at the same time (a scan and processing completed during an extraction, RAN B audit); the refusal is silent; closing is accepted at once; a late clean-up can abort the program (LOW, B-G41). |

### 3.2 Explanations for the "different" features

Each subsection gives what one table line cannot hold, and where to read more. GST-rule parts are described, never judged.

#### Row 1 - Start-up (A-F1 / B-F1)

**A's start-up.** A was made to start on any PC (`c463256`, 26-07-2026), then made to refuse a Windows start without pywin32 (`be6418c`). The reason given is that reports built without Excel lose Power Query. Started from a console, A prints a report and, if something is missing, offers to pip-install (pip = Python's own package installer). The offer installs the whole `requirements.txt` (pytest too) and restarts (`a5ad487`). Off Windows the offer comes back at every start.

**B's start-up.** B's window imports the Step 4 code, which imports pywin32 with no guard. So the window module itself cannot load without pywin32 (RAN (C3)). The advice printed, `pip install PyQt5 openpyxl`, names two packages that are already installed. B became this way in `c258047` (07-02-2026), when it merged the extractor into the window. At the split, the organiser window did not import `report_processor` (B audit §3 B-F1, Origin).

**Shared and Standard 13.** For the author's Standard 13 (`rebuild_pipeline_FINAL.md:900-1061`), facts only:

- Neither project has a `.bat`/`pythonw` start.
- Neither has `sys.excepthook` or `threading.excepthook` (hooks that catch errors nothing else caught).
- A's install offer changes the user's Python.
- Started without a console, A dies at A:main.py:17 with no trace, and B shows nothing on screen when it fails.

Read more: A audit §3 A-F1 (l. 1023); B audit §3 B-F1; MAP rows A-R124 to A-R128; gaps A-G22, A-G53, A-G59 / B-G32, B-G34.

#### Row 2 - The main window (A-F2 / B-F1, B-F29)

**Closing.** Both projects lose work when the window is closed mid-run (A-G27 / B-G40):

- A asks once, and only for Step 3. Its worker is a daemon thread and dies about 1 s after "Yes". A asks nothing during Step 4.
- B never asks. In the B audit, closing during an extraction ended the program with exit code 0. It left an unrefreshed copy under the "Refreshed" name and wrote no workbook, log or history.

**B's tab dots.** They are new in `b0e7d0e` (09-02-2026). They turn green when a step has merely *run*: an empty scan, a failed extraction and a stopped extraction all turned them green. They are never reset, and Step 2's dot never turns green at all. The B audit flags these four points and, in its final text, lists them under gap B-G71 (feedback gaps). [Corrected during review - originally "but Section 7 has no gap number of its own for them"; audit B's self-review (its corrections 11-19) placed them in B-G71 after this section was written.]

Read more: A audit §3 A-F2 (l. 1141); B audit §3 B-F1 and B-F29; MAP rows A-R131, A-R132, A-R155.

#### Row 3 - Choosing the four paths (A-F3 / B-F3)

**What each picker offers.** The template types offered differ, and the code shows three answers in each project:

- A's picker offers `.xltx/.xlsx`; its checklist accepts `.xlsx/.xltx/.xltm/.xlsm`; the unused constant lists `.xlsx/.xltx/.xltm`.
- B's picker offers `.xlsx/.xlsm`; its caption says ".xlsx or .xlsm"; the unused constant lists `.xlsx/.xltx/.xltm`.

**When the settings are saved.** In A, a Browse saves at once. A typed path is not saved until some other click saves. If the length box holds text, every save, and every Browse, fails silently (A-G20, A-G26). In B, the settings are saved at a scan and at the other moments listed in the table, but not when START is pressed.

**B's invisible buttons.** They came with the restyle `7e09e5c` (08-02-2026). The cause is the white background rule on the content area of each coloured section (B-G67).

Read more: A audit §3 A-F3 (l. 1207); B audit §3 B-F3; MAP rows A-R52, A-R114 to A-R116; gaps A-G20 / B-G30, B-G50, B-G67.

#### Row 6 - Client-folder length (A-F6 / B-F4)

**The rule.** The shortening rule itself is identical (`create_client_state_key`, IDENTICAL by function comparison). It shortens only `<client>-<state code>` (Level 2), never the bracketed category names or the file names (A-G21 / B-G24, shared).

**A's box.** Until `322717f` (25-07-2026), A's box had a Tk input check that allowed only an empty text or a whole number from 10 to 100. That commit dropped it. Since then any text can be typed. An empty box broke Start even before (A audit §3 A-F6, History).

**B's box.** B's box cannot hold bad text. But its value is used even when the box is greyed out, which the screen suggests should mean "not in use".

Read more: A audit §3 A-F6 (l. 1390); B audit §3 B-F4 and B-F2 (point 2); gaps A-G21, A-G26 / B-G24.

#### Row 7 - Scanning and recognising files (A-F7 / B-F6)

**The GST-rule part [GST RULE] ⚠.** This is only the IMS Reco trailing "(N)". It is described in MAP 2.1 item 1, with a worked example and its effect on duplicates and on Links cell B11:

- A accepts the "(N)" since `5c8ce95` (23-07-2026). The message says such names "were not detected" and that every other type already tolerated the suffix.
- B has the original pattern from `edac6a3`.
- A's tests pin A's form: 4 of them fail against B's code (RAN (C3)).

No version is recommended here.

**The technical part.** The trade-off is speed against visibility:

- A's scan reads the folder once, reuses the listing's file details and reads the 8-byte signatures 16 at a time (`a603c02`, `56aaad9`, `4f84ba9`). A's commits report the measured effect on the office share. But the scan runs on the window's thread: the window does not repaint until the scan ends.
- B runs three wildcard searches and then, one file after another, a size query and an 8-byte read per file (B:utils/helpers.py:255-257; B:core/file_parser.py:171-207). That is about two round trips per file, never overlapped, against about 1.2 in A after `56aaad9`, with the reads overlapped 16 at a time after `4f84ba9`. So B is likely slower per file on a share (MEDIUM; READ, not timed here). B had already cut its own per-file requests in `7e09e5c` (08-02-2026): at the split each file was checked with several separate requests, as in A before `a603c02`. The same commit made B list small and non-Excel files as variations and skip `~$` names; its message calls this only "handle Excel file selection and validation more efficiently" (READ `git show 7e09e5c`; B audit §3 B-F6, Origin). But B's scan happens in a background thread.

**Where the two differ in what the user is told:**

- A drops small or non-Excel files and ordinary lock files without trace (A-G37). A `~$` lock file of 1 KB or more with a ZIP signature is listed and becomes a variation (MAP 4.3 item 4).
- B lists small and non-Excel files as variations, but shows only a count (B-G22). It skips `~$` names.
- B's count lines do not add up for the user: 28 files listed in the log, "Files found: 27", "Variations: 3" (RAN B audit, B-G22).

Read more: A audit §3 A-F7 (l. 1457); B audit §3 B-F6, B-F7; MAP rows A-R1 to A-R10, A-R109, A-R110; gaps A-G37, A-G48 / B-G22, B-G46, B-G55.

#### Row 8 - The client list on Step 2 (A-F8 / B-F8, B-F7)

**The GST-rule part [GST RULE]: described only.** Both projects build the client key the same way (`<client as written>-<state code>`) and decide the status the same way ("Complete" / "Missing n files" / "Has duplicates"; `_analyze_client_completeness` IDENTICAL). The display differs: A tints any non-complete row the same pale orange; B colours "Has duplicates" red-orange and "Missing" orange.

**The technical part.** The difference is in what happens between the table and the run:

- A rebuilds each ticked client's key from the displayed text. Tk turns "0042" into 42, so the key `42-MH` does not exist, and Start stops with "Fatal error" (RAN A audit; worked example MAP 2.1 item 4).
- B stores the key inside the row (B:gui/views/validation_view.py:371).
- B ticks every row after every scan (B:gui/views/validation_view.py:370). Pressing START at once processes complete and incomplete clients alike.
- B's sortable headers sort numbers as text, and processing follows the screen order (RAN B audit).
- B's "Folder Name" column shows the scan key, not the real (abbreviated, shortened) folder name (B-G73).

Read more: A audit §3 A-F8 (l. 1571); B audit §3 B-F8, B-F7; MAP rows A-R17, A-R29, A-R138; gaps A-G25 / B-G73.

#### Row 9 - The variations list (A-F9 / B-F9)

The hint texts are a [GST RULE] (they name return file patterns ⚠) and are computed identically. What differs is whether the list reaches the user at all:

- A writes it into the summary workbook's "Variations" sheet. It is visible only after processing, and it omits the files A dropped silently (A-G67).
- B never shows it. B's summary sheet then says "No file variations found", which is false whenever the scan found problem files.

B's own commit `c258047` replaced the list with `[]`. At the split, the Tk code passed the real list.

Read more: A audit §3 A-F9 (l. 1646); B audit §3 B-F9; MAP 2.1 item 2; gaps A-G67 / B-G20.

#### Row 10 - Per-client folder-name option (A-F10 / B-F4, B-F5)

In A, the option works per client since `33d0baf` fixed the key. Its global/per-client interplay still disagrees with the screen, in three ways (all RAN A audit, A-G19):

- turning the global tick off does not turn the per-client flags off;
- after a scan the table shows "☐ No" while the global tick adds names;
- the summary reports only the global tick.

A's long-name warning uses two different limits. Switching a per-client toggle on warns above a fixed 10 characters (A:gui/handlers/client_handler.py:191). Pressing Start warns above the configured length, but only when the global tick is on (A:gui/handlers/processing_handler.py:90-112). READ.

B lost the control in `c258047` / `a62d3cc`. Its organiser still has the branch, and that branch works when filled by hand (RAN B audit), but nothing on screen fills it. Whether a per-client option is wanted at all is a decision for the author (technical, not a GST rule: A-R40 / B-R35 are not tagged).

Read more: A audit §3 A-F10 (l. 1709); B audit §3 B-F5; MAP rows A-R40, A-R136, A-R137.

#### Row 11 - Dry run (A-F11 / B-F11)

The gate is a [GST RULE] by the registries' tagging (it decides which inputs a run needs). Both sides are described in MAP 2.1 item 7, with a worked example:

- A: no templates chosen, one client ticked → "Please select both ITC and Sales template files", and no preview.
- B: the same situation → a preview promising both reports.

Technically, the two previews share every other weakness (RAN both audits):

- they never show which folders, which mode effect, or which of several files feeds a report;
- after a run they show cleaned names instead of the source names;
- they do not say that a report for a client without inputs will be an unfilled template copy.

Read more: A audit §3 A-F11 (l. 1783); B audit §3 B-F11; MAP rows A-R51, A-R139.

#### Row 13 - Organising files into folders (A-F13 / B-F12)

**The GST-rule part [GST RULE] ⚠: folder names.** The folder names are both a layout choice and the text written into Links cells B2, B7, B10 (ITC) and B2, B8 (Sales). So they must match what the templates' queries expect. Which names the templates need cannot be checked here: the templates are not in either repository. How to check: open a real template's queries and see whether they read the folder from the Links cell or contain a fixed name (MAP 2.1 item 8).

The history is described only:

- B keeps the names of `edac6a3` (31-08-2025).
- A shortened them in `7347383` (26-07-2026) for Excel's path limit.

Two details of A's change, both READ (`git show -s 7347383`; A:utils/constants.py:85-90):

- The commit message gives two sets of figures. Measured on the real run: worst path 270 characters, 1094 of 3030 files over ~218 ⚠. Modelled for a planned shorter network-share root: 235 characters and 27 files over with the old names, 215 and none over with the short names. A's code comment repeats only the modelled figures and calls them "Measured", without the shorter-root condition. The A audit records this under A-R41 and A-G41. Which figures fit the office share is for the author to confirm.
- The message itself calls the margin "thin (215 vs 218)" and says "a path-length guard remains the durable fix". Neither project has such a guard (A-G41 / B-G54).

**The technical part: name cleaning.** The only technical difference in organising is name cleaning, which affects the copied file names and the bracketed names:

- A uses whole words in any letter case (`33d0baf`), and abbreviates underscore-joined names again (`4219ffb`).
- B replaces "Private"/"Limited" case-sensitively, anywhere in the text.

The effect in B (RAN (C3), `sanitize_out.txt`): a client written in capitals gets a client folder `ClientA Pvt Ltd-DL` but copy names `ClientA_PRIVATE_LIMITED…`, so B's own names disagree with each other. For underscore-joined names A's own names disagree in the same way, and so do B's: `ClientA_Private_Limited` gets the client folder `ClientA_Private_Limited-DL` but copy names `ClientA_Pvt_Ltd…` in both projects (same run, re-run in VERIFY-C3), because the client-folder rule treats "_" as part of a word. A's `33d0baf` message says the new cleaning matches the client-folder rule; after `4219ffb` that holds only for names written with spaces. The same cleaning also builds the client part of the report names (`ITC_Report_<cleaned client>_<state code>_…`, A:gui/handlers/processing_handler.py:292-297, B:gui/controllers/processing_controller.py:112-127), so those differ in the same cases: `ITC_Report_ClientA_Pvt_Ltd_…` in A, `ITC_Report_ClientA_PRIVATE_LIMITED_…` in B for a client written in capitals. The report-name pattern itself (A-R59 / B-R51 [GST RULE]) is identical; the cleaning (A-R19 / B-R19) is tagged GST by neither registry, and is described here, not judged as a GST rule. Everything else is the same code with the same results (MAP P:organizer) and the same defects (Resume, same-minute overwrite, folder collisions, no path-length check: A-G10, A-G9, A-G8, A-G41 / B-G13, B-G27, B-G15, B-G54).

Read more: A audit §3 A-F13 (l. 1917); B audit §3 B-F12; MAP 2.1 items 6 and 8; MAP rows A-R19, A-R39 to A-R48.

#### Row 14 - Creating the ITC report (A-F14 / B-F13)

**The GST-rule part [GST RULE] ⚠.** The cell map is identical in both: ITC B2, B4-B5, B7-B8, B10-B11; Sales B2-B3, B5-B6, B8-B9. So are the sources and the "first file of a type feeds the cell" rule. In both, the template's own text stays in the cells of missing file types (A-G5 / B-G2). One exception through row 7: in A, an IMS "(1)" duplicate sorts first and feeds B11 (MAP 2.1 item 1).

**The technical part: where the report is built.** A's staging copy means a failed Excel session normally leaves nothing under the report's name. In B, the unfilled template copy stays under the real report name, and B's Step 4 later lists, refreshes and extracts it as if it were a report (B-G5, only B).

**Shared serious defect (A-G54 / B-G63).** Both take over the user's running Excel through `GetActiveObject` (the call that attaches to an Excel that is already running), hide it and always call `Quit()`. Both code comments say "Don't quit if Excel was already open", and both do the opposite. Facts against Standard 12 ("Excel through COM", `rebuild_pipeline_FINAL.md:844-878`) are in both audits' Appendix X1.

Read more: A audit §3 A-F14 (l. 2034); B audit §3 B-F13; gaps A-G12, A-G28, A-G54 / B-G5, B-G38, B-G63.

#### Row 15 - Creating the Sales report (A-F15 / B-F14)

The re-read of the template boxes is tagged [GST RULE] in the map (MAP 2.1 item 11), because it decides which template a client's report is built from. Described only:

- A re-reads the Step 1 template boxes and the per-client flags for every client. A template changed mid-run applies to the later clients (RAN A audit, `v_sales_midrun.py`).
- B takes a snapshot of the settings at Start (B:gui/pyqt_main_window.py:352-371).

The "[DEBUG]" lines are purely technical. They show full output paths with client names in the user's log, and they have been present since `edac6a3` (14 at every commit up to `bf2f1dd`, A audit §3 A-F15, History).

Read more: A audit §3 A-F15 (l. 2167); B audit §3 B-F14; gaps A-G14, A-G69.

#### Row 17 - Progress and the Step 3 log (A-F17 / B-F16)

**A's log.** A's `0890022` fixed a late-binding fault (log lines could show the next client's name). But A's worker still makes 45 direct log-line calls and 2 progress-bar calls into the window, and shows the message boxes and makes the button changes from the worker thread (READ, count in A:gui/handlers/processing_handler.py:182-530; A-G70).

**B's log.** B's design avoids that entirely, but its log has costs of its own:

- B's HTML log drops anything between "<" and ">". In a trace this cut an error message to "❌ ITC Report Error: '" (RAN B audit).
- It drops leading spaces.
- It grows into one text block that slows the window (B-G56, B-G65).

**Shared.** Both show a stopped run as "Processing Complete / Successfully processed 1/5" (A-G69 / B-G70). Both take Stop only between clients.

Read more: A audit §3 A-F17 (l. 2371); B audit §3 B-F16; MAP rows A-R135, A-R140, A-R141.

#### Row 18 - Step 4 inside the main window (A-F18 / B-F18, B-F24)

**A's Step 4.** It is the standalone extractor panel placed inside a tab (`bfcb24a`, `c463256`, 26-07-2026). Its strength: the folder box follows Step 1's target as typed. Its weaknesses (all RAN A audit):

- the panel's status line is not shown inside the main window;
- the "full" hand-over overwrites the box and stops the following;
- the loading notice can no longer appear for pywin32 (the extractor guards pywin32 itself) and blames pywin32 for any other failure.

**B's Step 4.** It is a native page (`b0e7d0e`). It reads `settings.target_folder` (the last saved value), so a target typed but not yet saved is ignored. It also rebuilds its list on the window's thread at four moments: at start-up, after processing, after an extraction that ends normally or is stopped (not after one that ends in an error, `B:gui/pyqt_main_window.py:549-565`) and on Refresh List. [Corrected during review - originally "after extraction"; Section 4, 4.7 item 2.] Every rebuild loses the ticks (B-G28, B-G58, B-G73).

Read more: A audit §3 A-F18 (l. 2455); B audit §3 B-F18, B-F24; MAP row A-R144; gaps A-G65 / B-G28, B-G58.

#### Row 19 - Which batch and version Step 4 shows (A-F19 / B-F18)

This is a technical rule (no registry tags A-R42 / B-R39 as GST). The three rules side by side:

| | A | B |
|---|---|---|
| Batch (run folder) | newest date modified (since `edac6a3`) | highest name |
| Version folder | newest date modified (since `c1583d6`; before it, by name) | all versions holding a report, ordered by name (`b0e7d0e`) |
| Re-Run / Resume in Step 3 | newest date modified | newest date modified (the same code as A) |

So B's Step 3 and Step 4 can work on different batches (RAN B audit).

RAN (C3) (`runfolder_rules.py`, both expressions copied verbatim):

- Dates as made: A picks the 05-01-2026 batch, B the 31-12-2025 batch.
- After a lock file is written into the older batch: A switches to 31-12-2025 as well.

The A audit RAN the same switch with a real file (A-G45). Pointing A's box at a batch folder itself finds nothing and silently keeps the old list (A-G45). B shows only one batch at a time, so clients of earlier Fresh runs disappear from Step 4 (B-G12).

Read more: A audit §3 A-F19 (l. 2528); B audit §3 B-F18; MAP row A-R42; gaps A-G45 / B-G12, B-G13.

#### Row 20 - Step 4 selection and "last refresh" (A-F20 / B-F18)

**The GST-rule part [GST RULE] (MAP 2.1 item 10): described only.** Both look for `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`:

- A ticks by a lower-case prefix test and then finds the original with the search.
- B uses the search for both.

The two give different results only for a hand-renamed file whose letter case differs from what the program writes (for example `.XLSX`, or `itc_report_…`), on a case-sensitive disk; on Windows both ignore letter case (HIGH). READ A:power_query_extractor/gui/extractor_window.py:20-42; B:gui/pyqt_main_window.py:461-462.

**The technical part.** The status texts differ, and both have defects:

- A's status follows the configured suffix, but over-matches some suffixes (A-G16).
- B's status is per version (not per report) and recognises only "_Refreshed".
- Both trust the file name for "refreshed" and the file's date for "when", so a failed refresh reads as a refresh (A-G2 / B-G7).
- A pre-ticks every existing report. B starts all unticked and lists each version as its own job (A-R65 / B-R61, B-R136).

Read more: A audit §3 A-F20 (l. 2596); B audit §3 B-F18; MAP rows A-R59, A-R65, A-R68, A-R145.

#### Rows 12, 16 and 21 - one line is enough

These three rows are *different*, but the table line holds the whole difference. Pointers only:

- **Row 12, export (A-F12 / B-F10).** A audit §3 A-F12 (l. 1858); B audit §3 B-F10; gaps A-G31 / B-G73. B's button is drawn and styled (B:gui/views/validation_view.py:186-221) but `btn_export` has no `.clicked.connect` anywhere in B (READ, grep).
- **Row 16, summary workbook (A-F16 / B-F15).** A audit §3 A-F16 (l. 2271); B audit §3 B-F15; gaps A-G19 / B-G19, A-G67 / B-G20. Precisely: in both projects the summary routine catches its own errors, writes them to the program log file and returns "False" (B:core/excel_handler.py:313-346; A's routine is IDENTICAL, MAP row A-F16). A's Step 3 log shows that outcome as "❌ Failed to create summary report" (A:gui/handlers/processing_handler.py:469-473). B's Step 3 log shows only the success line; its error line would need the routine to raise, which it does not (B:gui/controllers/processing_controller.py:189-193). READ.
- **Row 21, refresh (A-F21 / B-F19, B-F20).** A audit §3 A-F21 (l. 2655); B audit §3 B-F19 and B-F20 (traces of all 15 Excel steps, 39 s of fixed waits per report at the default wait, and the BB2 values that pass); gaps A-G3 / B-G6, A-G55 / B-G61 and B-G62, A-G2 / B-G7. The refresh function differs only by an unreachable `return` and the kinds of error caught (MAP row A-F21). The BB2 check (`_validate_refresh`) is IDENTICAL [GST RULE] ⚠.

#### Row 22 - Skip refresh and naming (A-F22 / B-F21)

**A's recognition.** A replaced the hard-coded `_Refreshed_` search with `refresh_naming.py` in `2aaa2b1` (25-07-2026). Its message: "Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one". It comes with 19 tests in `tests/test_refresh_naming.py` and 2 in `tests/test_extractor_panel.py`.

**B's recognition.** B keeps the `bd176bf` state (30-10-2025, shared): the suffix box is configurable, but the Skip search stays fixed. B's `b0e7d0e` added a second fixed text for the tree status.

**RAN (C3), `suffix_both.py`.**

| Suffix | A | B |
|---|---|---|
| default | (no copy present yet) finds nothing | finds nothing |
| `_PQ_{timestamp}` | finds the copy | does not find the copy |
| `_R{timestamp}` | also takes the original report as a "refreshed copy", because "_r" occurs inside "itc_report_…" | finds nothing |

**Shared defects.** None of these records whether a copy was ever refreshed successfully (A-G2 / B-G7). A suffix without `{timestamp}` overwrites the last copy, and an empty suffix fails every report (A-G16 / B-G9).

Read more: A audit §3 A-F22 (l. 2759); B audit §3 B-F21; MAP rows A-R67 to A-R69, A-R119.

#### Row 24 - The consolidated workbook (A-F24 / B-F23)

**The GST-rule part [GST RULE] ⚠: identical in both.** The 9 columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" (IGST/CGST/SGST) are filled from `Diff GSTR-3B!J15/K15/L15`, not from their configured cells (A-G1 / B-G1, RAN in both audits and in the map's probe). A failed refresh turns every figure of that report into "Missing" (A-G17 / B-G10). Which cell is meant is the author's decision.

**What differs is technical:**

- the destination folder (B fixed at Start; A whatever the last scan chose);
- the row label (A the client folder, one row per client; B "client / version", one row per ticked version);
- what is left after a save error (A: an error line in the Step 4 log and an error box, no record of its own on disk - the figures survive only in the program log, A-G56; B: one log line, plus the log file and the run history, which hold the figures);
- A's closing question "Open the consolidated report?" (`os.startfile`, Windows only).

Read more: A audit §3 A-F24 (l. 2899); B audit §3 B-F23; MAP rows A-R86 to A-R91; gaps A-G29 / B-G43.

#### Row 25 - Step 4 log, Stop and history (A-F25 / B-F24, B-F25)

**What only B has.** B added the Stop flag with `c258047`, and the log file with `b0e7d0e`. Stop takes effect after the current version: about 39 s per report at the default wait (RAN B audit). After Stop, B's run then ends like a normal one, with a green dot and "Extraction complete" (B-G70).

**Confidentiality.** B's log file and run history copy every extracted figure into plain-text files outside the reports, and the history keeps them for ever (B-G64). A writes the same figures only to its program log (`gst_organizer.log`, or `pq_extractor.log` when the extractor is started on its own) (A-G56). B writes them to `gst_organizer.log` as well when started through `main.py`: both `_log` methods log every line at INFO (READ A:power_query_extractor/core/report_processor.py:53-60; B:power_query_extractor/core/report_processor.py:30-37). Both are facts for the author's confidentiality decision.

Read more: A audit §3 A-F25 (l. 3004); B audit §3 B-F24, B-F25; MAP rows A-R123, A-R135, A-R143.

#### Row 27 - Settings (A-F27 / B-F27, B-F2)

**When each project saves.**

- A saves after a Browse, the workflow radio, the length arrows, the global include-name tick and the dark switch.
- B saves at a scan, after processing, on the dark switch, on File > Save Settings and on closing.
- Neither saves when a job starts.

**What can stop the start.** Both stop on wrong value types, in different files:

- A: a JSON list (JSON = the plain-text format of both settings files), or a length written as text (RAN A audit t27).
- B: `pq_wait_time` as text, or a JSON list (RAN B audit t27).

**The import between them.** B's import of A's file (row 32) means a merged tool could read A users' settings once. B's `session_state.json` is written, but its contents are never used.

Read more: A audit §3 A-F27 (l. 3157); B audit §3 B-F27, B-F2; MAP rows A-R114 to A-R119; gaps A-G24 / B-G36, B-G37, B-G82.

#### Row 28 - Dark mode (A-F28 / B-F28)

**A.** A walks every Tk widget and records its colours. It was rebuilt in `aeb83cf`, with later fixes `725fbaf`, `c89c1d4` and `169c7b5` (all 26-07-2026), plus 11 tests. Still, after one dark/light round the selected row and the selected tab keep changed colours (RAN A audit, A-G64).

**B.** B's single style sheet cannot reach widgets that carry their own colours. That is why Steps 1-2 stay white, and why the status lines become dark grey on near-black (RAN B audit pixel counts, B-G69).

Read more: A audit §3 A-F28 (l. 3229); B audit §3 B-F28; MAP row A-R148.

#### Rows 29 and 30 - Launcher and log file (A-F29, A-F30 / B-F30, B-F31, B-F32)

**A's launcher.** It opens the standalone panel window. It finds the target in the home settings file first: A fixed the old wrong file name in `c1583d6`; A's own older document `PROJECT_AUDIT_ANNUALSTATEMENT.md` had reported it as a confirmed bug.

**B's launcher.** B rewrote it in `c258047` to open the main window on tab 4. It sets up no logging, so INFO lines are lost and no file is written (RAN B audit). B's `extractor_main.py` still exists, but imports a window file deleted in `a62d3cc`.

**The main log.** Its set-up is unchanged since `edac6a3` in both (same file name, start folder, INFO, no encoding).

Read more: A audit §3 A-F29 (l. 3289), A-F30 (l. 3351); B audit §3 B-F30, B-F31, B-F32; MAP rows A-R118, A-R120 to A-R122; gaps A-G32 / B-G44, B-G45.

#### Row 31 - Help, tooltips and scrolling (A-F31 / B-F33)

**The card content [GST RULE] ⚠: described only.**

- A's card shows each pattern's own description text, for example "IMS Reco - ImsReco-ClientName-State-DDMMYYYY.xlsx" (A:utils/constants.py:19-59).
- B's card shows hand-typed lines without the extension, for example "ImsReco-Client-State-DDMMYYYY".

Both name the same six types in the same order (MAP 2.1 item 3).

**A's card is not drift-proof.** A's code comment says the card "can never drift from the regexes". It reads the description strings, not the regular expressions (text patterns). So it, too, omits `.xls`, the "(N)" forms, letter-case freedom and the hyphen rule (A audit §3 A-F31, finding 69). Map rows A-F31 and 4.1 (`322717f`) say "built from the patterns"; see correction 2 in 3.5.

**Technical.** A has 33 tooltips on Step 1 and B none. B's mode lines say "Re-Run: Overwrite existing" while its radio button says "Add new version folders". B's note promises warnings that do not exist (B-G72). A's hand-written wheel handling has gaps (`<MouseWheel>` only; small steps dropped in the Step 4 log, A-G68), but is guarded by 64 source-text tests.

Read more: A audit §3 A-F31 (l. 3425); B audit §3 B-F33; MAP rows A-R16, A-R146, A-R147.

#### Row 34 - How the work is run (A-F7, A-F17, A-F21 / B-F34)

**B's model.** B's controllers use QThread workers, Qt's background-thread objects (`c258047`). Each controller refuses a second job of its own kind, but only in the program log. Nothing ties the three together. In the B audit, a scan and a whole processing run finished while an extraction was still running, and the end of processing switched the Step 4 list to the new batch mid-extraction (B-G74).

**A's model.** A has three different models:

- the scan on the window thread (A-G48);
- Step 3 in a daemon thread that touches the window directly (A-G70);
- Step 4 in a daemon thread whose screen updates were moved onto the window thread in `c1f425e` (25-07-2026).

A's re-scan during Step 3 empties the dictionaries (the in-memory tables of files and clients) the worker is reading (A-G13, only A). B builds a new parser and result for every scan, so a running job keeps its own data.

Read more: A audit §3 A-F7, A-F17, A-F21; B audit §3 B-F34; gaps A-G13, A-G27, A-G48, A-G65, A-G70, A-G71 / B-G40, B-G41, B-G71, B-G74, B-G75, B-G76.

### 3.3 The features that are the same, only in A or only in B

- **Same (rows 23, 26).** Cell extraction and the error-dialog closing are the same code in both, so they carry the same defects. The first carries a [GST RULE] ⚠ mapping table with no tests in either project (A tests notes §12). Keeping either version keeps these behaviours unchanged.
- **Only A (rows 4, 5).** The live Setup checklist and the workflow choice. Both were added in A's July 2026 work (`322717f`, `510f7b4`) and are tested (19 + 13, 7 tests). Their help texts and the "full" hand-over have the defects listed in rows 4-5.
- **Only B (rows 32, 33).** The one-time import of the old settings file, and the run history. Both came with `c258047` (07-02-2026). B has no tests, so all evidence for them is from traces (RAN B audit).
- **Lost in B, kept in A.** These are not separate rows. B's `a62d3cc` says "All functionality has been reimplemented", but three features did not survive the move to PyQt5: the client-list export (row 12), the per-client folder-name toggle (row 10), and the variations list in the summary workbook (row 9). This is a fact about B's history, useful for the decision list.

### 3.4 Checks run for this section

Everything below ran on copies in `/root/merge-scratch/run/C3/` (`A/` = A-copy, `B/` = B-copy, both copied fresh with `__pycache__` removed), with `python -B`. `HOME` and `TMPDIR` pointed into that folder.

| # | Check | How | Result |
|---|---|---|---|
| C3-1 | Can each window module load without pywin32? | `import gui.pyqt_main_window` (venv-B, offscreen Qt) and `import gui.main_window` (venv-A); output `/root/merge-scratch/run/C3/import_out.txt` (RAN) | B: `ModuleNotFoundError: No module named 'pythoncom'`. A: imported, `WIN32_AVAILABLE = False`, with the warning "pywin32 not available … Power Query refresh is disabled". |
| C3-2 | A's whole test suite | `xvfb-run -a … pytest -p no:cacheprovider -o addopts="" -q` (xvfb = a virtual screen for Tk); output `/root/merge-scratch/run/C3/a_suite_out.txt` (RAN) | `1 failed, 287 passed`. The failure is `test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only` (a Windows path read on Linux), the same as recorded in the A tests notes. |
| C3-3 | A's pure-logic tests against B's code | `test_file_patterns.py`, `test_helpers.py`, `test_scan.py` copied into a copy of B (`/root/merge-scratch/run/C3/AtestsOnB/`), venv-B (RAN) | `7 failed, 66 passed`: 4 IMS "(N)" tests and 3 name-cleaning tests. `test_file_ops.py` cannot be collected: "cannot import name 'scan_excel_files'". |
| C3-4 | The same made-up source folder scanned by both parsers | `/root/merge-scratch/run/C3/scripts/scan_both.py` → `scan_A_out.txt`, `scan_B_out.txt` (RAN) | A: 4 files listed (including `.XLSX` and IMS "(1)"), 0 variations; the 500-byte file, the lock file, the `.csv` and the sub-folder file are not mentioned. B: 3 files listed, `.XLSX` missed (Linux), 2 variations (IMS "(1)" "No matching pattern found"; `tiny.xlsx` "File too small to be valid Excel"). |
| C3-5 | Which earlier copy "Skip refresh" finds | `/root/merge-scratch/run/C3/scripts/suffix_both.py` → `suffix_out.txt`. A's own `refresh_naming`; B's search expression from B:power_query_extractor/core/report_processor.py:102 (RAN) | Default suffix: neither finds a copy (none present). `_PQ_{timestamp}`: A finds the `_PQ_` copy, B nothing. `_R{timestamp}`: A finds the original AND the copy, B nothing. |
| C3-6 | Which batch Step 4 picks | `/root/merge-scratch/run/C3/scripts/runfolder_rules.py` → `runfolder_out.txt`; both selection expressions copied verbatim (RAN) | Dates as made: A `…050126 0900`, B `…311225 1000`. After a lock file in the older batch: both `…311225 1000`. |
| C3-7 | Name cleaning | `sanitize_filename` and `create_client_state_key` of each project on four made-up names → `/root/merge-scratch/run/C3/sanitize_out.txt` (RAN) | As in row 13. The Level-2 keys are identical in both; the copy names differ for capitals and for "Privateer". |
| C3-8 | Code re-read in both projects (READ) | the lines cited in the table | All confirmed, among them: B:…/report_processor.py:11-15 (unguarded imports) / A:…:20-36 (guard); A:gui/handlers/file_handler.py:31, 42 / B:gui/views/setup_view.py:612, 619; A:gui/tabs/setup_tab.py:440-443 / B:gui/views/setup_view.py:374-382 and B:gui/pyqt_main_window.py:203; A:gui/handlers/processing_handler.py:224 / B:gui/controllers/processing_controller.py:73; `btn_export` never connected (B grep); A:gui/handlers/processing_handler.py:38-41 (dry-run gate) / B:gui/pyqt_main_window.py:326-343; A:core/excel_handler.py:135-155, 212-215 / B:core/excel_handler.py:122-129; 14 "[DEBUG]" in A:gui/handlers/processing_handler.py, none in B's window code; A:…:469-475 / B:gui/controllers/processing_controller.py:189-193; A:power_query_extractor/gui/extractor_window.py:631, 641-648 / B:gui/pyqt_main_window.py:431-457; A:…extractor_window.py:945 / B:gui/controllers/extraction_controller.py:29, 92; wait boxes A:…extractor_window.py:273-274 / B:gui/views/extraction_view.py:127; no Stop in A's extractor (grep); A:gui/main_window.py:344-355 / B:gui/pyqt_main_window.py:601-608; A:gui/tabs/setup_tab.py:477-503 / B:gui/views/setup_view.py:488-512; B `setToolTip` 0 hits, A `Tooltip(` 9 call sites; A:gui/handlers/client_handler.py:152 / B:gui/views/validation_view.py:370, 446. |
| C3-9 | Commit messages (READ) | `git -C /root/merge-scratch/combined log -1` for `12d505f`, `a603c02`, `56aaad9`, `4f84ba9`, `c1583d6`, `c1f425e`, `2acffed`, `33d0baf`, `0890022`, `c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e` | Quoted where used. No message contains client data. |
| C3-10 | Second reading after the interrupted first run (READ) | the lines cited in rows 6, 10, 15, 16 and the "Rows 12, 16 and 21" note, re-read in both projects | All confirmed: B's length box disabled at start and enabled only by the include-name tick (B:gui/views/setup_view.py:374-382), its value passed on regardless (B:gui/controllers/processing_controller.py:51); A's two warning limits (A:gui/handlers/client_handler.py:191; A:gui/handlers/processing_handler.py:90-112); B's settings snapshot is a new object at Start (B:gui/pyqt_main_window.py:194-208, 363), later changed in place only for dark mode (:595); A re-reads the Sales template box per client (A:gui/handlers/processing_handler.py:332); the summary routine returns "False" on any error (B:core/excel_handler.py:313-346); `btn_export` has no `.clicked.connect` (B:gui/views/validation_view.py:184-297 connects five other buttons); 33 `Tooltip` attachments on A's Step 1 from 9 call sites (A:gui/tabs/setup_tab.py:314-467), 0 `setToolTip` in B; window sizes A:utils/constants.py:133-137 / B:gui/pyqt_main_window.py:71-73; A's category-folder comment A:utils/constants.py:85-90 against `git show -s 7347383`. |
| C3-11 | Independent re-run by the checker (VERIFY-C3) | Fresh copies of A-copy and B-copy in `/root/merge-scratch/run/VERIFY-C3/`, `python -B`, `pytest -p no:cacheprovider`; scripts in `scripts/` there (RAN) | C3-1: same result (B fails at `report_processor.py:11` through `gui/controllers/extraction_controller.py:10`). C3-2: `1 failed, 287 passed`, same failure; per-file counts confirmed by `--collect-only`. C3-3: `7 failed, 66 passed`, same seven tests; `test_file_ops.py` same import error. C3-4, C3-5, C3-6, C3-7: outputs identical to C3's. New: `quote_out.txt` (A's checklist marks a quoted folder ✕ "Folder not found" and keeps Scan unavailable; a file passes as the source folder); `b_status_out.txt` (B's tree status stays "Pending" with a `_PQ_` copy); `key_length_out.txt` (`create_client_state_key` with length 8 drops the state code: `ClientA ` from "ClientA Traders", Maharashtra; the same function gives the same result in B, whose box cannot go below 15). |

Nothing was written into A-copy, B-copy, the original checkouts, or anything under `/home/user`.

### 3.5 Corrections and clarifications found while writing this section

1. **Map row A-F13 / B-F12, "only the category folder names differ".** This is incomplete. Copied file names (and the bracketed names, and - found during the check of this section - the client part of the report names, which use the same cleaning) also differ when "Private"/"Limited" is written in capitals or sits inside a word: A uses whole words in any case, B a case-sensitive replace anywhere (RAN (C3) `sanitize_out.txt`; READ B:utils/helpers.py:78). The map records this in its rule row A-R19 / B-R19, but not in the feature row.
2. **Map row A-F31 and §4.1 (`322717f`), "a file-name card built from the real patterns".** More precisely: built from the `description` strings stored next to each pattern (A:gui/tabs/setup_tab.py:494-503 reads `info['description']`; texts at A:utils/constants.py:19, 27, 35, 43, 51, 59), not from the regular expressions. So the card can still drift; it already omits `.xls`, the "(N)" forms and letter-case freedom (A audit §3 A-F31, finding 69).
3. **Map row A-F24 / B-F23, "B writes into the run folder shown on Step 4".** More precisely: the run folder shown when Start Extraction was pressed. The worker keeps it (B:gui/controllers/extraction_controller.py:29, 92), so if the Step 4 list is rebuilt during the run, the workbook still goes to the batch being processed (RAN B audit, B-F34 Run 1).
4. **Map row A-F27, the list of A's saved settings.** A's settings file also holds `recent_folders` (A:gui/utils/cache_manager.py:32-43), which nothing reads (A-R117). Not wrong, only incomplete.
5. **A's commit `12d505f` message: "On failure the destination is no longer left with a half-written template copy".** This holds for a failed Excel session. A final copy broken by a network drop still leaves a partial file at the destination (A:core/excel_handler.py:215 raises → :226-228 returns False, no clean-up), as the A audit records under A-G12. This is not an audit error; it is a limit on the commit message's claim, and it matters when rows 14-15 are weighed.
6. **A's own code comment on the category folder names (A:utils/constants.py:85-90).** It calls 235 characters / 27 files over (old names) and 215 / none over (short names) "Measured on a real run". The commit `7347383` that wrote it says those are modelled figures that assume a planned shorter share root; the measured run gave 270 characters and 1094 files over (READ `git show -s 7347383`). This is not an audit error: the A audit already records it (A-R41, A-G41), and MAP 2.1 item 6 quotes the message correctly. It is listed here because anyone reading only the code would take the comment's figures as measured.
7. **Map gap row A-G20 / B-G30, B-G50, "Only B: quotes from 'Copy as path' are not removed".** A does not remove them either. A folder pasted with its quotes is marked ✕ "Folder not found" by A's checklist, which keeps Scan unavailable, and A's own scan check would refuse it with "Source folder does not exist" (RAN (VERIFY-C3), `quote_out.txt`; READ A:gui/utils/setup_validation.py:36-53, A:gui/handlers/file_handler.py:120-130). The A audit does not record this. The difference is only when the user is told: A while typing, B when SCAN is pressed. Found during the check of this section (VERIFY-C3).
8. **Map row A-F24 / B-F23, "code identical".** The output is the same, but the code is not quite identical. A's `power_query_extractor/core/data_consolidator.py:73-83` has a leftover header block whose two cells are overwritten at once (A-R92 records it), and A keeps two bare `except:` lines where B catches `(TypeError, ValueError)` (READ `diff` of the two files). No effect on the workbook. Found during the check of this section (VERIFY-C3).
9. **A's commit `33d0baf` message: sanitize "now uses word-boundary, case-insensitive matching ..., matching create_client_state_key".** After `4219ffb` (which turns underscores into spaces before abbreviating) this is true only for names written with spaces. For underscore-joined names the copies are abbreviated and the client folder is not (`ClientA_Private_Limited` → copies `ClientA_Pvt_Ltd…`, folder `ClientA_Private_Limited-DL`; RAN (C3) `sanitize_out.txt`, re-run in VERIFY-C3). B shows the same pair. Map row A-R19 states the two rules correctly. This is not an audit error; it limits the commit message's claim. Found during the check of this section (VERIFY-C3).
10. **B's scan behaviour came with `7e09e5c` (08-02-2026), not with the PyQt5 rewrite.** At the split (and still at `c258047`) B's scan dropped small and non-Excel files silently, as A still does. `7e09e5c` made B list them as variations, skip `~$` names and cut its per-file requests (READ `git show 47a1492:utils/helpers.py`, `git show 7e09e5c:core/file_parser.py`). The B audit records this (§3 B-F6, Origin). It is noted here because that commit's message describes it only as "handle Excel file selection and validation more efficiently". Found during the check of this section (VERIFY-C3).

## 4. Fixes only one side got

After the two projects split (after the shared commit `47a1492`, 30-10-2025 — a *commit* is one saved change in a project's history), each project went on correcting and improving its own copy. Project A made 34 commits of its own (`fc279af` … `bf2f1dd`: two in March 2026, 32 in July 2026). Project B made 5 (`c258047` … `b0e7d0e`, 07-02-2026 to 09-02-2026). This section takes every fix and every improvement in those 39 commits — one row each, because one commit often holds several — and checks **in the other project's code** whether the other project still has the problem, or still lacks the improvement. No answer in the "other project" column is assumed: each one cites the other project's code, and 35 of the 76 rows were also run (the others rest on reading the code of both projects).

Think of two article assistants who each corrected a different set of mistakes in their own copy of the same working papers after the file was divided. Before the merged file is signed off, you want a list of which corrections exist in only one copy. This section is that list.

### 4.1 How to read this section

**The two projects.** A = `goyal-sidhant/AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026): a Tkinter window (Python's built-in window toolkit) calling itself "GST File Organizer v3.0", with 288 tests. B = `goyal-sidhant/AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026): a PyQt5 window (a separate, add-on window toolkit) calling itself "v4.0", with no tests. The names mislead: A is not the older project and B is not the newer or better one. GitHub's records say B's repository was created on 31-08-2025 and A's on 07-02-2026.

**Columns of the tables in 4.2 and 4.3**

| Column | What it holds |
|---|---|
| Row | `4A-n` for a fix or improvement made in A, `4B-n` for one made in B. Other sections and the decision list can point to these numbers. |
| Commit, date | The commit's short code, its date and time (India time), and its own title in quotes. |
| What it fixed or improved | In plain English, with the place in the project's own code. |
| Does the other project still have it? | A verdict word in bold, then the evidence **from the other project's code**. |
| How checked | READ or RAN (below). |
| Consequence for the author | What it means for "Annual Statement 2.0". |

**Verdict words**

| Verdict | Meaning |
|---|---|
| **Yes** | The other project still has the problem, or still lacks the improvement. |
| **Partly** | Some of the problem is still there; the row says which part. |
| **No - fixed independently** | The other project fixed the same thing in its own way. |
| **No - never had it** | The problem came from code that only one project ever had — often that project's own earlier commit. Such a problem is a *regression*: something that worked and was broken by a later change. |
| **No such feature** | The other project does not have the part of the program the fix is about. |

**Evidence marks**

- `A:path:line` and `B:path:line` point into A at `bf2f1dd` and B at `b0e7d0e`; `path:line@commit` points into older code.
- **READ** = established by reading the code of both projects.
- **RAN (C4)** = run for this section, in `/root/merge-scratch/run/C4/`, with `python -B`, on made-up inputs only (invented clients `ClientA`, `ClientB`, `ClientC`; no client data). The scripts and their outputs are listed in 4.6. An earlier, interrupted attempt at this section had left some of the same scripts and outputs there; every one was re-run for this version, with the same results.
- **RAN (VERIFY-C4)** = run (or re-run) by the independent check of this section, in `/root/merge-scratch/run/VERIFY-C4/` (see the end of 4.6).
- **RAN (A audit)**, **RAN (B audit)**, **RAN (MAP)** = run by that earlier piece of work; the result was re-read against the code here.
- **HIGH / MEDIUM / LOW** = how confident an inference is that was not run.
- **COM** (Windows' way for one program to remote-control another; here Python drives Excel) cannot run on this Linux machine, and **pywin32** (the Windows-only Python package that provides COM) is not installed here. So everything Excel itself would do is READ only. Where B's window code had to be imported, empty stand-in modules took pywin32's place; the row says so.

**GST rules.** A *GST rule* is anything that decides a tax figure; which return, table, sheet or cell a figure belongs to (for example GSTR-9 Table 6J or 8D ⚠); a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. When in doubt, a point is treated as one. Such rows are tagged **[GST RULE]**. Their "Consequence" column says "GST rule - the author decides", and section 4.4 sets out what each project does, runs one worked example through both, and says when each version came in and what the commit messages hint about why. No version of a GST rule is recommended anywhere in this section. **⚠** marks values that change over time (file-name conventions, return and table names, template cell positions, folder names written into templates, state lists). "Which works better" remarks are made for technical behaviour only (crashes, safety of Excel and files, speed, tests, clarity of messages).

**Terms used below** (plain meanings; each is also explained where it first matters)

| Term | Plain meaning |
|---|---|
| Step 1 / 2 / 3 / 4 | The four tabs: Setup (choose folders and templates, scan), Validation (client list), Processing (sort files, create the ITC and Sales reports), Extract (refresh the reports' Power Query and collect the figures). |
| Run folder, Version folder, category folders | `Annual Statement-<DDMMYY HHMM>` (one per batch), `Version-<DDMMYY HHMM>` inside each client folder (one per processing run), and the three folders inside a Version folder that hold the GSTR-3B, ITC and Sales source files. |
| Links sheet | The sheet in each ITC and Sales template into which the program writes folder paths and file names for the template's queries to read. |
| Power Query, refreshed copy | Excel's data-import tool inside the templates; Step 4 copies each report to a "refreshed copy" and refreshes that copy. |
| openpyxl | A Python library that writes `.xlsx` files without Excel. It drops Power Query when it saves a workbook. |
| Variation | The code's word for a file the scan could not accept. |
| Network share, round trip | A folder on the office server; one request to the server and its answer (slow over a network or VPN). |
| Listing, stat | Reading a folder's list of entries in one request; asking for one file's size and dates. |
| Thread, worker thread | A second line of work running beside the window. A window can repaint only when its own thread is free. |
| Qt signal | Qt's way for one part of the program to send a message to another, safely across threads. |
| Test, pytest | A small program that checks the code automatically and fails when behaviour changes; Python's testing tool. |
| Console, `pythonw` | The black text window; the Windows Python that runs a program without one (the author's Standard 13 starts apps that way). |
| `requirements.txt` | The list of packages that `pip install -r requirements.txt` installs (pip = Python's own package installer). |
| strace | A Linux tool that lists every request a program makes to the operating system; used here to count file-system requests. |

### 4.2 Project A's fixes and improvements (34 commits)

#### 23-07-2026, night: network speed and the first correctness fixes

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-1 | `a603c02`, 23-07-2026 00:22, "Speed up network file scanning (behavior-preserving)" | The scan searched the source folder three times (once each for `.xlsx`, `.xls`, `.xlsm`) and asked the file system up to four separate questions per file. Now: one listing of the folder (`os.scandir`), fewer questions per file, and a repeated second validity check per file removed. Today's code: A:utils/helpers.py:309-366. Side effect: on Linux and macOS, names with capital extensions such as `.XLSX` are now found (Windows ignores letter case anyway). | **Partly.** B made its own, smaller saving in `7e09e5c` (one stat and one 8-byte read per file, B:core/file_parser.py:171-215). But B still searches the folder three times (B:utils/helpers.py:255-257). On Linux B misses a capital extension completely: `GSTR3B-ClientB-Delhi-May.XLSX` was scanned by A and appears nowhere in B's result, not even as a variation (B-G46). | READ both. RAN (C4) strace over 50 made-up files: A opened the folder listing once, B three times; per file both made 1 metadata request and 1 open (on Linux). RAN (C4) core probe for the `.XLSX` name. | Technical (speed): A's single listing is the faster design. On the office's Windows PCs the capital-extension difference does not arise (HIGH: Windows file names ignore letter case). |
| 4A-2 | `12d505f`, 23-07-2026 00:34, "Build Excel reports on local disk, copy to network once" | Step 3 used to copy the template straight to the report's final place on the share and let Excel edit it there. Now each report is built in a local temporary folder (`%TEMP%\gst_excel_temp`) and the finished file is copied to the share once: COM path A:core/excel_handler.py:147-155, 212-215; openpyxl path :273-279, 305-322. If building fails, nothing is left at the destination. | **Yes.** B copies the template to the report's final name first and fills it in place, on both paths (B:core/excel_handler.py:124 COM path, :228 openpyxl path). With a damaged template, B leaves an unfilled copy of the template under the report's normal name (B-G5). B's Step 4 offers every `ITC_Report_*.xlsx` it finds (B:gui/pyqt_main_window.py:461), so it would offer that unfilled copy for refreshing (READ). | READ both. RAN (C4) core probe, openpyxl path, damaged made-up template: A left nothing in the destination folder; B left `ITC_Report_ClientA_DL_211026_1030.xlsx` there. The COM path cannot run here (READ only). | Technical (file safety and speed): A's version is safer. Still open in A: an interrupted final copy leaves a partial file under the final name (A-G12). Unchanged in both: the report code takes over the user's own open Excel and quits it (A-G54 = B-G63). Open question from A's audit: does a report saved in `%TEMP%` record that folder inside the workbook? (check on Windows). |
| 4A-3 | `275d9a4`, 23-07-2026 00:46, "Cache level-1 folder; single-listing scans in PQ extractor" (part 1) | Step 3 now works out the run folder once per run and reuses it for every client (A:core/file_organizer.py:150-161). Before, in Re-run and Resume modes, the target folder was listed and every entry's date was asked again for every client; in Fresh mode the run folder was created again for every client. Side effect: each client's text report no longer counts the run folder once per client in its "Folders Created" line. | **Yes.** B:core/file_organizer.py:146-179 runs for every client (called at :90). | READ both. RAN (C4) core probe, three made-up clients, Fresh mode: "Folders Created" read 6 / 11 / 16 in A and 6 / 12 / 18 in B. | Technical (speed in Re-run and Resume on the share). In both projects the "Folders Created" figure is a running total for the whole run, not a per-client count (A-G19 = B-G19, B-G21). |
| 4A-4 | `275d9a4`, 23-07-2026 00:46 (part 2) | Step 4 checks the newest Version folder of each client with one listing for "has an ITC / Sales report" and one for "when last refreshed" (A:power_query_extractor/gui/extractor_window.py:20-42, 45-82), instead of several file searches. | **Yes, in B's own form.** B's Step 4 (written in `b0e7d0e`) runs three file searches in every Version folder of every client (B:gui/pyqt_main_window.py:461-462, 468), asks every refreshed file's date once to find the newest, and then asks the newest one's date a second time (:475-477) — the "double stat() per winner" that A's commit removed. | READ both. | Technical (speed of Step 4's list on the share). Both walk these folders on the window's own thread, so the window can freeze meanwhile (A-G48 = B-G58). B also lists every version, not only the newest (row 4B-25). |
| 4A-5 | `5c8ce95`, 23-07-2026 01:02, "Detect IMS Reco files with a trailing (N) duplicate suffix" **[GST RULE]** ⚠ | IMS Reco file names may now end with a duplicate marker "(N)" after the 8-digit date (the kind a browser or Windows adds to a second copy of a download — MEDIUM, the commit does not say), for example `ImsReco-ClientA-Delhi-01012025 (1).xlsx` (A:utils/constants.py:23). Such files used to be "not recognised". | **Yes - B does not accept them.** B:utils/constants.py:23, pattern unchanged since the first commit `edac6a3`. | READ both. RAN (C4): core probe, links probe, and A's IMS tests run against B's code (4 of them fail on B). | **GST rule - the author decides** which names count as an IMS reconciliation file, and which of two IMS files goes into Links `B11` ⚠. Details and worked example: 4.4.1. |
| 4A-6 | `c1583d6`, 23-07-2026 01:49, "Fix extractor cache path and version-folder selection" (part 1) | The stand-alone extractor looked for the main program's settings (the last target folder) in whatever folder it was started from, while the main program writes them to `~/.gst_organizer_cache.json` in the user's home folder, so "load the last folder automatically" never worked. Now it reads the home-folder file first (A:power_query_extractor/extractor_main.py:36-63; A:power_query_extractor/gui/extractor_window.py:118). | **No - the problem went away another way.** B deleted the old extractor window (`a62d3cc`). Its Step 4 takes the target from B's own settings (B:gui/pyqt_main_window.py:421), saved in `~/.gst_organizer/settings.json` (B:models/cache_manager.py:17, 35). The old look-up is still written in B:power_query_extractor/extractor_main.py:38-42, but that file cannot be imported at all (its line 9 imports the deleted window; B-G45). | READ both. | None for the merge, except that B's dead `extractor_main.py` should not be carried over. |
| 4A-7 | `c1583d6`, 23-07-2026 01:49 (part 2) | Step 4 picked "the newest" `Version-<DDMMYY HHMM>` folder by its name. The names start with the day, so they do not sort by date across month ends ("300626" sorts after "020726"). Now: the newest by date modified (A:power_query_extractor/gui/extractor_window.py:648) — the rule A already used for the run folder (:631). | **Partly.** B no longer picks one Version folder: it lists all of them (B:gui/pyqt_main_window.py:457; ordered Z to A by name, for display only). But B picks the **run folder** by its name, Z to A (B:gui/pyqt_main_window.py:432-448) — the same day-first trap one level up (B-G12). | READ both. RAN (C4) Step 4 probe: with run folders "Annual Statement-310126 1000" (31-01-2026) and "Annual Statement-010226 0900" (01-02-2026, the newer), A used the 01-02-2026 batch; B used the 31-01-2026 batch. | Technical: B can show an older batch, and write its consolidated workbook and extraction log into it. A's date-modified rule has its own weak spot: anything written into an older folder later makes it "newest" (A-G45). |
| 4A-8 | `0890022`, 23-07-2026 01:54, "Fix late-binding closures in processing log/progress messages" | Screen lines scheduled from inside a loop showed the wrong client or file. This is a *late-binding* bug: a small function made inside a loop reads its variables when it finally runs, not when it was made. In the Dry Run preview every client was shown as the last client. Now the values are fixed at the moment each line is scheduled (A:gui/handlers/processing_handler.py:23-36). | **No - fixed independently** by B's rewrite (`c258047`). B's Dry Run writes each line directly (B:gui/pyqt_main_window.py:326-343), and the Step 3 worker sends finished text through Qt signals (B:gui/controllers/processing_controller.py:20-24). | READ both. RAN (C4) B's real Dry Run with two made-up clients: "ClientA - Delhi" with its one file, then "ClientB - Maharashtra" with its two files, each in the right place. (A's fix was RAN in A's audit.) | None: both are correct now. |

#### 23-07-2026, morning: tests first, then fixes guarded by tests

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-9 | `6c1e1c6`, 23-07-2026 07:58, "Add pytest suite for data-integrity logic (81 tests)" | The first automated tests: 81 tests of file-name recognition, name cleaning, state codes and the scan. At the snapshot A has 288 tests (RAN (A audit): 287 pass; the one failure, `tests/test_setup_validation.py:66`, expects a Windows path to be split at `\`, so it can pass only on Windows). | **Yes - B has no tests** at all (no `tests` folder, no `pytest.ini`). | READ. RAN (C4): A's four pure-logic test files run against a copy of B's code. `test_file_ops.py` cannot even load (B has no `scan_excel_files`). The other three: 66 passed, 7 failed — 4 on IMS "(N)" names, 3 on "Private"/"Limited" shortening. The same 95 tests all pass on A. | A's tests are a ready-made safety net for the new tool. The 7 failures mark exactly where A and B behave differently. Four of them pin a GST rule (4.4.1), so the author must decide that rule before those tests are carried over. |
| 4A-10 | `33d0baf`, 23-07-2026 08:48, "Fix sanitize abbreviation, processing re-entrancy, folder-setting key" (fix 1) | When names are cleaned for files and folders, "Private"/"Limited" now become "Pvt"/"Ltd" only as whole words and in any letter case (A:utils/helpers.py:92-93). Before, they were replaced inside other words too ("Privateer" → "Pvter"), and only when written with exactly that capitalisation. | **Yes.** B:utils/helpers.py:78-79 replaces the text anywhere, and only with that capitalisation. | READ both. RAN (C4) core probe: "Privateer Traders" → A `Privateer_Traders`, B `Pvter_Traders`; "CLIENTA PRIVATE LIMITED" → A `CLIENTA_Pvt_Ltd`, B `CLIENTA_PRIVATE_LIMITED`; "ClientA private limited" → A `ClientA_Pvt_Ltd`, B `ClientA_private_limited`; but "ClientAPrivateLimited" (no spaces) → A leaves it whole, B gives `ClientAPvtLtd`. RAN (C4) A's tests on B: 3 failures here. | Technical (names of the reports, of the copied files and of the bracketed Level-4 folders). A is right for whole words and capitals; B shortens run-together names, which A does not (a limitation A's own tests pin). Not a GST rule: the client *folder* name is built by a different function, identical in both and already whole-word (A:utils/helpers.py:620-644 = B:utils/helpers.py:513-537), and no state code is touched. |
| 4A-11 | `33d0baf`, 23-07-2026 08:48 (fix 2) | A second press of Start while Step 3 is running is now ignored (A:gui/handlers/processing_handler.py:75). Tk buttons wired by a mouse binding react even when shown as "disabled", so a second worker could start. | **No - never had it.** A disabled Qt button sends no click. B disables Step 3's Start during a run (B:gui/views/processing_view.py:134, switched at B:gui/pyqt_main_window.py:138-139). Step 2's START stays clickable, but B's controller then refuses a second run with only a log line (B:gui/controllers/processing_controller.py:259-261) — after that second press has already switched to Step 3 and replaced the window's settings with the values now on screen (B:gui/pyqt_main_window.py:363-364; B-G74). | READ both. | None. Still open in A: its other Step 2 and Step 3 buttons are wired by mouse binding (A:gui/tabs/validation_tab.py:118, 130, 142, 190; A:gui/tabs/processing_tab.py:74, 87), so they cannot be used from the keyboard (A-G62). |
| 4A-12 | `33d0baf`, 23-07-2026 08:48 (fix 3) **[GST RULE]** (state-code key) | The per-client choice "include the client name in folder names" was saved under "<client>-<state name>", while the folder builder looks it up under "<client>-<state code>", so a per-client "Yes" did nothing. Now it is saved under the code (A:gui/handlers/client_handler.py:181, 218). | **No such feature.** B's PyQt5 window (`c258047`) has no control for the per-client switch, and `a62d3cc` then deleted the old Tk screens that had it. Nothing in B writes this setting; its only definition is an empty default (B:models/app_state.py:83), handed on unchanged (B:gui/pyqt_main_window.py:370). B's folder builder reads the same code key (B:core/file_organizer.py:118, 126). B's Step 1 still says "This overrides individual client settings if checked" (B:gui/views/setup_view.py:385), a promise B cannot keep. | READ both. RAN (C4) links probe, identical in both: a setting stored under "ClientA-DL" adds "(ClientA)" to the three category folders; one stored under "ClientA-Delhi" is ignored. | **GST rule (state-code key) - the author decides** whether the per-client option returns (A-F10 / B-F5). If it does, A's key form is the one the folder builder reads. No test protects A's fix (A audit, Appendix X3: the *mutation check* — an old bug put back on purpose into a copy of the code, to see whether a test notices — that keys the setting by the state name again, M16, survived). Details: 4.4.4. |
| 4A-13 | `064c634`, 23-07-2026 08:49, "Log hygiene: downgrade [DEBUG] to logger.debug; drop stray print; add pywin32" (part 1) | 35 "[DEBUG]" lines that filled the log file at INFO level (the normal level, which is written to the log file) — including client names, states and complete dumps of the folder paths and file names written into the Links sheet (`core/excel_handler.py:72-75, 557-656@47a1492`) — were made debug-level (a detail level that the program's log set-up does not write), so they no longer reach `gst_organizer.log` (in `core/excel_handler.py` and `core/file_parser.py`). | **No - fixed independently.** B's `c258047` removed 36 such lines from the log-file code, and also the 14 "[DEBUG]" lines of the on-screen Step 3 log, which A still shows (A:gui/handlers/processing_handler.py:333 and 13 more). A also keeps one `logger.error("[DEBUG] Template file does not exist …")` (A:core/excel_handler.py:83). | READ both. RAN (C4) text search: 15 "[DEBUG]" texts in A's program files, 0 in B's. | Technical. Still in both: each Links value, cut to 50 characters, is written to the log at INFO level (A:core/excel_handler.py:195; B:core/excel_handler.py:165), so client paths stay in the log file (A-G56 = B-G64). |
| 4A-14 | `064c634`, 23-07-2026 08:49 (part 2) | A stray `print(f"PROCESSING: …")` above a function's description in the refresh code was removed. | **No - fixed independently** in B's `c258047`. | RAN (C4) text search: neither `report_processor.py` contains it. | None. |
| 4A-15 | `064c634`, 23-07-2026 08:49 (part 3) | `requirements.txt` now names pywin32, for Windows only (A:requirements.txt:7-9). | **Yes.** B lists pywin32 only as a comment (B:requirements.txt:13-14), although B's window cannot open without it (row 4A-32). | READ both. | Technical: in B, `pip install -r requirements.txt` leaves out the one package the window needs to start. |
| 4A-16 | `ba2a753`, 23-07-2026 08:54, "De-duplicate the shared 'already-open workbook' detection" | The loop "is this workbook already open in Excel?", written twice, is now one shared helper (A:utils/excel_com.py:19-29), used by both Excel paths (A:core/excel_handler.py:337; A:power_query_extractor/core/report_processor.py:709). | **Yes** (the loop is still written twice: B:core/excel_handler.py:274-277; B:power_query_extractor/core/report_processor.py:657-660). | READ both. | Code clarity only; behaviour is identical. In both, the check compares full paths, so a drive-letter path and a `\\server\share` path to the same file do not match (A-G46, same in both). |

#### 25-07-2026: scan speed on the real share, and the new Step 1

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-17 | `4219ffb`, 25-07-2026 18:45, "Fix sanitize_filename: abbreviate underscore-joined names again" | Repairs A's own regression from `33d0baf` (row 4A-10): names joined with underscores (`ClientA_Private_Limited`) are shortened again (A:utils/helpers.py:91). | **No - never had it.** B's plain text replacement always shortened them. | RAN (C4) core probe: `ClientA_Private_Limited` → `ClientA_Pvt_Ltd` in both. | None. |
| 4A-18 | `56aaad9`, 25-07-2026 18:55, "Reuse directory-listing metadata: 81% fewer round-trips per file" | Each file's size and type are now taken from the folder listing itself (A:utils/helpers.py:328-341), which on Windows comes with the listing at no extra cost, instead of asking the server again for every file. | **Yes.** B asks once per file (`file_path.stat()`, B:core/file_parser.py:176): on Windows that is one round trip to the server per file. | READ both. RAN (C4) strace on Linux: here both made 50 metadata requests for 50 files, because Linux needs one either way. The Windows saving comes from Python's documentation of `os.DirEntry.stat` (MEDIUM-HIGH; not run on Windows). | Technical (speed on the share). |
| 4A-19 | `4f84ba9`, 25-07-2026 19:09, "Read Excel signatures concurrently: 11 min -> 41s on the real share" | The 8-byte check that a file really is an Excel file (its *signature*) is now read for up to 16 files at a time (A:utils/helpers.py:271-272, 343-359). The commit message reports, for 2,508 files on the real share, about 11 minutes projected before and 40.78 seconds measured after (not verifiable here). | **Yes.** B reads one file after another (B:core/file_parser.py:157-159, 195-196). | READ both. RAN (C4) speed probe: 100 made-up files, with an artificial 20 ms wait added to every file opened in the source folder (a stand-in for a network round trip): A 0.15 s, B 2.03 s. | Technical (speed): a large difference on the share. In A the scan still runs on the window's own thread, so the window cannot repaint while it runs (A-G48; see row 4B-1). |
| 4A-20 | `322717f`, 25-07-2026 19:24, "Restructure the Setup tab: live validation checklist" (part 1) | A new Step 1: (a) a live checklist that shows ✓ / ○ / ✕ for the source folder, both templates and the target, re-checked 350 ms after typing stops (a *debounce*) with each answer remembered (A:gui/tabs/setup_tab.py:135, 531-608; A:gui/utils/setup_validation.py); (b) Scan is enabled only when the source folder exists; (c) the Step 1 buttons use Tk's `command=`, so "disabled" is respected and the keyboard works; (d) a *tooltip* (a small pop-up text shown while the pointer rests on an element) with the full path on each path box (A:gui/widgets/tooltip.py); (e) the hint beside the length box now states the configured limit, "(default 35 · names over <limit> are shortened)" (A:gui/tabs/setup_tab.py:446-451, 610-615), instead of the hard-coded "⚠️ Names >10 chars will show warning" (`gui/tabs/setup_tab.py:238@47a1492`); (f) the Step 1 wheel handler no longer takes over mouse-wheel scrolling for the whole window (the old tab used `bind_all`, `gui/tabs/setup_tab.py:66@47a1492`; the first replacement was itself repaired in `d64d3a9`, row 4A-22). | **No such feature** for (a), (b) and (d). B checks the source folder only when SCAN is pressed ("Please select a source folder first." / "Folder not found: …", B:gui/pyqt_main_window.py:218-225), and the templates and target only when START is pressed (:353-361). B has no tooltips anywhere. (c) and (f): B's Qt buttons and scroll areas already behave correctly. (e): **Yes** — B's Step 1 still says "Names >10 chars will show warning. This overrides individual client settings if checked." (B:gui/views/setup_view.py:385), but B shows no long-name warning at all (lost in `c258047`; RAN (VERIFY-C4) text search: no long-name warning text in B's code), so the note promises a warning that never comes (B-G72). | READ both. RAN (C4) text search: `setToolTip` occurs 0 times in B. | Only-A feature (A-F4), with A's own gaps: a FILE passes as the source folder in the checklist (A-G20), and the checks run on the window's thread (A-G49). **Regression in this commit:** the old check on what is *typed* into the length box was dropped (A:gui/tabs/setup_tab.py:440-443). RAN (C4) A's real Step 1: typing 8 gives a maximum of 8, and two made-up clients then share one client folder (**[GST RULE]**, state code dropped; row 4B-20 and 4.4.3). Typing letters makes Step 3 fail (A-G26). On (e): A itself still warns at a fixed 10 characters when a client's "📁 Name" cell is switched on (A:gui/handlers/client_handler.py:191), and at the configured limit on Start only when the Step 1 tick is on (A:gui/handlers/processing_handler.py:90-97) (A-G66). |
| 4A-21 | `322717f`, 25-07-2026 19:24 (part 2) **[GST RULE]** ⚠ (display only) | The "EXPECTED FILE NAMES" card on Step 1 is now built from the file-name patterns' own descriptions (A:gui/tabs/setup_tab.py:477-503, loop at :494), e.g. "IMS Reco - ImsReco-ClientName-State-DDMMYYYY.xlsx". | **Partly.** B shows a hand-typed list (B:gui/views/setup_view.py:492-499), e.g. "ImsReco-Client-State-DDMMYYYY", restored in `7e09e5c` from the old v3 screen. | READ both. | **GST rule (display) - the author decides** the wording users see. Neither card mentions `.xls` or the "(N)" forms; A's card does not mention the IMS "(N)" its own pattern accepts. Details: 4.4.6. |
| 4A-22 | `d64d3a9`, 25-07-2026 19:31, "Fix Setup tab scrolling" | Repairs A's own regression from `322717f`: mouse-wheel scrolling did nothing over the content of the new Step 1 (only its thin margins scrolled), and small touch-pad movements were lost (A:gui/tabs/setup_tab.py:137-138, 188-248). | **No - never had it.** B scrolls with Qt's own scroll area (B:gui/views/setup_view.py:95-146). | READ both. | None. |

#### 25-07-2026, evening: preparing Step 4 inside the main window

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-23 | `c1f425e`, 25-07-2026 22:46, "Extractor: marshal worker-thread UI updates onto the Tk thread" | Every screen update from Step 4's worker thread is now handed to the window's own thread (A:power_query_extractor/gui/extractor_window.py:952-1029). Tk is not safe to use from two threads; once the extractor shares one window with the main program, this can cause random freezes (the commit says "intermittent freezes"). | **No - never had it.** B's extraction worker only sends Qt signals (B:gui/controllers/extraction_controller.py:43-46, 67). | READ both. | None for Step 4. Still open in A: the Step 3 worker touches the window directly (A:gui/handlers/processing_handler.py:319-530; A-G70) — see row 4B-2. |
| 4A-24 | `2acffed`, 25-07-2026 22:53, "Extractor: update the refresh timestamps after a run" | Step 4's "Last Refresh Status" column is re-read after a run (it used to show the state from before the run until a manual re-scan); the user's ticks are kept (A:power_query_extractor/gui/extractor_window.py:759-782, 977-987). It also runs after a failed run. | **No - fixed independently** in B's `b0e7d0e`: B rebuilds Step 4's list from the disk after every *completed* extraction (B:gui/pyqt_main_window.py:561). Unlike A, B does not rebuild after an extraction that ends in its error path (:563-565), and its rebuild clears every tick. | READ both. | Technical: both now show fresh status after a normal run. |
| 4A-25 | `a677670`, 25-07-2026 23:01, "Split the extractor into ExtractorPanel + a standalone window shell" | The extractor became one reusable block of screen elements, used both by the stand-alone launcher and (from `bfcb24a`) by Step 4 (A:power_query_extractor/gui/extractor_window.py:85-95, 1032-1063): one implementation, two places. | **No - fixed independently.** B merged the extractor into its one window in `c258047`; B's `launch_extractor.py` opens that window on Tab 4 (B:launch_extractor.py:13-19). | READ both. | None: both have a single extractor. |
| 4A-26 | `2aaa2b1`, 25-07-2026 23:11, "Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one" | With a customised "Refreshed File Suffix" (the text added to a refreshed copy's name), Step 4 showed "Never refreshed" right after a refresh, and "Skip Refresh" (reuse an existing refreshed copy instead of refreshing again) refreshed every report again. Now refreshed copies are recognised by the configured suffix, plus the old `_Refreshed_` marker (A:power_query_extractor/core/refresh_naming.py:14-66; used at A:power_query_extractor/core/report_processor.py:140-159 and A:power_query_extractor/gui/extractor_window.py:45-82, 784-797). The message says both problems were "Reported". | **Yes.** B's Skip Refresh looks only for `<report>_Refreshed_*.xlsx` (B:power_query_extractor/core/report_processor.py:102), and its status column only for `*_Refreshed*` (B:gui/pyqt_main_window.py:467-470) (B-G9). | READ both. RAN (C4) Step 4 probe, suffix `_Done_{timestamp}`, an existing `…_Done_010226_0930.xlsx`: A reused it ("Skip refresh enabled - using existing: …_Done_010226_0930.xlsx"); B said "No existing refreshed file found - will refresh" and made a new copy; B's status column showed that version as "Pending". | Technical: B ignores its own suffix setting. Which earlier copy Skip Refresh may reuse at all is a GST rule for the author, because it decides which file's figures reach the consolidated workbook (Section 5, 9.4, tags it [GST RULE] when in doubt). [Corrected during review - the GST pointer was added; originally "Technical." alone.] A's own limits: very short suffixes such as `_R{timestamp}` make every original report look refreshed (A-G16); the suffix is not remembered between sessions (A-R119); no test protects the Skip Refresh part of the fix (A audit, Appendix X3, mutation M21 survived). |

#### 26-07-2026: shorter folder names, the Step 4 tab and the workflow choice

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-27 | `7347383`, 26-07-2026 08:20, "Shorten category folder names to fit Excel's path limit" **[GST RULE]** ⚠ | The three category folders became `GSTR3B`, `ITC` and `Sales` instead of `GSTR-3B Exports`, `Other ITC related files` and `Sales related files` (A:utils/constants.py:85-95), 9 to 20 characters shorter, because Excel failed on long paths (about 218 characters ⚠). These folder paths are written into the templates' Links cells. | **Yes - B keeps the long names** (B:utils/constants.py:85-89). | READ both. RAN (C4) links probe: the ITC Links cell `B2` holds `…\Version-<DDMMYY HHMM>\GSTR3B` in A and `…\Version-<DDMMYY HHMM>\GSTR-3B Exports` in B. | **GST rule - the author decides** (folder paths written into Links cells ⚠). Technical part only: shorter paths are less likely to exceed Excel's limit; neither project measures a path before creating it (A-G41 = B-G54). Details: 4.4.2. |
| 4A-28 | `bfcb24a`, 26-07-2026 20:40, "Add "Step 4: Extract" tab hosting the extractor in the main app" (part 1) | Step 4 became a fourth tab of the main window (A:gui/tabs/extract_tab.py; A:gui/main_window.py:146), filled from Step 1's target folder without scanning the network at start-up. | **Partly.** The tab itself: fixed independently — B has had a fourth tab since `c258047` (row 4B-14). The start-up part: **yes** — B walks the saved target folder for Step 4 at every start-up, on the window's own thread, before the window appears (B:gui/pyqt_main_window.py:63 → :190 → :419-504; B-G58), which A's tab deliberately avoids (`auto_load=False`, A:gui/tabs/extract_tab.py:79; A:power_query_extractor/gui/extractor_window.py:129-132). | READ both. | None. Note: this commit made A impossible to start without pywin32 for about two and a half hours, until `c463256` (row 4A-32). |
| 4A-29 | `bfcb24a`, 26-07-2026 20:40 (part 2) | Fixes a crash the new tab exposed: the tab-change handler used a fixed list of three tab names and failed when the fourth tab was selected (A:gui/main_window.py:331-338). | **No - never had it.** B has no tab-change handler at all (READ: no `currentChanged` in B:gui/pyqt_main_window.py). | READ both. | None. |
| 4A-30 | `510f7b4`, 26-07-2026 22:58, "Add a workflow choice: organise-only or full pipeline" | Step 1 asks "WHAT SHOULD THIS RUN DO?": organise files and create reports, then stop; or "full pipeline", where Step 3 hands over to Step 4, fills its folder and lists the clients (A:utils/constants.py:163-172; A:gui/handlers/processing_handler.py:136-175). The choice is remembered. | **No such choice.** After every processing run B rebuilds Step 4's list from the disk and switches to Step 4 (B:gui/pyqt_main_window.py:384, 401) — in effect A's "full", always. | READ both. | Only-A feature (A-F5). In neither project does the refresh start by itself, although A's texts promise it ("Also refresh Power Query and extract the data", A-G66). |

#### 26-07-2026, late evening: dark mode and the pywin32 start-up chain

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-31 | `aeb83cf`, 26-07-2026 23:04, "Rebuild dark mode so it restores the exact original colours" | Dark mode used to guess the colours when it was switched off. Now each element's own colours are recorded before anything is changed and are put back exactly (A:gui/utils/dark_mode_manager.py). | **No - different design.** B switches the whole application *stylesheet* (a list of colour and style rules for the window; B:gui/theme/theme_manager.py:414-419, called from B:gui/pyqt_main_window.py:594-596). Switching off applies the light stylesheet again, so there are no recorded colours to restore. | READ both. | Not comparable. Both have their own dark-mode defects: in A, elements created while dark mode is on stay light (A-G64); in B, parts of Steps 1 and 2 stay light and some status text becomes unreadable (B-G69; RAN (B audit) screenshots). |
| 4A-32 | `c463256`, 26-07-2026 23:10, "Fix: app failed to start when pywin32 was missing" | A's window starts without pywin32: the refresh module loads pywin32 inside a guard (A:power_query_extractor/core/report_processor.py:11-36), and Step 4 refuses each report with "pywin32 is not installed for this Python … Install it with: pip install pywin32" (:62-77). Steps 1-3 keep working. | **Yes.** B's window cannot open without pywin32, on any PC: B:main.py:38 → B:gui/pyqt_main_window.py:25 (which loads the package file B:gui/controllers/__init__.py, whose line 4 loads the extraction controller; line 27 would load it too) → B:gui/controllers/extraction_controller.py:10 → B:power_query_extractor/core/report_processor.py:11-15 (unguarded) (B-G32). [Corrected during review - originally `B:gui/pyqt_main_window.py:27`; audit B's self-review (its correction 29, rule B-R124) found that line 25 reaches the controller first; READ both files.] | READ both. RAN (C4) without pywin32: A's main window and Step 4 code import normally; B's main window import stops with "No module named 'pythoncom'"; running B's `main.py` prints "Missing module: No module named 'pythoncom'" and the advice `pip install PyQt5 openpyxl`, and exits with code 1. | Technical. Since `be6418c` A refuses to start on Windows without pywin32 anyway (row 4A-34), so this matters off Windows and for a pywin32 that is installed but broken. A's own defect: the Step 4 notice "Power Query refresh is unavailable" never appears for a missing pywin32 (RAN (C4), re-run (VERIFY-C4): the notice's error value `EXTRACTOR_IMPORT_ERROR` is `None`) — Step 4 looks normal and then refuses each client's reports (A-R129; listed in the A audit under A-G66, "help texts that promise behaviour the code does not have"). |
| 4A-33 | `f363946`, 26-07-2026 23:15, "Check dependencies at startup and report them clearly" | Before the window is loaded, A checks the packages it needs and prints a report (A:utils/requirements_check.py; A:main.py:66-73), instead of a message that blamed missing source files. | **Yes.** B checks nothing; its only message names the wrong packages (B:main.py:49-54). | RAN (C4) on this Linux PC: A prints "[OK] openpyxl", "[not found] pywin32 …", what would not work, the Python being used and the `py`/`python` note. B: see row 4A-32. | Technical (clarity of messages). Both messages go to the console only, which a `pythonw` start does not have (A-G22 / B-G34). |
| 4A-34 | `be6418c`, 26-07-2026 23:20, "Require pywin32 on Windows so reports cannot lose Power Query" (part 1) | On Windows a missing pywin32 now stops the start with the reason "Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them." (A:utils/requirements_check.py:45-60). Without pywin32 the reports would be written by openpyxl, which drops Power Query; the author's Standard 12 says never to let openpyxl save over such a workbook (`rebuild_pipeline_FINAL.md`:845-850). | **No in effect, by accident.** B's window cannot open without pywin32 at all (row 4A-32), so B's openpyxl fallback (B:core/excel_handler.py:83-93) is never reached from the window. But B explains nothing, gives the wrong advice, and its `requirements.txt` leaves pywin32 out (row 4A-15). | READ both. RAN (C4) for B's message. A's Windows behaviour is READ (here `essential` is false because this is Linux). | Technical. A's protection is a start-up check only: it confirms that pywin32 can be *found*, not that it loads, and nothing checks a workbook for Power Query before openpyxl saves it (A-G7). |
| 4A-35 | `be6418c`, 26-07-2026 23:20 (part 2) | The "Press Enter to exit" pause no longer adds a confusing second error when there is no keyboard input (A:main.py:45-54). | **Yes - B has no pause at all** (B:main.py:49-59; removed in `c258047`). Started by double-click, B's console window would close at once, with the message (MEDIUM; not run on Windows). | READ both. | Technical. A's pause does not help when there is no console at all: A stops even earlier, at A:main.py:17 (row 4B-5). |
| 4A-36 | `a5ad487`, 26-07-2026 23:27, "Offer to install missing dependencies, then restart" | If packages are missing, A offers (Yes/No box) to install them with the Python that is running, then restarts itself (A:utils/dependency_installer.py; A:main.py:78-95). | **Yes - B has no offer.** | READ both. | Technical. A's own gaps: the window freezes while pip runs (A-G53); it installs the whole `requirements.txt`, with open-ended versions and pytest (A-G59). Fact only: the author's Standard 13 says "Nothing is installed on users' PCs" (`rebuild_pipeline_FINAL.md`:924). |

#### 26-07-2026 and 27-07-2026: dark-mode and scrolling corrections, hover help

| Row | Commit, date | What it fixed or improved | Does B still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4A-37 | `c89c1d4`, 26-07-2026 23:37, "Fix stale state colours after a dark-mode toggle" | After a dark-mode switch, a Step 1 checklist row could keep the previous state's colour, e.g. green for an input that had become invalid (A:gui/main_window.py:180-194). | **No such feature** (B has no checklist). | READ. | None. |
| 4A-38 | `725fbaf`, 26-07-2026 23:45, "Fix dark mode not fully reverting: tree background and tab padding" | After switching dark mode off, the client list stayed dark grey and the tab strip looked different (A:gui/utils/dark_mode_manager.py:87-102; A:gui/main_window.py:199-207). | **No - different design** (stylesheet swap; row 4A-31). | READ. | None. |
| 4A-39 | `169c7b5`, 26-07-2026 23:50, "Give the Setup checklist dark-mode colours" | The Step 1 checklist got its own dark colours (A:gui/tabs/setup_tab.py:34-67, 531-608), repairing a side effect of `c89c1d4`. | **No such feature** (no checklist). B's Step 1 has the wider problem: its sections keep their light colours in dark mode (B-G69). | READ. | None. |
| 4A-40 | `bf2f1dd`, 27-07-2026 07:48, "Fix scrolling dying after visiting other tabs; add hover help" (part 1) | Scrolling died on Step 1 after the pointer had passed over Step 4's client list, because that list removed every application-wide wheel handler (A:power_query_extractor/gui/extractor_window.py:431-481). A repair of A's own regression from `bfcb24a`. | **No - never had it** (Qt's own scrolling; B has no application-wide wheel handlers). | READ. | None. |
| 4A-41 | `bf2f1dd`, 27-07-2026 07:48 (part 2) | Hover help on every Step 1 control, including consequences not visible on screen (source folder: "nothing here is moved, renamed or deleted"; ITC template: "the template itself is never modified"; long paths: Excel "cannot open files whose full path exceeds roughly 218 characters" ⚠) (A:gui/tabs/setup_tab.py:73-133). | **Yes - B has no tooltips** (RAN (C4) text search: `setToolTip` 0 times in B). | READ; RAN (C4) text search. | Only-A improvement. Some of A's help texts promise more than the code does (A-G66). |
| 4A-42 | `bf2f1dd`, 27-07-2026 07:48 (part 3) | Guard tests that fail if any screen file adds the wheel-binding pattern that broke scrolling twice (A:tests/test_scroll_bindings.py, 64 tests). | **Not applicable** (B has no tests, and Qt does not need this pattern). | READ. | Only relevant if the new tool keeps Tkinter. |

#### A's commits with no change to the program

| Commit, date | What it is | Relevance to B |
|---|---|---|
| `fc279af`, 04-03-2026 22:16, "Project Audit" | A report on the program as it stood at the split (`PROJECT_AUDIT_ANNUALSTATEMENT.md`, 856 lines); no code. Lines 130-232 of the later version hold worked examples with a company-like name, which is not copied here. | None. |
| `f7eb8f1`, 05-03-2026 08:44, "Project Audit 2.0" | A rewrite of the same report. Six of its ten "Known Improvements for Rebuild" (`PROJECT_AUDIT_ANNUALSTATEMENT.md:900-920@f7eb8f1`) were done in July 2026: the extractor's settings path (`c1583d6`), the late-binding lines (`0890022`), guarded pywin32 imports (`c463256`), pywin32 in `requirements.txt` and the "[DEBUG]" clean-up (`064c634`), and whole-word abbreviations (`33d0baf`). Four were not done: the Windows-only `os.startfile` call (still at A:gui/handlers/client_handler.py:255), the two template checks for a "Links" sheet, and the dead-code removal (`calculate_file_hash` still at A:utils/helpers.py:529) (READ). | None. |
| `b276542`, 25-07-2026 22:39, "Add design spec for merging the Power Query Extractor into the main app" | A design note (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md`); no code. | None. |

### 4.3 Project B's fixes and improvements (5 commits)

#### `c258047`, 07-02-2026 11:27, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"

One large AI-assisted rewrite of the window layer. Each row below is one fix or improvement inside it. The same commit also lost several behaviours (Variations sheet always empty, Dry Run without its checks, no question before closing during a run, Export List not wired, the long-name question gone); those are not fixes and are covered in section 3 and the gap map.

| Row | What it fixed or improved | Does A still have it? | How checked | Consequence for the author |
|---|---|---|---|---|
| 4B-1 | The scan runs in a worker thread, so the window stays responsive; a second scan while one runs is refused (B:gui/controllers/scan_controller.py; guard at :57-61). | **Yes.** A's scan runs on the window's own thread (A:gui/handlers/file_handler.py:71-72; A-G48). A's status bar is not repainted during it (A:gui/main_window.py:279-281 only changes the label text), so during A's own measured ~41 s scan of the share the window stays frozen ("Not Responding" on Windows) (HIGH, READ). | READ both. | Technical: B's way is better for large network folders. |
| 4B-2 | Every worker talks to the window only through Qt signals — Step 3 too (B:gui/controllers/processing_controller.py:20-24). | **Partly.** A made Step 4 safe (`c1f425e`) and fixed the scheduled Step 3 lines (`0890022`), but the Step 3 worker still changes the window directly — log lines, message boxes, buttons (A:gui/handlers/processing_handler.py:319-530; A-G70). | READ both. | Technical: B's pattern is the safe one. |
| 4B-3 | Step 4 has a Stop button (B:gui/views/extraction_view.py:272-276; B:gui/controllers/extraction_controller.py:60-63, 167-170). | **Yes.** A's Step 4 cannot be stopped; A's Stop exists for Step 3 only (A:gui/handlers/processing_handler.py:177; A-G50, A-G65). | READ both. | Technical. B's Stop acts only between versions; a refresh in progress runs to its end (B-G39, B-G57). |
| 4B-4 | A second extraction while one runs is refused (B:gui/controllers/extraction_controller.py:142-144). | **Yes.** A's Step 4 start has no "already running" check (A:power_query_extractor/gui/extractor_window.py:845-883), and a Scan or Browse during a run re-enables its Start button (:755-757; A-G65, RAN (A audit)). | READ both. | Technical: two runs at once would fight over the same Excel and files (A-G71 = B-G74). B refuses silently (log line only). |
| 4B-5 | Start-up survives a missing console: the line that re-wraps the console output is guarded (B:main.py:12-13). | **Yes.** A:main.py:17 is unguarded. | RAN (C4), start-up lines run with no output stream (as under `pythonw`): A stops with "AttributeError: 'NoneType' object has no attribute 'buffer'"; B runs on. | Technical. The author's Standard 13 starts apps with `pythonw.exe` (`rebuild_pipeline_FINAL.md`:957, 969). In B, any later start-up error is still invisible without a console, because B only prints it (B-G34). |
| 4B-6 | Start-up no longer creates folders and files inside the program folder (B:main.py:26-29). | **Yes.** A:main.py:29-43 creates `core`, `utils`, `gui`, `logs` and `temp` (and `__init__.py` files where missing) in the program folder at every start. | RAN (C4) on fresh copies: A's start-up step added `logs` and `temp` to the program folder; B's added nothing. | Technical. Standard 13 keeps settings and logs out of the shared app folder (`rebuild_pipeline_FINAL.md`:1039-1040). Both still open `gst_organizer.log` in whatever folder the program is started from (A:main.py:24; B:main.py:20; A-G23 = B-G33). |
| 4B-7 | 36 "[DEBUG]" lines removed from the log-file code, and 14 from the on-screen Step 3 log. | **Partly** — the log-file part was fixed independently by A's `064c634` (row 4A-13); the 14 on-screen lines are still in A (A:gui/handlers/processing_handler.py:333 and 13 more; A-G69). | RAN (C4) text search (A 15, B 0). | Technical. |
| 4B-8 | Code clean-up with no effect on results: all 26 bare `except:` (catch-every-error lines) given a type; the "# ADD THIS …" markers left from pasting code cleared (9 existed at the split; the commit deletes 8, as its message says, and rewords the ninth to a plain comment, B:power_query_extractor/core/report_processor.py:139 — RAN (VERIFY-C4) `git grep` at `47a1492` and `git show c258047`; none is left in B); an unreachable `return` removed from the refresh code; a duplicated, immediately overwritten header block removed from the consolidated-workbook code; the AI tool's permission file `.claude/settings.local.json` taken out of the repository and ignored (B:.gitignore:82). | **Yes.** A still has 23 bare `except:` lines (A:core/excel_handler.py 7, A:power_query_extractor/core/report_processor.py 12, A:power_query_extractor/core/data_consolidator.py 2, A:gui/main_window.py 2); 8 paste markers (e.g. A:utils/helpers.py:17, A:core/excel_handler.py:39; the ninth, in `core/file_parser.py`, went in A's own `56aaad9`); the unreachable `return` (A:power_query_extractor/core/report_processor.py:382-385); the duplicated header block (A:power_query_extractor/core/data_consolidator.py:73-80); and A's repository still contains `.claude/settings.local.json`. (The stray `print` was removed in both: row 4A-14.) | RAN (C4) text searches and `git ls-tree` (git's list of the files a commit holds) of both branches; re-run (VERIFY-C4): bare `except:` 7 + 12 + 2 + 2 = 23 in A, 0 in B, 26 at `47a1492`. | Tidiness only, except one changed outcome on the openpyxl path (row 4B-9). (The commit message also claims the 344-line Step 3 worker was split into 7 named methods; those methods were written into the old Tk file that `a62d3cc` then deleted, and B's Step 3 worker at the snapshot is one `_process` method of about 160 lines, B:gui/controllers/processing_controller.py:42-203 — READ.) |
| 4B-9 | **[GST RULE]** ⚠ (which Links cell receives a value). On the openpyxl path, an error while writing one Links cell is no longer silently skipped: only two error kinds are caught (B:core/excel_handler.py:246-252); any other error makes the whole report fail. | **A still skips silently** (bare `except: pass`, A:core/excel_handler.py:297-303). | RAN (C4) merged-cell probe: a made-up ITC template whose Links `B2` lies inside a merged range (cells joined into one, so only the first can hold a value). A: success; `B2` empty; `B4`, `B5`, `B7`, `B8` filled. B: failure; the template copy left under the report's name with none of the values written. | **GST rule - the author decides** whether a report with a missing Links value may be produced. Technical part only: neither outcome is clean — A's success message hides a missing value; B's failure leaves an unfilled file under the report's name. In B this path is unreachable (the window needs pywin32); in A it runs only without a working pywin32. Details: 4.4.7. |
| 4B-10 | **[GST RULE]** (client key with state code). Step 2 hands back the exact key the scan made ("<client>-<state code>"), stored with each row (B:gui/views/validation_view.py:371; returned at :412-421). | **Yes.** A rebuilds the key from the table's text (A:gui/handlers/client_handler.py:166-170), and Tk turns number-like text into a number: a client written "0042" becomes 42, the key "42-MH" does not exist, and Start stops with "Fatal error" (A-G25). | READ both. RAN (MAP, `scripts/key0042.py`) and RAN (A audit, traces T7 and T12); re-run (VERIFY-C4, `scripts/key0042_tk.py`): A's real `get_selected_clients` on a real Tk table row "0042" / "Maharashtra" reads the values back as `[42, 'Maharashtra']` and returns `42-MH`. | **GST rule (key with state code) - same formula in both**; the difference is technical: B's way avoids A's crash. Details: 4.4.5. |
| 4B-11 | Side effect of the Qt rewrite: every button can be used from the keyboard and respects "disabled"; number boxes accept numbers only, within their range. | **Yes.** A's Step 2 and Step 3 buttons are still mouse-bound (row 4A-11; A-G62). A's length box accepts letters (Step 3 then fails, A-G26) and, since `322717f`, numbers below its range (row 4A-20). | READ both. RAN (C4): typing 8 into the length box — A's real Step 1 keeps 8, B's real Step 1 stays at 35. | Technical: B's widgets are safer. |
| 4B-12 | A run history (the last 50 runs, per client) and a "Last Run" column on Step 2 (B:models/run_history.py:44-71; B:gui/views/validation_view.py:329, 382-385). | **Yes - A keeps no history** (A-F25). | READ both. | Only-B feature (B-F17). B's own gaps: extraction results are stored under a key no screen reads (B-G25); per-client lists are never trimmed (B-G59); extracted figures are kept in plain JSON (a readable text data format) in the user's home folder (B-G64). Fact only: Standard 13 puts settings and logs under `%LOCALAPPDATA%\[App Name]\` (`rebuild_pipeline_FINAL.md`:1039); neither project does. |
| 4B-13 | Settings moved to B's own folder `~/.gst_organizer/`, with a one-time import of the old settings file (B:models/cache_manager.py:17, 33-37, 93-131). The new settings also keep Step 4's wait time, refreshed-file suffix and Skip Refresh between sessions (B:models/app_state.py:21-23; saved at B:gui/pyqt_main_window.py:205-207, restored at :185-187). | **Not applicable** for the folder — A still uses the old file `~/.gst_organizer_cache.json` (A:gui/utils/cache_manager.py:15; A:gui/main_window.py:358-389). **Yes** for the Step 4 options: A starts them at their defaults on every launch (A-R119). | READ both. | A design change, not a fix; remembering the Step 4 options is a small only-B convenience. The import loses the list of recent folders (B-G24, RAN (B audit)). Where the new tool keeps its settings is a fresh decision (Standard 13). |
| 4B-14 | The two programs (organiser and Power Query extractor) became one window with four tabs. | **No - done independently** by A's `a677670` and `bfcb24a` (rows 4A-25, 4A-28). | READ both. | None. |
| 4B-15 | The Dry Run's late-binding bug was not carried into the new code. | **No - fixed independently** by A's `0890022` (row 4A-8). | RAN (C4) B's Dry Run. | None. |

#### `a62d3cc`, 07-02-2026 11:29, and `c21f788`, 07-02-2026 11:57

`a62d3cc` ("Remove old Tkinter GUI files replaced by PyQt5 in v4.0") deleted the 17 old Tkinter screen files. It fixed nothing; it removed the code of the per-client folder-name switch (with its state-name key bug, compare row 4A-12), the client-list export, and the old extractor window — all already out of use since `c258047`, whose new window offers none of them (with its settings-path bug, compare row 4A-6). It left B:power_query_extractor/extractor_main.py:9 importing a deleted file (B-G45), and it stopped B from starting until the next commit.

| Row | Commit, date | What it fixed or improved | Does A still have it? | How checked | Consequence for the author |
|---|---|---|---|---|---|
| 4B-16 | `c21f788`, 07-02-2026 11:57, "Fix gui/__init__.py to import PyQt5 main window instead of deleted Tkinter module" | Repairs B's own start-up break from `a62d3cc` (B:gui/__init__.py:7). | **No - never had it.** A's `gui/__init__.py` imports a file that exists (A:gui/__init__.py). | READ both. RAN (C4): A's main window imports normally. | None. |

#### `7e09e5c`, 08-02-2026 10:52, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections"

| Row | What it fixed or improved | Does A still have it? | How checked | Consequence for the author |
|---|---|---|---|---|
| 4B-17 | Files the scan rejects no longer vanish: under 1 KB → "File too small to be valid Excel"; wrong signature → "Invalid or corrupted Excel file"; unreadable → "Cannot read file: …" (B:core/file_parser.py:171-215). | **Yes.** A drops such files silently, before they are counted anywhere (A:utils/helpers.py:290-306 checks, :343-359 filtering; A-G37). | RAN (C4) core probe, made-up folder: a 104-byte `Sales-ClientA-Delhi-Apr-Mar.xlsx` and a 3 KB `AnnualReport-ClientA-Delhi-2025.xlsx` with a wrong signature are listed by B with those two reasons; in A neither appears anywhere. | Technical. Only partly better: B shows the user only a *count* (B:gui/pyqt_main_window.py:262, 294), and B's summary workbook receives an empty list (B:gui/controllers/processing_controller.py:73; B-G20). A at least names unrecognised files in its summary workbook's Variations sheet (A:gui/handlers/processing_handler.py:224). Neither names every rejected file with its reason. Also, B still drops a file whose details cannot be read at all, with only a program-log warning ("Cannot stat …", B:core/file_parser.py:175-179; READ). |
| 4B-18 | Excel's temporary `~$…` lock files are skipped by name (B:utils/helpers.py:259-260). | **Partly.** A drops an ordinary lock file (about 165 bytes) by size, but a `~$` file of 1 KB or more with a ZIP signature becomes a variation in A (no `~$` rule in A's code). | RAN (C4) core probe: `~$Big-lock.xlsx` (2 KB, ZIP signature) is a variation in A and skipped in B. | Technical, minor. |
| 4B-19 | A live scan log on Step 1, scan progress "Scanning... P% (i/N)" in the status bar, and SCAN FILES disabled while a scan runs (B:gui/pyqt_main_window.py:128-129, 231, 234-242). | **Yes.** A writes progress text to the status bar (A:gui/handlers/file_handler.py:67-68), but the scan holds the window's thread, so the text is not drawn until the scan ends; A's slow part (the signature reads) gives no progress messages at all (A:utils/helpers.py:343-359) (HIGH, READ). | READ both. | Technical: B shows progress; A cannot while its scan runs on the window thread (rows 4A-19, 4B-1). |
| 4B-20 | **[GST RULE]** (state code in the client-folder name). The client-folder length box goes back to 15-100 (B:gui/views/setup_view.py:373-376), repairing B's own `c258047` range of 5-100, with which two clients could fall into one folder. | **Yes, by another route.** A's arrows run 15-100 (A:gui/tabs/setup_tab.py:440-442), but since `322717f` a *typed* value is not checked, and Step 3 uses it as it is (A:gui/handlers/processing_handler.py:194; A:core/file_organizer.py:42). | RAN (C4) on both real Step 1 screens, typing 8: A keeps 8, and two made-up clients ("ClientA Private Limited", Maharashtra; "ClientA Traders", Gujarat) both go into one folder `ClientA ` with no state code; B stays at 35 and gives `ClientA Pvt Ltd-MH` and `ClientA Traders-GJ`. | **GST rule - the author decides** the shortest allowed client-folder name. Technical part only: A's box does not enforce its own stated range on typed input; B's does. Details: 4.4.3. |
| 4B-21 | The processing-mode button texts went back to the wording of the constants ("Re-Run - Add new version folders to existing structure", B:gui/views/setup_view.py:343-346), repairing B's own `c258047` wording. | **No - never had it.** A's mode buttons read the constants (A:gui/tabs/setup_tab.py:401). | READ both. | None. B's help card still says "Re-Run: Overwrite existing" (B:gui/views/setup_view.py:510-512; B-G72); A's help texts have other inaccuracies (A-G66). |
| 4B-22 | `gui/widgets/__init__.py` no longer imports the files deleted by `a62d3cc` (B:gui/widgets/__init__.py:4-7). | **No - never had it.** | READ. | None. |
| 4B-23 | Step 4's list was filled from the run history at start-up and after processing, with a warning when no folder is found. Replaced the next day by `b0e7d0e`; only two unused summary keys remain (B:gui/controllers/processing_controller.py:139-140). | **Not applicable** (superseded inside B). | READ. | None. |
| 4B-24 | An error while showing scan results is caught and logged, and SCAN FILES is enabled again (B:gui/pyqt_main_window.py:244-316). | **No - never had it.** A's scan handler already wraps the scan and the display in one error catcher with a message box (A:gui/handlers/file_handler.py:62-95). | READ both. | None. |

#### `b0e7d0e`, 09-02-2026 08:42, "feat: Enhance extraction process with logging and statistics tracking"

| Row | What it fixed or improved | Does A still have it? | How checked | Consequence for the author |
|---|---|---|---|---|
| 4B-25 | Step 4 lists every Version folder of every client from the disk, and the user ticks which version(s) to refresh (B:gui/pyqt_main_window.py:419-504; B:gui/views/extraction_view.py:291-356). | **Yes - A offers only the newest Version folder of each client** (A:power_query_extractor/gui/extractor_window.py:641-648). | READ both. RAN (C4) Step 4 probe with one client and two Version folders: B listed both ("Version-300126 1000" Refreshed, "Version-010226 0900" Pending); A offered only "Version-010226 0900". | Technical improvement, with B's own problems: only one run folder is ever shown, chosen by name (row 4A-7; B-G12); ticks are cleared at every rebuild (B-G73); "Select All" ticks every version of every client, and the consolidated workbook then has one row per version (B-R87). |
| 4B-26 | Each extraction writes a text log `Extraction_Log_<YYYYMMDD_HHMM>.txt` into the run folder, beside the consolidated workbook (B:gui/controllers/extraction_controller.py:100-121). | **Yes - A writes no per-run log file**; Step 4's messages go to the screen and to the general program log (A-F25). | READ both. RAN (C4) text search: no `Extraction_Log` in A. | Only-B improvement. B's own gaps: two extractions in the same minute overwrite the log (B-G27); warnings that the refresh code writes only to the program log are not in this file (B-R122). |
| 4B-27 | The consolidated workbook goes back into the run folder (B:gui/controllers/extraction_controller.py:92), repairing B's own `c258047` move to the target folder. | **No - never had it** (A:power_query_extractor/gui/extractor_window.py:945 writes into the run folder). | READ both. | None. |
| 4B-28 | A full statistics block after a scan (files, parsed, clients, complete, variations, files per type; B:gui/controllers/scan_controller.py:31; B:gui/pyqt_main_window.py:250-299), restoring the parsed count and the files-per-type list of the v3 summary that B's own `c258047` had dropped (`c258047` kept only total files, clients and variations, `gui/pyqt_main_window.py:230-242@c258047`). | **No - never had it.** A kept its "SCAN RESULTS SUMMARY" (A:gui/main_window.py:299-301). | READ both. | None. B's "Files found" counts only accepted files while its progress line says "Found N files" for all of them (B-G22). |
| 4B-29 | Step 4 ticks can be set only where the report exists, and all start unticked (B:gui/views/extraction_view.py:327-339), repairing B's own `c258047`, where every box could be ticked. | **No - never had it.** A disables a missing report's tick and pre-ticks the existing ones (A:power_query_extractor/gui/extractor_window.py:693-694, 724, 734). | READ both. | None. The default differs (B unticked, A ticked). |
| 4B-30 | Coloured status dots on the tabs (green done, orange ready) (B:gui/pyqt_main_window.py:569-590). | **Yes - A has none**; it shows "Step N: <name>" in the status bar on a tab change (A:gui/main_window.py:331-338). | READ both. | Cosmetic. B's dots are never reset, Step 2 never turns green, and Step 4 turns green even when every refresh failed (B-F29). |
| 4B-31 | After processing, the window switches to Step 4 with a fresh list (B:gui/pyqt_main_window.py:384, 401). | **Partly** — A does this only in its "full pipeline" mode (row 4A-30). | READ both. | See row 4A-30. |
| 4B-32 | Step 1's two help cards fold away when a scan starts (B:gui/views/setup_view.py:590-594; B:gui/pyqt_main_window.py:228). | **Not applicable** (A's Step 1 has no fold-away help cards). | READ. | Cosmetic. |
| 4B-33 | Step 4's list is rebuilt from the disk after an extraction, so the refresh status is current (B:gui/pyqt_main_window.py:561). | **No - fixed independently** by A's `2acffed` (row 4A-24). | READ both. | None. |
| 4B-34 | A "Refresh List" button re-reads Step 4's list from the disk (B:gui/views/extraction_view.py:180-193). | **No - never lacked it.** A's Step 4 has its own Scan button and folder box (A:power_query_extractor/gui/extractor_window.py:611). | READ both. | None. |

### 4.4 The GST-rule rows in detail

Described only. Nothing below says which version is right; that is the author's decision. All clients, names and dates are invented.

#### 4.4.1 IMS Reco names with a "(N)" after the date (row 4A-5) · [GST RULE] · ⚠

| | Project A | Project B |
|---|---|---|
| Pattern | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` (A:utils/constants.py:23): one "(N)" or " (N)" may follow the 8-digit date; the date part is still exactly the 8 digits. | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` (B:utils/constants.py:23): nothing may follow the date. |
| Effect on Links | When both a plain IMS file and its "(1)" copy are present, both are IMS files of the client; the ITC Links cells `B10` (folder) and `B11` (file name) ⚠ (A:utils/constants.py:111-112) take the first IMS file in name order (`ims_files[0]`, A:core/excel_handler.py:648-651), and a space sorts before a dot, so that is the "(1)" copy. | The "(1)" copy is a variation, shown only as a count; `B11` takes the plain file (same first-file rule, B:core/excel_handler.py:581-584). |

**Worked example** (RAN (C4), core probe and links probe; the same inputs through both):

| Input | A | B |
|---|---|---|
| `ImsReco-ClientA-Delhi-01012025.xlsx` | IMS Reco, date `01012025` | IMS Reco, date `01012025` |
| `ImsReco-ClientA-Delhi-01012025 (1).xlsx` | IMS Reco, date `01012025` | not recognised ("No matching pattern found") |
| `ImsReco-ClientA-Delhi-01012025(2).xls` | IMS Reco, date `01012025` | not recognised |
| `ImsReco-ClientA-Delhi-01012025 - Copy.xlsx` | not recognised | not recognised |
| `GSTR3B-ClientA-Delhi-Apr (1).xlsx` (another type, for comparison) | GSTR-3B, month text `Apr (1)` | GSTR-3B, month text `Apr (1)` (same) |
| Source folder for ClientA (Delhi) with the six expected types plus `ImsReco-ClientA-Delhi-31032026 (1).xlsx` | 2 IMS files; status "Has duplicates"; Links `B11` = `ImsReco-ClientA-Delhi-31032026_(1)` | 1 IMS file; status "Complete"; the "(1)" file a variation; Links `B11` = `ImsReco-ClientA-Delhi-31032026` |

**When each version came in.** B: the original pattern from the first commit `edac6a3` (31-08-2025), never changed in B. A: `5c8ce95` (23-07-2026 01:02), "Detect IMS Reco files with a trailing (N) duplicate suffix". Its message says such names "were not detected", that "the date group still captures exactly the 8 digits" and that "all other types already tolerated the suffix via their flexible trailing group". Hint why (A audit): HIGH that real IMS files arrived with such a duplicate marker and were listed as unrecognised (a fix was made for them); only MEDIUM that the marker comes from a browser saving the same download twice — the message does not say where the names came from. A's tests pin the new behaviour (`tests/test_file_patterns.py`); those 4 tests fail on B (RAN (C4)).

**The author decides:** whether an IMS file with "(N)" counts as the client's IMS reconciliation file, and which of two IMS files should be linked in `B11` (A-G6). Neither project's Step 1 card mentions the "(N)" form (4.4.6).

#### 4.4.2 Category folder names written into the templates (row 4A-27) · [GST RULE] · ⚠

| | Project A | Project B |
|---|---|---|
| Folder names | `GSTR3B ({client})`, `ITC ({client})`, `Sales ({client})` (A:utils/constants.py:91-95) | `GSTR-3B Exports ({client})`, `Other ITC related files ({client})`, `Sales related files ({client})` (B:utils/constants.py:85-89) |
| Which file type goes where | GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder (A:core/file_organizer.py:302-315) | the same mapping, identical code (B:core/file_organizer.py:285-298) |
| Bracketed client name | dropped unless the global or per-client option is on (`({client})` removed) | the same code; but in B only the global option can be switched on (4.4.4) |

**Worked example** (RAN (C4) links probe; made-up ClientA, Delhi, Fresh mode; `<target>` and the time stamps shortened):

| Links cell ⚠ | A | B |
|---|---|---|
| ITC `B2` (GSTR-3B folder) | `<target>\Annual Statement-<DDMMYY HHMM>\ClientA-DL\Version-<DDMMYY HHMM>\GSTR3B` | `…\Version-<DDMMYY HHMM>\GSTR-3B Exports` |
| ITC `B7`, `B10` (GSTR-2B and IMS folders) | `…\Version-<DDMMYY HHMM>\ITC` | `…\Version-<DDMMYY HHMM>\Other ITC related files` |
| Sales `B2`, `B8` (Sales and Sales Reco folders) | `…\Version-<DDMMYY HHMM>\Sales` | `…\Version-<DDMMYY HHMM>\Sales related files` |
| File-name cells and the Annual Report folder (ITC `B4`, Sales `B5`) | identical in both | identical in both |
| With the client-name option on | `GSTR3B (ClientA)`, `ITC (ClientA)`, `Sales (ClientA)` | `GSTR-3B Exports (ClientA)`, `Other ITC related files (ClientA)`, `Sales related files (ClientA)` |

**When each version came in.** B: the names from the first commit `edac6a3` (31-08-2025), never changed. A: `7347383` (26-07-2026 08:20), "Shorten category folder names to fit Excel's path limit". The message reports a real run of 3,030 output files on the share: worst path 270 characters, 1,094 files (36%) over Excel's limit of about 218 characters ⚠, "which is why long client names have been failing". It also gives figures modelled on the same 3,030 paths with a planned shorter share root: 235 characters and 27 files over with the shorter root alone, 215 characters and none over with the shorter root plus the short names (the message itself warns the margin is thin, 215 against 218). The code comment (A:utils/constants.py:85-90) presents exactly those modelled pairs (235 / 27 and 215 / none) as "Measured on a real run of 3030 files", without mentioning the shorter share root. The share path the message names is not repeated here.

**The author decides** which names the folders carry. **Not sure — how to check:** the templates are not in either repository, so it cannot be seen here whether their queries take every folder only from the Links cells, or also name "Other ITC related files", "Sales related files" or "GSTR-3B Exports" themselves. Open each real template's Power Query editor and search its queries for those three names. Output folders made before the change keep their old names.

#### 4.4.3 Client-folder names shorter than 15 characters (rows 4A-20 and 4B-20) · [GST RULE]

| | Project A | Project B |
|---|---|---|
| The box | Arrows 15-100 (A:gui/tabs/setup_tab.py:440-442); a typed value is not checked; Step 3 reads it as it is (A:gui/handlers/processing_handler.py:194), and any value above 0 is used (A:core/file_organizer.py:42). | Qt number box 15-100 (B:gui/views/setup_view.py:373-376); typed values outside the range are refused. Greyed out unless "Include client name" is ticked, but it still applies when greyed out (B-G24). |
| How the folder name is cut | `create_client_state_key` (A:utils/helpers.py:620-644): if fewer than 6 characters remain for the client before "-<state code>", the whole text is cut, state code included. | the identical function (B:utils/helpers.py:513-537) |

**Worked example** (RAN (C4); typing 8 into each project's real Step 1, then building the folders for two made-up clients):

| | A | B |
|---|---|---|
| Value after typing 8 | 8 | 35 (the box refuses 8) |
| "ClientA Private Limited", Maharashtra | folder `ClientA ` | folder `ClientA Pvt Ltd-MH` |
| "ClientA Traders", Gujarat | the **same** folder `ClientA ` | folder `ClientA Traders-GJ` |

With the same function in both (RAN (C4)): a maximum of 8 or less drops a two-letter state code (`ClientA `); a maximum of 9 keeps it (`Client-GJ`) but drops a three-letter fallback code (an unknown state "Xyz" gives `XYZ`, and the folder becomes `ClientA T`); at 10, `ClientA-GJ` and `Client-XYZ`; at 15, `ClientA Trad-GJ` and `ClientA Tra-XYZ`. On Windows the trailing space of `ClientA ` is normally dropped; either way the two clients share one folder (MEDIUM, not run on Windows).

**When each version came in.** Shared code at the split: arrows 15-100 plus a typing check that allowed 10-100 (`gui/tabs/setup_tab.py:252-275@47a1492`). A: the typing check was dropped in `322717f` (25-07-2026 19:24, "Restructure the Setup tab: live validation checklist"); the message does not mention it. B: range 5-100 in `c258047` (07-02-2026), back to 15-100 in `7e09e5c` (08-02-2026); neither message mentions it.

**The author decides** the shortest allowed client-folder name, and whether the state code must always survive the cut.

#### 4.4.4 The per-client "client name in folder names" setting (row 4A-12) · [GST RULE]

| | Project A | Project B |
|---|---|---|
| Writers | after a scan, every client set to "off" under "<client>-<state code>" (A:gui/handlers/client_handler.py:49); the per-client toggle (:181) and the global switch (:218) also write under the code since `33d0baf` | none (B's PyQt5 window from `c258047` has no such switch; `a62d3cc` deleted the old screens that had it); the setting is always empty (B:models/app_state.py:83) |
| Reader | the folder builder looks up "<client>-<state code>" (A:core/file_organizer.py:122, 130) | the same look-up (B:core/file_organizer.py:118, 126) |

**Worked example** (RAN (C4) links probe, global option off, made-up ClientA in Delhi): a setting stored under `ClientA-DL` gives the folders `GSTR3B (ClientA)`, `ITC (ClientA)`, `Sales (ClientA)` in A and `GSTR-3B Exports (ClientA)`, `Other ITC related files (ClientA)`, `Sales related files (ClientA)` in B; a setting stored under `ClientA-Delhi` is ignored in both. So the reader is the same; only A's screen can write the setting.

**When each version came in.** Shared code (the first commit): the toggle wrote "<client>-<state name as shown>". A: `33d0baf` (23-07-2026 08:48); its message says "per-client folder-name toggles were silently ignored". B: the switch left the screen with the PyQt5 rewrite `c258047` and its code was deleted in `a62d3cc` (both 07-02-2026), while B's Step 1 note (B:gui/views/setup_view.py:385) still mentions individual client settings.

**The author decides** whether the new tool keeps a per-client choice. No test in A protects the key form.

#### 4.4.5 The client key handed back by Step 2 (row 4B-10) · [GST RULE]

Both projects build the same key at the scan, "<client as written>-<state code>" (shared code). A then rebuilds it from the table's text (A:gui/handlers/client_handler.py:166-170); B returns the key stored with the row (B:gui/views/validation_view.py:371, 412-421).

**Worked example** (RAN (MAP), `scripts/key0042.py`; RAN (A audit) traces T7 and T12): `AnnualReport-0042-Maharashtra-2024-25.xlsx` gives the key `0042-MH` in both. Ticking that row and pressing Start: A rebuilds `42-MH`, which does not exist, and stops with "Fatal error"; B returns `0042-MH` and continues.

**When each version came in.** A: the shared Tkinter code (`edac6a3`). B: `c258047` (07-02-2026, the PyQt5 rewrite; the message does not mention this point).

The key formula and state codes are the same in both; there is no GST-rule difference to decide here. Technical part only: B's way avoids A's crash for clients whose names look like numbers.

#### 4.4.6 The "EXPECTED FILE NAMES" card on Step 1 (row 4A-21) · [GST RULE] · ⚠ (display only)

| | Project A | Project B |
|---|---|---|
| Source of the text | each pattern's own `description` text, read in a loop (A:gui/tabs/setup_tab.py:494); those descriptions are themselves hand-written strings beside the patterns (A:utils/constants.py:19, 27, 35, 43, 51, 59), not generated from them | a hand-typed list (B:gui/views/setup_view.py:492-499) |
| IMS line | "IMS Reco - ImsReco-ClientName-State-DDMMYYYY.xlsx" | "ImsReco-Client-State-DDMMYYYY" |

Both show the same six types in the same order. Neither mentions `.xls`, the ban on hyphens inside a name, or the "(N)" forms; A accepts "(N)" for IMS files (4.4.1) without saying so on its card. The recognition itself is the patterns in `utils/constants.py`, not these texts. **When:** A `322717f` (25-07-2026), whose code comment says the card "can never drift from the regexes the parser actually uses" — true for the list of types, not for the wording, which can still differ from the patterns (A-G66); B `7e09e5c` (08-02-2026), which restored the v3 screen's list. **The author decides** the wording.

#### 4.4.7 A Links cell that cannot be written, on the openpyxl path (row 4B-9) · [GST RULE] · ⚠

| | Project A | Project B |
|---|---|---|
| Catcher around each Links cell | bare `except: pass` (A:core/excel_handler.py:297-303): the cell is skipped | `except (KeyError, ValueError)` with a warning (B:core/excel_handler.py:246-252): any other error ends the whole report |

**Worked example** (RAN (C4) merged-cell probe; a made-up ITC template whose Links `B2` lies inside the merged range `A2:B2`; made-up folder values): A returned success; `B2` stayed empty; `B4` (Annual Report folder), `B5`, `B7`, `B8` were written. B returned failure ("Fallback method error: 'MergedCell' object attribute 'value' is read-only" in the log — RAN (VERIFY-C4) with the log shown); the template copy stayed on disk under the report's name with no Links value written.

**When each version came in.** A: the shared code (`edac6a3`). B: `c258047` (07-02-2026), whose message says "Fixed all 26 bare except: clauses with specific exception types". Reachability: in B never from the window (pywin32 is needed to start); in A only where pywin32 is missing or broken. **The author decides** whether a report with a missing Links value may be produced at all.

### 4.5 Summary

**Fixes and improvements A has that B lacks** (B still has the problem, checked in B's code):

- Scan speed on the share: one folder listing (4A-1, partly), per-file details from the listing (4A-18), signatures read 16 at a time (4A-19).
- Step 3 reports built locally and copied once, with nothing left behind on failure (4A-2); the run folder worked out once per run (4A-3); Step 4's single listings (4A-4).
- Step 4's choice of the newest batch by date modified (4A-7, partly: B picks the run folder by its day-first name).
- Refreshed copies recognised by the configured suffix, in the list and in Skip Refresh (4A-26).
- The test suite (4A-9): A's pure-logic tests fail 7 times on B's code.
- Name cleaning by whole words and any letter case (4A-10); the shared "already open" helper (4A-16).
- Start without pywin32, a clear start-up package report, the Windows rule that refuses to make reports without Power Query, the tolerant pause and the install offer (4A-32 to 4A-36); pywin32 in `requirements.txt` (4A-15).
- The live Step 1 checklist and hover help (4A-20, 4A-41); the length-box hint that states the configured limit, where B still promises a ">10 chars" warning it never shows (4A-20 (e)); no network walk for Step 4 at start-up (4A-28, partly); the workflow choice (4A-30).
- GST-rule rows where only A changed (the author decides): IMS "(N)" names (4A-5), the shorter category folder names (4A-27), the per-client folder-name key (4A-12), the expected-file-name card built from the patterns (4A-21).

**Improvements B has that A lacks** (A still has the problem, checked in A's code):

- The scan in a worker thread, with live progress (4B-1, 4B-19); every worker talking to the window only through Qt signals (4B-2).
- Step 4 Stop, and a second extraction refused (4B-3, 4B-4); every Version folder selectable (4B-25); a per-run extraction log file (4B-26).
- Start without a console, and no folders created in the program folder (4B-5, 4B-6).
- Rejected files counted instead of vanishing, and `~$` lock files skipped (4B-17, 4B-18).
- Qt widgets: keyboard-usable buttons and number-only boxes (4B-11); the on-screen "[DEBUG]" lines removed (4B-7); code clean-up (4B-8); run history (4B-12); Step 4's options remembered between sessions (4B-13); tab dots (4B-30).
- GST-rule rows where B's behaviour differs (the author decides): the length box that refuses typed values below 15 (4B-20; A accepts them since `322717f`), the Step 2 key kept with each row (4B-10; same formula, technical crash in A), and the stricter catcher on the openpyxl path (4B-9).

**Fixes both made independently:** the Dry Run late-binding bug (4A-8 / 4B-15); the "[DEBUG]" lines in the log file and the stray `print` (4A-13, 4A-14 / 4B-7, 4B-8); one window that includes the extractor (4A-25, 4A-28 / 4B-14; but B walks the target folder for Step 4 at every start-up, 4A-28); Step 4's status refreshed after a run (4A-24 / 4B-33); fewer file-system requests per scanned file (4A-1 / 4B-17's commit `7e09e5c`, A further than B); the stand-alone extractor's settings-path bug (fixed in A by 4A-6, removed in B by deleting that window).

**Each project's repairs of its own regressions** (the other never had the problem): A — 4A-17, 4A-22, 4A-29, 4A-37, 4A-39, 4A-40 (and 4A-32, whose problem B has from its own rewrite; 4A-38 is mixed: its client-list colour was A's own regression from `aeb83cf`, but its tab-padding change came from the shared code, `gui/main_window.py:179-180@47a1492`, which B left behind when it moved to Qt); B — 4B-16, 4B-21, 4B-22, 4B-27 to 4B-29 (and 4B-20, whose problem A has by another route).

**For the decision list:** every [GST RULE] row above (4A-5, 4A-12, 4A-21, 4A-27, 4B-9, 4B-10, 4B-20) needs the author's choice before the new tool and its tests are built.

### 4.6 Checks run for this section

All runs were in `/root/merge-scratch/run/C4/`, on copies of the two snapshots (`A/`, `B/`, and `B_Atests/` = B plus A's four pure-logic test files, `pytest.ini` and `conftest.py`; each confirmed identical to `/root/merge-scratch/A-copy` and `/root/merge-scratch/B-copy` before use). `HOME` and `TMPDIR` pointed inside the run folder, so no real settings file was read or written. Python 3.14 (`venv-A` with Tk under a virtual display, `venv-B` with Qt `offscreen`), Linux, no Excel, no pywin32. `scripts/run_all.sh` re-runs the whole set except `lenbox_real_A.py`, `dryrun_probe.py` and the three strace scripts, which were run on their own; the outputs are in `out2/`. The earlier, interrupted attempt's outputs in `out/` agree with `out2/` in every value (only time stamps and timings differ).

| Script | What it checks | Rows |
|---|---|---|
| `scripts/core_probe.py` | IMS names; name cleaning and client-folder key on six made-up names; the scan of one made-up folder (valid, tiny, wrong-signature, `~$`, capital-extension, unrecognised and IMS "(1)" files); "Folders Created" for three clients; a damaged template on the openpyxl path | 4A-1, 4A-2, 4A-3, 4A-5, 4A-10, 4A-17, 4B-17, 4B-18 |
| `scripts/make_strace_src.py`, `scripts/scan_only.py`, `scripts/strace_count.py` | 50 made-up files scanned under strace; requests counted per kind | 4A-1, 4A-18 |
| `scripts/scanspeed_probe.py` | 100 made-up files with a 20 ms wait added to every open in the source folder | 4A-19 |
| `scripts/step4_probe.py`, `scripts/step4_probe_newonly.py` | which run and Version folders Step 4 shows; refreshed-file recognition and Skip Refresh with suffix `_Done_{timestamp}` (B with stand-in pywin32 modules; `time.sleep` replaced in memory) | 4A-7, 4A-26, 4B-25 |
| `scripts/links_probe.py` | made-up ClientA with six types plus an IMS "(1)" copy: IMS count, status, Links values; per-client setting under the code key and the name key | 4A-5, 4A-12, 4A-27 |
| `scripts/lenbox_real_A.py`, `scripts/lenbox_probe.py` | typing 8 into A's real Step 1 (built with the stand-in app object of A's own smoke tests) and B's real Step 1 (`lenbox_probe.py`'s A half is the earlier stand-in box); folder names for two made-up clients | 4A-20, 4B-11, 4B-20 |
| `scripts/startup_probe.py` and B's `main.py` | imports without pywin32; start-up lines without a console; folders created at start; A's package report; B's start message and exit code | 4A-32 to 4A-34, 4B-5, 4B-6, 4B-16 |
| `scripts/merged_probe.py` | a made-up template with a merged Links cell through each openpyxl path | 4B-9 |
| `scripts/dryrun_probe.py` | B's real Dry Run with two made-up clients (stand-in pywin32 modules) | 4A-8, 4B-15 |
| pytest in `B_Atests/` and `A/` | A's four pure-logic test files on B (7 failed, 66 passed; `test_file_ops.py` cannot load) and on A (95 passed) | 4A-9, 4A-10 |
| text searches and `git ls-tree` | "[DEBUG]", bare `except:`, "ADD TH…", `setToolTip`, `Extraction_Log`, `currentChanged`, `~$`; the tracked `.claude/settings.local.json` | 4A-13, 4A-29, 4A-41, 4B-8, 4B-26 |

Not run: anything in Excel itself (COM), anything on Windows, and the real office share. The full Step 2 → Start round trip of 4B-10 was not run here; it rests on READ plus the MAP's and A audit's runs (the key step itself was re-run during verification, below).

**Verification re-run (VERIFY-C4).** Every script above was copied to `/root/merge-scratch/run/VERIFY-C4/scripts/` (paths changed only) and re-run on fresh copies of both snapshots, with `HOME` and `TMPDIR` inside that folder; outputs in `/root/merge-scratch/run/VERIFY-C4/out2/`. Every value above came out the same (timings: A 0.15-0.16 s, B 2.03-2.06 s over three runs; strace counts identical). Added there: `keycut.py` (the client-folder cut at maxima 8, 9, 10 and 15 in both projects — the values in 4.4.3), `key0042_tk.py` (A's real `get_selected_clients` on a real Tk table row "0042": `42-MH`), `merged_probe_log.py` (the merged-cell probe with the log shown: B's "Fallback method error: 'MergedCell' object attribute 'value' is read-only"), `coverage.py` (all 34 + 5 own commits present; rows 4A-1…4A-42 and 4B-1…4B-34 continuous; all 66 cited IDs exist) and `cites.py` (every `A:`/`B:` citation printed from the code). (The MAP's `key0042.py`, re-run for 4.4.5, rewrote its own made-up folders `tmp_key_A`/`tmp_key_B` under `/root/merge-scratch/run/MAP/` with the same content.)

### 4.7 Corrections and clarifications found while writing this section

The code was taken as the truth.

1. **MAP §3, row A-G21** says "typed values unchecked (8 or less drops the state code)". That holds for two-letter state codes. For a three-letter fallback code (an unknown state such as "Xyz" gives `XYZ`), a maximum of 9 also drops it (`ClientA T`) (RAN (C4), `create_client_state_key`, identical in both projects).
2. **B audit, S12own Phase 7 item 5, and S3b line 36** say Step 4's list is rebuilt "after every extraction". The code rebuilds it only after a completed extraction (B:gui/pyqt_main_window.py:561); the error path (`_on_extraction_error`, :563-565) does not rebuild it (READ).
3. **MAP §4.1, row a603c02** gives the commit as "without changing results". That is true on Windows; on Linux and macOS the commit did change the result, because names with capital extensions are found from that commit on (RAN (A audit) on the parent commit and `a603c02`; RAN (C4) at the snapshot: A finds `GSTR3B-ClientB-Delhi-May.XLSX`, B does not).
4. **B's `c258047` message against its code.** The message says "Removed 8 '# ADD THIS' comments" and that the 344-line Step 3 worker was broken up into seven named methods. In the code, 9 markers existed at the split: 8 were deleted and the ninth reworded (B:power_query_extractor/core/report_processor.py:139), so none is left; and the seven methods were written into the old Tk file `gui/handlers/processing_handler.py` (`:112-392@c258047`), which `a62d3cc` deleted — B's working Step 3 worker is one `_process` method (B:gui/controllers/processing_controller.py:42-203). Both points agree with the B audit (X345 CM5; S12BF) (READ; RAN `git grep` / `git show` in VERIFY-C4).
5. **A's real Step 1 confirms the length-box finding.** The A audit and the MAP reached it by their own runs, and the interrupted earlier attempt at this section used a stand-in box built with the same options. This section re-ran it on A's real Step 1 screen (`scripts/lenbox_real_A.py`), with the same result: typing 8 keeps 8, and two clients share one folder.

## 5. Rules side by side

This section puts every business rule (a fixed decision the program makes on its own, such as "a file whose name starts with `ImsReco-` is an IMS reconciliation") of project A next to the matching rule of project B. This first part covers the rule categories 1 to 8 of the two registries (the numbered lists of all features and rules made for each audit): input-file recognition, client identity and name cleaning, state codes, financial year and tax period, expected files and completeness, folder layout and naming, copying and versioning, and report creation with the template "Links" cells (the cells of a sheet called "Links" in each ITC and Sales report template, into which the program writes folder paths and file names; see Category 8). Categories 9 to 16 (Power Query refresh onwards; Power Query is Excel's data-import tool, whose queries live inside the workbook) follow in the next part.

Reminder of the two projects: **A** = `AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026, Tkinter window - Tkinter is Python's built-in window toolkit - "v3.0", 288 tests); **B** = `AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026, PyQt5 window - PyQt5 is a separate, add-on window toolkit - "v4.0", no tests). Despite the names, A is not the older one. Both share their first 12 commits (`edac6a3` 31-08-2025 … `47a1492` 30-10-2025) and split after `47a1492`.

### How to read this section

Each category starts with one table that lists **every** rule of both projects in that category, in A's order, followed by the B rules that no A rule covers. The columns are:

| Column | What it holds |
|---|---|
| A rule / B rule | The rule's ID in each registry (A-R#, B-R#). A rule that one registry files under another category is shown where its partner sits, with a pointer in the other category. |
| What it decides | The rule in plain words. |
| A's version | What A does, with the place in A's code (`A:path:line`). |
| B's version | What B does, with the place in B's code (`B:path:line`). |
| Relation | **same** = the same behaviour and values in both; **different** = the behaviour or values differ (a numbered worked example follows the table); **only A** / **only B** = the rule exists in one project only. |
| Evidence | How it was checked (see the legend below). |

Every **different** row is followed, under the table, by a numbered block (for example "1.1") with: what each project does; one realistic, made-up example that was **run through the real code of both projects** with both results shown; what the difference changes in a client's annual working (which figure, report cell, folder or file); and, for every difference, when each version came in (commit, date, message) and any hint in the code or messages about why.

| Mark | Meaning |
|---|---|
| **[GST RULE]** | The rule decides something only the author can decide: a tax figure; which return, table, sheet or cell a figure belongs to; a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; a due date. When in doubt the rule is tagged (if either registry tags it, it is tagged here). For these rules this section only describes; it never says which version is right and never recommends one. |
| **⚠** | A value that can change over time and must be checked by the author: return and table names, file-naming conventions of the portal, the state list, date formats, template cell positions, Excel limits. |
| **READ** | Established by reading the code of both projects. |
| **RAN** | The real functions of both projects were executed on the same made-up inputs (see "How the runs were done"). |
| **HIGH / MEDIUM / LOW** | Confidence of a statement that could not be run here. |
| `A:path:line`, `B:path:line` | A place in A at `bf2f1dd` or in B at `b0e7d0e`. `path:line@commit` is a place in older code. |
| "technical point" | A remark about crashes, file safety, speed or clarity of messages. Such remarks never judge a GST rule. |

### How the runs were done

- Both projects were copied to `/root/merge-scratch/run/C5a/A` and `/root/merge-scratch/run/C5a/B` and run there with `python -B` (Python 3.14; project A with its own `venv-A`, project B with `venv-B`). Nothing was run inside the original repositories. The scripts are in `/root/merge-scratch/run/C5a/scripts/` and their outputs in `/root/merge-scratch/run/C5a/out/`:
  - `core_probe.py` - file names through both name parsers (the code that splits a file name into type, client, state and period), a source folder with unusual entries, the name cleaner, the client-folder name builder, state codes, the folder builder ("organiser") and the Links values with three test templates;
  - `merged_probe.py` - a Links cell inside a merged cell range;
  - `step3_probe.py` - each project's real Step 3 worker (the code behind "Start Processing") on three made-up clients, including a change of template during the run (`out/step3_A.json`, `out/step3_B.json`) and a damaged template (`out/step3broken_*.json`);
  - `step4_probe.py` and `step4_probe_newonly.py` - each project's real Step 4 code that chooses the run folder and the version folders;
  - `gui_probe_A.py`, `gui_probe_B.py`, `display_probe.py` - the real Step 1 and Step 2 screen code (A's Tkinter screens under `xvfb-run`, a virtual screen; B's Qt screens - Qt is the window toolkit PyQt5 is built on - in Qt's "offscreen" mode, which draws without a monitor);
  - `key42_probe.py`, `collide_probe.py`, `modes_probe.py`, `ims_only1_probe.py` - smaller single-question runs named in the blocks below.
- The clock was fixed at 21-10-2026 10:30 inside each run by replacing the time-stamp function in memory (the code on disk was not changed), so folder names read `Annual Statement-211026 1030` and `Version-211026 1030`.
- All clients are invented: `ClientA` (Delhi), `ClientB PRIVATE LIMITED` (Maharashtra), `ClientD` (Tamil Nadu), `ClientE` and `ClientF` (Delhi), `0042` (Maharashtra; a client written as a number), `Zeta Traders` and `Alpha Traders` (Delhi). The "Excel files" are 2 KB stand-ins that start with the bytes of a real Excel file; they hold no return data.
- This machine is Linux without Excel. Everything Excel itself does through COM (COM is Windows' way for one program to drive another; here Python drives Excel) could only be **READ**. Report building was **RAN** through each project's openpyxl path (openpyxl is a Python library that writes `.xlsx` files without Excel; both projects fall back to it when Excel cannot be driven). Where a result could differ on Windows this is said, with a confidence mark.
- B's window code imports pywin32 (the Python add-on for driving Excel on Windows) when its modules are loaded, so on Linux those modules cannot even be imported. For the B screen runs, empty stand-in modules named `pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui` were placed in memory first (`scripts/fake_win32.py`) only so that the import succeeds. None of their functions is called on the code paths that were run.
- Function-by-function comparison: `funcdiff.py` (it sits in `/root/merge-scratch/run/C5a/` itself, next to `scripts/`, not inside `scripts/`) compared the source text of every function in `utils/helpers.py`, `core/file_parser.py`, `core/file_organizer.py` and `core/excel_handler.py` of both projects (RAN). It is cited below as "funcdiff: IDENTICAL" or "funcdiff: DIFF". (It compares functions by name only, so for a name used twice in one file, such as `__init__`, only the first one is compared.)
- Verification re-run (VERIFY-C5a): every probe above was run again on fresh copies of both projects in `/root/merge-scratch/run/VERIFY-C5a/` (outputs in `out/`), with the same results except the byte sizes of the written report files (the folder paths written into them are longer there). Three extra checks were added there (`scripts2/links_case_probe.py`, `scripts2/spin_probe.py`, and A's pattern and name-cleaning tests run against B's code); they are cited where used.

### Category 1: Input-file recognition

This category decides which files in the source folder (the folder of downloaded returns and working files for the season) are looked at, and which client, state, return type and period each file name is read as. All six file-name patterns are regular expressions (text patterns the computer matches against a file name) kept in one table, `FILE_PATTERNS`, in each project.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R1 | B-R6 | GSTR-2B Reco file name ⚠ [GST RULE] | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$`, any letter case; parts client, state, period. A:utils/constants.py:14-21 (pattern :15) | Identical. B:utils/constants.py:14-21 (pattern :15) | same | RAN core_probe: `GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` gives period `Apr25-Mar26` in both; with " (1)" before `.xlsx` the period becomes `Apr25-Mar26 (1)` in both. Since `edac6a3` (shared). |
| A-R2 | B-R7 | IMS Reco file name ⚠ [GST RULE] | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` - also accepts one "(N)" or " (N)" after the 8-digit date. A:utils/constants.py:23 | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` - nothing may follow the date. B:utils/constants.py:23 | **different** | 1.1 below (RAN) |
| A-R3 | B-R8 | GSTR-3B Export file name ⚠ [GST RULE] | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$`; a leading "(1) " is accepted, a trailing " (1)" ends up inside the month. A:utils/constants.py:30-37 (pattern :31) | Identical. B:utils/constants.py:30-37 (pattern :31) | same | RAN core_probe: `(1) GSTR3B-ClientA-Delhi-Apr.xlsx` gives month `Apr`; `GSTR3B-ClientA-Delhi-Apr (1).xlsx` gives month `Apr (1)` - in both. |
| A-R4 | B-R9 | Sales file name ⚠ [GST RULE] | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$`; exactly four parts (client, state, start month, end month). A:utils/constants.py:38-45 (pattern :39) | Identical. B:utils/constants.py:38-45 (pattern :39) | same | RAN core_probe: `Sales-ClientA-Delhi-Apr25-Mar26 (1).xlsx` gives start `Apr25`, end `Mar26 (1)` in both. |
| A-R5 | B-R10 | Sales Reco file name ⚠ [GST RULE] | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$`. A:utils/constants.py:46-53 (pattern :47) | Identical. B:utils/constants.py:46-53 (pattern :47) | same | RAN core_probe: period `FY2025-26` in both. |
| A-R6 | B-R11 | Annual Report file name ⚠ [GST RULE] | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$`. A:utils/constants.py:54-61 (pattern :55) | Identical. B:utils/constants.py:54-61 (pattern :55) | same | RAN core_probe: year `2025-26`, and `2025-26 (1)` for a " (1)" name, in both. |
| A-R7 | B-R5 | How a name is matched [GST RULE]: patterns tried in a fixed order, first match wins, parts trimmed, empty parts skipped | A:core/file_parser.py:82-134 | B:core/file_parser.py:78-124 | same | READ: the two functions differ only in debug lines for Sales files that write to the log file (A:core/file_parser.py:77-79, 123-126); these lines are shared code that A kept and B removed in `c258047` (07-02-2026), so they are not an A addition. RAN core_probe: 15 names give the same results except the IMS "(N)" names of 1.1. |
| A-R8 | B-R1 | Which files the scan looks at | One folder listing; names ending `.xlsx`, `.xls`, `.xlsm` in any letter case; regular files only (a folder named like a workbook is skipped). A:utils/helpers.py:309-365 (extensions :327, test :330, regular file :296) | Three wildcard searches (searches by name pattern, in which `*` stands for any text: `*.xlsx`, `*.xls`, `*.xlsm`; case-sensitive on Linux), names starting `~$` (Excel's lock files) skipped, folders not excluded. B:utils/helpers.py:249-266 | **different** | 1.2 below (RAN) |
| A-R9 | B-R2 | Minimum size 1024 bytes | Smaller files dropped with no message and no count. A:utils/helpers.py:302-303 | Smaller files listed as the variation (the code's word for a problem file) "File too small to be valid Excel". B:core/file_parser.py:181-191 | **different** | 1.3 below (RAN) |
| A-R10 | B-R3 | Excel signature (first 8 bytes must be `PK\x03\x04` or `\xd0\xcf\x11\xe0`) | Wrong signature dropped with no message and no count. A:utils/helpers.py:263-287 | Listed as the variation "Invalid or corrupted Excel file" (unreadable: "Cannot read file: <error>"). B:core/file_parser.py:193-215 | **different** | 1.4 below (RAN) |
| (no A rule; A skips silently at A:utils/helpers.py:332-338) | B-R4 | A file whose size and dates cannot be read | Skipped; nothing on screen, nothing in the log file | Skipped; nothing on screen; the log file (`gst_organizer.log`) gets "Cannot stat <name>: <error>". B:core/file_parser.py:174-179 | **different** | 1.5 below (RAN) |
| A-R11 | B-R5 | Unrecognised names become variations ("No matching pattern found"; a match with an empty client or state gives "Could not parse filename") [GST RULE] (tagged here only because B-R5, the matching rule this row shares with A-R7, is tagged in B's registry; A's registry does not tag A-R11) | A:core/file_parser.py:130-132, 209-220 | B:core/file_parser.py:120-122, 245-252 | same (for names); B's list additionally holds the rejects of 1.3 to 1.5 | RAN core_probe: `Notes for ClientA.xlsx` → "No matching pattern found" in both. funcdiff: `_process_file` DIFF only in the size/signature part and in how the file details (size, dates) are gathered (A reuses the listing's details through `get_file_info`, B builds them itself from one `stat`, a single request to the file system for the file's size and dates), READ A:core/file_parser.py:186-193, B:core/file_parser.py:173-226. |
| A-R12 | B-R12 | Hint for an unrecognised name ⚠ [GST RULE] | Same hint texts and order; the hint is shown in the Step 3 summary workbook's "Variations" sheet. A:core/file_parser.py:288-306; list handed over at A:gui/handlers/processing_handler.py:224; sheet A:core/excel_handler.py:554-585 | Same hint computed, never shown: the summary workbook receives an empty list. B:core/file_parser.py:316-334; empty list at B:gui/controllers/processing_controller.py:73 | **different** | 1.6 below (RAN) |
| A-R13 | B-R14 | `EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm'}` - never used | A:utils/constants.py:196 | B:utils/constants.py:179 | same (dead in both) | RAN: a search of both copies finds no use outside the definition and the package's re-export list (A:utils/__init__.py:9, 26 = B:utils/__init__.py:9, 26), which only passes the name on |
| A-R14 | B-R16 | `MAX_FILE_SIZE_MB = 100` - never used; no upper size limit anywhere | A:utils/constants.py:198 | B:utils/constants.py:181 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R15 | B-R15 | `EXCEL_SIGNATURES` - never used (the bytes are typed again elsewhere) | A:utils/constants.py:199-202 | B:utils/constants.py:182-185 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R16 | B-R13 (with B-R141) | The patterns' extra labels ⚠ [GST RULE]: `category` and `folder` are copied into every file record and never read in either; the `description` (example name) is read by A's Step 1 card and by nothing in B | A:utils/constants.py:17-19, 25-27, 33-35, 41-43, 49-51, 57-59; card A:gui/tabs/setup_tab.py:477-503 | B:utils/constants.py:17-19, 25-27, 33-35, 41-43, 49-51, 57-59; B's card is typed by hand at B:gui/views/setup_view.py:488-499 | **different** (only the example names shown on Step 1) | 1.7 below (RAN) |

#### 1.1 IMS Reco names with a browser "(N)" after the date (A-R2 / B-R7) · [GST RULE] · ⚠

**What A does.** A name such as `ImsReco-ClientA-Delhi-31032026 (1).xlsx` or `…31032026(2).xlsx` (the "(1)" a web browser adds when the same file is downloaded twice) is an IMS Reco file of ClientA; the date part stays `31032026`. Only one such marker is allowed. A:utils/constants.py:23.

**What B does.** Nothing may follow the 8 digits; the same name is unrecognised ("No matching pattern found"). B:utils/constants.py:23.

**Worked example (RAN, `core_probe.py` and `ims_only1_probe.py`).**

| Input | A | B |
|---|---|---|
| `ImsReco-ClientA-Delhi-31032026 (1).xlsx` | IMS Reco, client `ClientA`, state `Delhi`, date `31032026` | variation "No matching pattern found" |
| `ImsReco-ClientA-Delhi-31032026(2).xlsx` | IMS Reco, date `31032026` | variation "No matching pattern found" |
| `ImsReco-ClientA-Delhi-31032026 (1) (2).xlsx` | variation (only one marker allowed) | variation |
| ClientA's folder holds both `ImsReco-ClientA-Delhi-31032026.xlsx` and the " (1)" copy, plus the other five types: Step 2 status | "Has duplicates", remark "Multiple IMS Reco (2 files)" | "Complete" |
| Same folder: files copied for ClientA | 8 (the duplicate as `ImsReco-ClientA-Delhi-31032026_(1).xlsx` in the ITC folder) | 7 |
| Same folder: ITC report Links cell B11 (the IMS file name the template reads) | `ImsReco-ClientA-Delhi-31032026_(1)` - the duplicate, because a space sorts before a dot (rule A-R57) | `ImsReco-ClientA-Delhi-31032026` |
| ClientF has **only** `ImsReco-ClientF-Delhi-31032026 (1).xlsx` (plus the other five types): status | "Complete"; B10 and B11 written (`…_(1)`) | "Missing 1 files" (IMS Reco); B10 and B11 not written, so the template's own text stays in them; the file counted only as a variation |
| Step 3 summary workbook for the same three-client folder (`step3_probe.py`) | "Total Files" 10; ClientA "Files" 7 | "Total Files" 9; ClientA "Files" 6 |

**What it changes in the annual working.** Which IMS reconciliation file feeds the ITC report's IMS part (Links B10/B11, and from there the IMS figures the template's queries compute). With a re-downloaded IMS file in the folder, A links the "(1)" copy (whether that is the newer download is not checked by either project), and Step 2 flags the client "Has duplicates"; B ignores the "(1)" copy and links the original. With only a "(1)" download, A treats IMS Reco as present; B reports the client as missing IMS Reco and leaves the IMS cells of the ITC report as the template had them. For all other types a trailing " (1)" is accepted by both and simply becomes part of the period text (RAN, rows A-R1, A-R3 to A-R6).

**When each version came in.** B's pattern is the original one from `edac6a3` (31-08-2025, "Initial commit"), never changed in B. A's came in `5c8ce95` (23-07-2026, "Detect IMS Reco files with a trailing (N) duplicate suffix"). Its message says such names "were not detected: the pattern required exactly 8 digits immediately before the extension", that "the date group still captures exactly the 8 digits, so metadata is unchanged", and that "all other types already tolerated the suffix via their flexible trailing group". A's tests pin the new behaviour (A:tests/test_file_patterns.py:46-60, the three IMS names at :48-50, and :63-67); run against B's code these tests fail (four failures: the three IMS names and the date test; RAN again in `run/VERIFY-C5a/Btests`; `notes2/A_tests.md` records these together with the three name-cleaning failures of 2.2 as "7 failures"). Neither project's Step 1 help mentions the "(N)" form (1.7). Which version to keep is the author's decision.

#### 1.2 Which files the scan looks at (A-R8 / B-R1)

**What A does.** One listing of the top level of the source folder. A name is kept if, in lower case, it ends in `.xlsx`, `.xls` or `.xlsm`; it must be a regular file (A:utils/helpers.py:296), at least 1 KB (1.3) and carry an Excel signature (1.4). Anything else is dropped without a message. A:utils/helpers.py:309-365.

**What B does.** Three wildcard searches (`*.xlsx`, `*.xls`, `*.xlsm`; `*` stands for any text) of the top level; names starting `~$` (Excel's lock files, made while a workbook is open) are dropped; everything else found goes to the parser, which applies 1.3 to 1.5. On Linux the wildcard search respects letter case; on Windows Python's search ignores it (HIGH). B:utils/helpers.py:249-266; parser B:core/file_parser.py:171-256.

**Worked example (RAN, `core_probe.py`, folder `src_discovery`; B ran on Linux).**

| Entry in the source folder | A | B |
|---|---|---|
| `GSTR3B-ClientA-Delhi-Apr.xlsx`, `GSTR3B-ClientA-Delhi-May.xls` | listed, GSTR-3B | listed, GSTR-3B |
| `GSTR3B-ClientA-Delhi-Jun.XLSX` (capital extension) | listed, GSTR-3B | not found at all on Linux; on Windows found (HIGH, not run) |
| `~$GSTR3B-ClientA-Delhi-Apr.xlsx` (165-byte lock file) | dropped (under 1 KB), no message | skipped by its `~$` name |
| `~$Big-lock.xlsx` (2 KB, Excel signature) | listed → variation "No matching pattern found" | skipped by its `~$` name |
| `GSTR3B-ClientA-Delhi-Sep.xlsm` | listed → variation (no pattern accepts `.xlsm`) | listed → the same variation |
| `Old folder.xlsx` (a **folder**, not a file) | skipped, no message | listed → variation "Cannot read file: Is a directory: …" (on Linux). On Windows the size of a folder is reported as 0, so the variation would probably read "File too small to be valid Excel" (MEDIUM, not run) |
| `GSTR3B-ClientA-Delhi-Dec.csv`, `sub/GSTR3B-ClientA-Delhi-Nov.xlsx` | ignored (wrong extension; sub-folder) | ignored |
| Result for ClientA | three GSTR-3B months (Apr, May, Jun) | two (Apr, May) on Linux |

The same folder also holds `Claims register ClientA.xlsx` (2 KB, Excel signature; listed by both and a variation in both, see 1.6) and the July, August and October files of 1.3 to 1.5; the screen counts of 5.2 are for this whole folder (RAN, `out/core_A.json` and `out/core_B.json`, keys `disc.*`).

**What it changes in the annual working.** On the office's Windows PCs the capital-letter case does not apply (HIGH), so in practice the two scans differ in three places: a lock-named file of 1 KB or more (rare) is a variation in A and invisible in B; a folder whose name ends in `.xlsx` adds a variation in B and is invisible in A; and files that are too small or not Excel are invisible in A but counted in B (1.3, 1.4). None of this changes which genuine return file goes where.

**When each version came in.** At the split both projects used B's three wildcard searches and dropped small and non-Excel files silently (`47a1492:utils/helpers.py:249-266`, where `validate_excel_file` filtered them). A replaced the searches with one listing in `a603c02` (23-07-2026, "Speed up network file scanning (behavior-preserving)"), then reused the listing's file details in `56aaad9` (25-07-2026, "Reuse directory-listing metadata: 81% fewer round-trips per file") and read the signatures 16 at a time in `4f84ba9` (25-07-2026, "Read Excel signatures concurrently: 11 min -> 41s on the real share"); the messages give speed on the network share (the office's shared network drive where the season's files are kept) as the reason. B added the `~$` filter and moved the size and signature checks into the parser in `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections"; the message says only "Added functionality to handle Excel file selection and validation more efficiently"). Not a GST rule.

#### 1.3 Files under 1 KB (A-R9 / B-R2)

**What A does.** Drops them with no message, no count and no log line. A:utils/helpers.py:302-303.
**What B does.** Lists each one as a variation "File too small to be valid Excel" with its size. B:core/file_parser.py:181-191.

**Worked example (RAN, `core_probe.py`).** `GSTR3B-ClientA-Delhi-Jul.xlsx` of 500 bytes (for example a download that broke off): A - not in any list or count; B - variation `{"filename": "GSTR3B-ClientA-Delhi-Jul.xlsx", "issue": "File too small to be valid Excel", "size": 500}`, counted under "Variations" on Steps 1 and 2.

**What it changes.** In A a broken July GSTR-3B download disappears without trace; ClientA simply has one month fewer in the GSTR-3B folder. In B the variation count rises by one, but the name of the file is shown nowhere on screen, because B never displays its variation list (1.6). Technical point: B at least counts the file; neither tells the user which file it was.

**When.** Silent drop since `edac6a3` (shared; at the split `47a1492:utils/helpers.py:232`, inside `validate_excel_file`, applied as a filter at `:260`). B's variation since `7e09e5c` (08-02-2026). A unchanged in behaviour. Not a GST rule.

#### 1.4 Files without an Excel signature (A-R10 / B-R3)

**What A does.** Reads the first 8 bytes; anything that does not start like a ZIP-based Excel file (`PK\x03\x04`) or an old `.xls` (`\xd0\xcf\x11\xe0`) is dropped with no message. A file that cannot be opened for reading is dropped too, with nothing on screen but a line "Error validating Excel file …" in the log file (A:utils/helpers.py:285-287). A:utils/helpers.py:263-287.
**What B does.** The same test, but the file becomes the variation "Invalid or corrupted Excel file"; if it cannot be read, "Cannot read file: <error>". B:core/file_parser.py:193-215.

**Worked example (RAN, `core_probe.py`).** `GSTR3B-ClientA-Delhi-Aug.xlsx` that is really a 2,240-byte text file: A - invisible; B - variation "Invalid or corrupted Excel file", size 2240.

**What it changes.** As 1.3: in A a mis-saved August GSTR-3B (for example a portal error page saved as `.xlsx`) vanishes from the client's working; in B it is counted but not named on screen.

**When.** As 1.3: silent in the shared code (`47a1492`), variation in B since `7e09e5c` (08-02-2026). The variation text itself is older: the shared `_process_file` checked each file a second time and would have listed "Invalid or corrupted Excel file" (`47a1492:core/file_parser.py:185-193`, from `edac6a3`), but it could never fire because the earlier search had already dropped such files; A removed that second check in `a603c02` (23-07-2026), and B's `7e09e5c` made the text reachable by moving the checks into the parser (READ). Not a GST rule.

#### 1.5 Entries whose size cannot be read (no A rule / B-R4)

**What A does.** If the listing finds a name but its details cannot be read (for example a link whose target has gone, or a file deleted from the share between listing and reading), the entry is skipped with no message and no log line. A:utils/helpers.py:332-338.
**What B does.** The same skip, but the log file gets "Cannot stat <name>: <error>". B:core/file_parser.py:174-179. Any other per-file error goes into a list nothing reads, in both (A:core/file_parser.py:222-224; B:core/file_parser.py:254-256).

**Worked example (RAN, `core_probe.py`).** A broken link named `GSTR3B-ClientA-Delhi-Oct.xlsx`: A - nothing anywhere; B - nothing on screen, log line "WARNING core.file_parser: Cannot stat GSTR3B-ClientA-Delhi-Oct.xlsx: [Errno 2] No such file or directory: …".

**What it changes.** Only whether a trace is left in the log file for later investigation; on screen both are silent.

**When.** A's present skip (A:utils/helpers.py:332-338) came with `56aaad9` (25-07-2026, "Reuse directory-listing metadata: 81% fewer round-trips per file"), which first read the file details inside the single listing; before that (in the shared code, and still after `a603c02` of 23-07-2026), such a file was dropped silently inside `validate_excel_file` (`47a1492:utils/helpers.py:224`, applied as a filter at `:260`). The behaviour (silent) did not change. B's log line since `7e09e5c` (08-02-2026). Not a GST rule.

#### 1.6 Showing the hint for an unrecognised file name (A-R12 / B-R12) · [GST RULE] · ⚠

**What both do.** For every unrecognised name a hint is chosen by plain text tests on the lower-cased name, first hit wins: contains "gstr" and "2b" → `GSTR-2B-Reco-ClientName-State-Period.xlsx`; "ims" → `ImsReco-ClientName-State-DDMMYYYY.xlsx` ⚠; "gstr" and "3b" → `GSTR3B-ClientName-State-Month.xlsx`; "sales" and "reco" → `SalesReco-ClientName-State-Period.xlsx`; "sales" → `Sales-ClientName-State-StartMonth-EndMonth.xlsx`; "annual" → `AnnualReport-ClientName-State-Year.xlsx`; otherwise "Check file naming convention". funcdiff: `_suggest_pattern` IDENTICAL.

**What differs.** A hands the scan's variation list to the summary workbook (A:gui/handlers/processing_handler.py:224), which writes it to the "Variations" sheet with the hint as "Expected Pattern" (A:core/excel_handler.py:554-585). B's worker builds its summary with `'variations': []` (B:gui/controllers/processing_controller.py:73); the scan's list stays in memory and is never displayed.

**Worked example (RAN, `core_probe.py` and `step3_probe.py`).** Source folder with `Claims register ClientA.xlsx` (and, for B, the IMS "(1)" file of 1.1):

| | A | B |
|---|---|---|
| Hint computed for `Claims register ClientA.xlsx` | `ImsReco-ClientName-State-DDMMYYYY.xlsx` (the word "claims" contains "ims") | the same |
| Hint for `IMS Reco-ClientA-Delhi-31032026.xlsx` (space in "IMS Reco") | `ImsReco-ClientName-State-DDMMYYYY.xlsx` | the same |
| Summary workbook, sheet "Variations" | one row: `Claims register ClientA.xlsx` · "No matching pattern found" · `ImsReco-ClientName-State-DDMMYYYY.xlsx` · 2004 · "Rename file" | "No file variations found" (although B's scan held two variations) |

**What it changes.** In A, after Step 3 the staff can open `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` and see which source files were not used and how they should have been named (for example an IMS download saved as `IMS Reco-…` with a space). In B the only trace is the number "Variations: n" on Steps 1 and 2; the file names and hints are not shown anywhere, so a mis-named file of a client is noticed only if someone spots the missing type.

**When.** The hint and the Variations sheet date from `edac6a3` (31-08-2025, shared); at the split the Tk code (the shared Tkinter window code; Tk is the toolkit Tkinter drives) passed the real list (`47a1492:gui/handlers/processing_handler.py:161`). B's empty list came with `c258047` (07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"); its message does not mention variations. The new worker is given the settings, the client data, the ticked keys and the (always empty) per-client folder flags (B:gui/pyqt_main_window.py:366-371), but not the scan's variation list (READ, HIGH that this is how the list was lost; why is not stated anywhere). Which version to keep is the author's decision.

#### 1.7 Example file names shown on Step 1 (A-R16 / B-R13, with B-R141) · [GST RULE] · ⚠

**What A does.** The "EXPECTED FILE NAMES" card on Step 1 is built in a loop over `FILE_PATTERNS` and shows each pattern's type and `description` text (A:gui/tabs/setup_tab.py:477-503; texts at A:utils/constants.py:19, 27, 35, 43, 51, 59).
**What B does.** The card "EXPECTED FILE NAMES / MODES" shows a list typed directly in the screen code (B:gui/views/setup_view.py:488-499); the patterns' `description` texts are never read.

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py` - the text of the real widgets).**

| Type | A's card | B's card |
|---|---|---|
| GSTR-2B Reco | `GSTR-2B-Reco-ClientName-State-Period.xlsx` | `GSTR-2B-Reco-Client-State-Period` |
| IMS Reco | `ImsReco-ClientName-State-DDMMYYYY.xlsx` | `ImsReco-Client-State-DDMMYYYY` |
| GSTR-3B Export | `GSTR3B-ClientName-State-Month.xlsx` | `GSTR3B-Client-State-Month` |
| Sales | `Sales-ClientName-State-StartMonth-EndMonth.xlsx` | `Sales-Client-State-Start-End` |
| Sales Reco | `SalesReco-ClientName-State-Period.xlsx` | `SalesReco-Client-State-Period` |
| Annual Report | `AnnualReport-ClientName-State-Year.xlsx` | `AnnualReport-Client-State-Year` |

Both name the same six types in the same order. A shows `.xlsx` and the type label; B shows no extension. Neither mentions that `.xls` is also accepted, that a hyphen inside a client or state name breaks the reading (rule A-R7/B-R5), or the "(1)" forms each project accepts (1.1). Technical point: A's code comment says the card "can never drift from the regexes the parser actually uses" (A:gui/tabs/setup_tab.py:478-479), but the card shows the hand-written `description` text kept next to each pattern, not the pattern itself; the IMS description (A:utils/constants.py:27) does not reflect the "(N)" that A's IMS pattern (A:utils/constants.py:23) accepts.

**What it changes.** Only what staff are told when they rename downloads at the start of the season; the recognition itself is 1.1 and the other pattern rows.

**When.** B's list is the shared original text: the Tk card at the split held the same six hand-typed lines (`47a1492:gui/tabs/setup_tab.py:362-369`, from `edac6a3`). B's first PyQt5 screen (`c258047`, 07-02-2026) showed an HTML list with `.xlsx` names (`c258047:gui/views/setup_view.py:169` for IMS); `7e09e5c` (08-02-2026) brought back the short hand-typed list. A replaced the list in `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"). Which wording to keep is the author's decision.

### Category 2: Client identity and name cleaning

This category decides which files belong to one client, and how a client's name is cleaned before it is used in folder and file names.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R17 | B-R17 | The client key [GST RULE]: files are grouped under "<client as written in the file name>-<state code>"; the key handed from Step 2 to Step 3 | Built at the scan, A:core/file_parser.py:226-236. Step 2 rebuilds the key from the text of the table row, A:gui/handlers/client_handler.py:161-173 | Built the same way, B:core/file_parser.py:258-268. Step 2 stores the key inside the row and hands it back unchanged, B:gui/views/validation_view.py:371, 412-421 | **different** (only the key Step 2 hands back) | 2.1 below (RAN) |
| A-R18 | B-R18 | Client and state text checks (empty, shorter than 2, longer than 100, characters `< > : " / \ ? *` and the vertical bar); the result is stored and never read | A:utils/helpers.py:428-449; called A:core/file_parser.py:108-119 | B:utils/helpers.py:321-342; called B:core/file_parser.py:104-115 | same (no effect in either) | funcdiff: `validate_client_name` IDENTICAL. RAN core_probe: `''`, `'A'`, 101 characters, `'a<b'`, `'Delhi'` give the same five answers in both. |
| A-R19 | B-R19 | "Private" → "Pvt", "Limited" → "Ltd" | File names (copies, report names, the bracketed name in category folders): whole words, any letter case, after turning underscores into spaces, A:utils/helpers.py:91-93. Client folder: whole words, any case, A:utils/helpers.py:624-625 | File names: plain, case-sensitive replacement anywhere in the text, B:utils/helpers.py:77-79. Client folder: whole words, any case (identical to A), B:utils/helpers.py:517-518 | **different** for file names; same for the client folder | 2.2 below (RAN) |
| A-R20 | B-R20 | Name cleaning (`sanitize_filename`): characters replaced, runs of spaces and underscores joined, length cut so that name plus extension is at most 190 characters (200 − 10; that is 185 + `.xlsx`, 186 + `.xls`, 190 with no extension; RAN in both), "unnamed" for empty text | A:utils/helpers.py:68-126 | B:utils/helpers.py:66-112 | **different** only through the abbreviation step of A-R19 / B-R19 | funcdiff: DIFF. RAN core_probe: the other test names give the same result in both, e.g. `ClientA/Delhi Branch` → `Delhi_Branch` (text before a slash is lost in both), `ClientB Pvt. Ltd` → `ClientB_Pvt. Ltd` (text after the last dot treated as an extension in both), `ClientA's #1, Shop!` → `ClientA_s_1_Shop`, `''` and `'___'` → `unnamed`. |

#### 2.1 The client key handed back by Step 2 (A-R17 / B-R17) · [GST RULE]

**What both do at the scan.** A client is "<client text exactly as written>-<state code>". For `AnnualReport-0042-Maharashtra-2025-26.xlsx` (a client written as an office code number) the key is `0042-MH` in both (RAN).

**What A does when the user ticks clients.** Step 2's table (a Tk "Treeview") gives back the cell texts; Tk turns number-looking text into numbers, so `0042` comes back as the number 42. A rebuilds the key as `f"{values[0]}-{state_code}"` and gets `42-MH`, which does not exist (A:gui/handlers/client_handler.py:167-170). Step 3 looks the key up outside its per-client protection (A:gui/handlers/processing_handler.py:232), so the run stops.

**What B does.** The key is stored inside the row when the table is filled and handed back as it is (B:gui/views/validation_view.py:371, 418).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`, `key42_probe.py`).**

| | A | B |
|---|---|---|
| Value read back from the Step 2 table for client `0042` | `42` (a number) | `0042` |
| Keys handed to Step 3 after ticking all three clients | `42-MH`, `ClientA-DL`, `ClientE-DL` (`42-MH` is not a scanned client) | `0042-MH`, `ClientA-DL`, `ClientE-DL` |
| Start Processing with `42-MH` first | "💥 Fatal error: '42-MH'" and an error box; nothing written for any client, no summary workbook | (not affected) |

**What it changes.** A client whose name in the file names is a pure number with leading zeros cannot be processed in A, and the run stops at that client: clients ticked after it are not processed and no summary workbook is written (clients before it are processed; READ, HIGH, from the loop at A:gui/handlers/processing_handler.py:228-232). Other number-looking names were not run. Technical point: this is a crash in A's screen code; the key-building rule itself is the same in both projects, and nothing here judges it.

**When.** Key building since `edac6a3` (shared). A's rebuild from the table text is the shared Tk code (`47a1492:gui/handlers/client_handler.py:161-170`). B's stored key came with the PyQt5 screen in `c258047` (07-02-2026); the message does not mention it.

#### 2.2 "Private"/"Limited" in file names (A-R19, A-R20 / B-R19, B-R20)

**What A does.** In every cleaned file name: runs of spaces and underscores become one space, then whole words "Private" and "Limited" in any letter case become "Pvt" and "Ltd" (A:utils/helpers.py:91-93). Words that merely contain them ("Privateer", "Unlimited") are left alone.
**What B does.** `base_name.replace('Private', 'Pvt')` and `.replace('Limited', 'Ltd')`: exact capitals only, anywhere in the text (B:utils/helpers.py:77-79).
**Same in both.** The client-folder name (Level 2) uses whole words in any case and does not first turn underscores into spaces (A:utils/helpers.py:624-625 = B:utils/helpers.py:517-518).

**Worked example (RAN, `core_probe.py`, `collide_probe.py`, `step3_probe.py`).**

| Client text in the file names | A: cleaned name | B: cleaned name | Client folder (both) |
|---|---|---|---|
| `ClientB PRIVATE LIMITED` | `ClientB_Pvt_Ltd` | `ClientB_PRIVATE_LIMITED` | `ClientB Pvt Ltd-MH` |
| `ClientB Private Limited` | `ClientB_Pvt_Ltd` | `ClientB_Pvt_Ltd` | `ClientB Pvt Ltd-MH` |
| `ClientB_Private_Limited` | `ClientB_Pvt_Ltd` | `ClientB_Pvt_Ltd` | `ClientB_Private_Limited-MH` |
| `Privateer Traders Unlimited` | `Privateer_Traders_Unlimited` | `Pvter_Traders_Unlimited` | `Privateer Traders Unlimited-DL` |
| `Delimited Exports Private Limited` | `Delimited_Exports_Pvt_Ltd` | `Delimited_Exports_Pvt_Ltd` | - |

Where the cleaned name is used, for `ClientB PRIVATE LIMITED` (Maharashtra):

| Use | A | B |
|---|---|---|
| Copy of the April GSTR-3B | `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx` | `GSTR3B-ClientB_PRIVATE_LIMITED-Maharashtra-Apr.xlsx` |
| ITC report name | `ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx` | `ITC_Report_ClientB_PRIVATE_LIMITED_MH_211026_1030.xlsx` |
| Category folder with the include-name tick on | `GSTR3B (ClientB_Pvt_Ltd)` | `GSTR-3B Exports (ClientB_PRIVATE_LIMITED)` |
| ITC Links cell B8 (GSTR-2B file name) | `GSTR-2B-Reco-ClientB_Pvt_Ltd-Maharashtra-FY2025-26` | `GSTR-2B-Reco-ClientB_PRIVATE_LIMITED-Maharashtra-FY2025-26` |

Two spellings of one client in one state (`ClientB PRIVATE LIMITED` with an April GSTR-3B, `ClientB Private Limited` with April and May; different contents), Fresh mode:

| | A | B |
|---|---|---|
| Client folders | both clients in `ClientB Pvt Ltd-MH` | both clients in `ClientB Pvt Ltd-MH` |
| April copies | both cleaned to `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx`: the second overwrites the first, which is kept as `…-Apr_backup_211026_1030.xlsx`; both reported "Success" | `GSTR3B-ClientB_PRIVATE_LIMITED-Maharashtra-Apr.xlsx` and `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx` side by side |
| ITC report names | identical for both clients (`ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx`), so the second client's report replaces the first's (the A audit's verifier ran this: `run/VERIFY-A-S45/scripts/v_abbrev_merge.py`; this section's `collide_probe.py` RAN only the names, which are identical, see the row above) | different names, both kept |

**What it changes in the annual working.** The names of copied files, report files and bracketed folder suffixes for clients written in capitals (common for company names) or containing words like "Privateer". Within each project the Links file-name cells match that project's copy names for every file copied successfully in the run (a skipped or failed file keeps its original name, rule A-R48 / B-R57, the same in both), so the templates find the copied files in both (RAN, the B8 row above). In B the same client is "Pvt Ltd" in its client folder but "PRIVATE_LIMITED" in its file and report names, and such names are 8 characters longer each time they appear (the A commits cite Excel's ~218-character path limit ⚠). In A, two spellings of one client in one state merge into the same file and report names and overwrite each other; in B only the client folder is shared.

**When.** B keeps the original text replacement from `edac6a3` (shared; `47a1492:utils/helpers.py:78-79`). A changed it in `33d0baf` (23-07-2026, "Fix sanitize abbreviation, processing re-entrancy, folder-setting key"); the message says the old "case-sensitive substring replace mangled words that merely CONTAIN them ('Privateer'->'Pvter') and missed ALL-CAPS names". Because "_" counts as part of a word, `33d0baf` stopped abbreviating underscore-joined names; `4219ffb` (25-07-2026, "Fix sanitize_filename: abbreviate underscore-joined names again") restored that, citing "path length (Excel's ~218 char limit, and there is no full-path guard anywhere)". A's tests pin the new behaviour (A:tests/test_helpers.py:43-48, 69-74); three of them fail against B's code (word boundaries, ALL-CAPS, and the run-together name `ABCPrivateLimited`, which B turns into `ABCPvtLtd`; RAN again in `run/VERIFY-C5a/Btests`; `notes2/A_tests.md` counts them in its "7 failures"). Not a GST rule.

### Category 3: State, GSTIN and PAN codes

The state part of each file name is turned into a two-letter code that goes into the client key, the client folder name and the report names.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R21 | B-R21 | State name → two-letter code table ⚠ [GST RULE]: 39 lower-case names, 36 codes (e.g. delhi DL, maharashtra MH, tamil nadu TN, odisha/orissa OD, puducherry/pondicherry PY); text trimmed and lower-cased first; letter codes only, no numeric state codes | A:utils/constants.py:251-292; look-up A:utils/helpers.py:592-606 | B:utils/constants.py:234-275; look-up B:utils/helpers.py:485-499 | same | RAN core_probe: 39 names and 36 codes in both; `Delhi`, `" delhi "`, `DL`, `Tamil Nadu`, `Orissa`, `National Capital Territory of Delhi`, `Ladakh` give DL, DL, DL, TN, OD, DL, LA in both (`DL` is not in the table; it gives DL only through the made-up-code rule of A-R22, which upper-cases its first three characters). funcdiff: `get_state_code` IDENTICAL. Since `edac6a3` (shared); B changed no state value after the split. |
| A-R22 | B-R22 | Made-up code for a state not in the table ⚠ [GST RULE] (A's registry does not mark ⚠; B's does): two or more words → first letters of the first two words; one word → first three characters of the original text, upper-cased | A:utils/helpers.py:608-618 | B:utils/helpers.py:501-511 | same | RAN core_probe: `Unknown Place` → `UP` (the same code as Uttar Pradesh), `Tamilnadu` → `TAM`, `Daman and Diu` → `DA`, `27` → `27`, `" Xyz"` → `" XY"`, `''` → `''`, in both. |
| A-R23 | B-R21 (use sites) | Where the code and where the state text are used [GST RULE]: the CODE in the client key, client folder and report names; the state TEXT as written in the organisation text report, the Step 2 table and the summary workbook | A:core/file_parser.py:229-236; A:utils/helpers.py:628-631; A:gui/handlers/processing_handler.py:293; A:core/file_organizer.py:82, 330 | B:core/file_parser.py:261-268; B:utils/helpers.py:521; B:gui/controllers/processing_controller.py:113; B:core/file_organizer.py:78, 313 | same | READ; RAN step3_probe: report names `…_MH_…`, `…_TN_…`, summary "State" column `Maharashtra`, `Tamil Nadu` in both. |
| A-R24 | B-R21 (note) | No GSTIN or PAN is read, checked or derived [GST RULE] | none | none | same | RAN: a whole-word, any-case search for "gstin" or "pan" over every `.py` file of both copies (tests excluded) finds nothing. |

There are no differences in this category.

### Category 4: Financial year and tax period

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R25 | B-R23 | Period parts of the file names ⚠ [GST RULE]: the period (GSTR-2B Reco, Sales Reco), the 8-digit date (IMS Reco), the month (GSTR-3B), start and end month (Sales) and the year (Annual Report) are kept as plain text and never read again; no financial-year or tax-period check exists | stored at A:core/file_parser.py:97-106 | stored at B:core/file_parser.py:93-102 | same | RAN: a search for `metadata` finds only the two lines that create and fill it in each project (A:core/file_parser.py:74, 106; B:core/file_parser.py:74, 102), apart from A's unrelated helper `_passes_metadata_checks` and comments in A:utils/helpers.py and A's tests, which read it only to check the parser. RAN core_probe: periods such as `Apr25-Mar26 (1)`, `Apr (1)`, `2025-26 (1)` are kept as text in both. The time stamps in folder names are clock times, not tax periods (Category 6). |

There are no differences in this category. In both projects, files of different years for one client are grouped together; which of several files of one type is linked is decided by alphabetical order (rule A-R57 / B-R56, Category 8).

### Category 5: Expected files and completeness

This category decides which return files every client is expected to have and how complete each client is shown to be on Step 2.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R26 | B-R24 | Six expected types for every client ⚠ [GST RULE]: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report | A:utils/constants.py:68-75; A:core/file_parser.py:252-260 | B:utils/constants.py:68-75; B:core/file_parser.py:280-288 | same | funcdiff: `_analyze_client_completeness` IDENTICAL; RAN core_probe: `0042-MH` missing GSTR-2B Reco, GSTR-3B Export, IMS Reco, Sales, Sales Reco in both. |
| A-R27 | B-R25 | Duplicates ⚠ [GST RULE]: GSTR-3B may have any number of files ("monthly filings - unlimited"); any other type with 2 or more files gets "Multiple <type> (<n> files)" | A:core/file_parser.py:262-271 | B:core/file_parser.py:290-299 | same | RAN core_probe: ClientD with two GSTR-2B Reco files → "Multiple GSTR-2B Reco (2 files)" in both. (The IMS "(1)" file of 1.1 makes a duplicate in A only.) |
| A-R28 | B-R26 | Completeness = types found ÷ 6 × 100, one decimal [GST RULE] | A:core/file_parser.py:273-275 | B:core/file_parser.py:301-303 | same | RAN core_probe: 100.0, 33.3, 16.7 in both. |
| A-R29 | B-R27 | Client status and how Step 2 shows it [GST RULE]: "Complete" / "Missing <n> files" (wins over duplicates) / "Has duplicates" | Status A:core/file_parser.py:277-286; display A:gui/handlers/client_handler.py:37-59 (row tint), 147-159 ("Complete Only") | Status B:core/file_parser.py:305-314; display B:gui/views/validation_view.py:387-410 (text colour), 441-447 ("Complete Only") | **different** (display only) | 5.1 below (RAN) |
| A-R30 | B-R28 | "Critical" types ⚠ [GST RULE] `['Annual Report', 'GSTR-3B Export']` in `validate_results`, which nothing calls | A:core/file_parser.py:338-371 (list :359) | B:core/file_parser.py:366-399 (list :387) | same (dead in both) | funcdiff: `validate_results` IDENTICAL; RAN: a search of both copies finds no caller. |
| A-R31 | B-R29 | Scan statistics and how they are shown | Statistics A:core/file_parser.py:308-336; Step 2 text box A:gui/main_window.py:294-328 | Statistics B:core/file_parser.py:336-364; Step 1 scan log, Step 2 summary and status bar B:gui/pyqt_main_window.py:247-312 | **different** (input and display) | 5.2 below (RAN) |

#### 5.1 How Step 2 shows a client's status (A-R29 / B-R27) · [GST RULE]

**What both do.** The status texts and the order of the tests are identical (funcdiff IDENTICAL). "Complete Only" ticks only clients whose status is "Complete".
**What A does.** Shows the status with an icon ("✅ Complete", "⚠️ Missing 5 files", "⚠️ Has duplicates"; A:gui/handlers/client_handler.py:39) and tints the whole row: light green `#E8F5E8` for Complete, light orange `#FFF4E5` for every other status (A:gui/handlers/client_handler.py:52-59; colours A:utils/constants.py:154-155). "Complete Only" tests "status contains Complete" (A:gui/handlers/client_handler.py:152).
**What B does.** Shows the status without the icon and colours the text of the row: green `#107C10` for Complete, orange `#FF8C00` when the status contains "Missing", red-orange `#D83B01` otherwise (B:gui/views/validation_view.py:387-398). "Complete Only" tests "status equals Complete" (B:gui/views/validation_view.py:446). A count label reads "3 clients (1 complete)" (B:gui/views/validation_view.py:407-410).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`; the same three clients).**

| Client | Status (both) | A on Step 2 | B on Step 2 |
|---|---|---|---|
| `ClientA` (all six types) | Complete | "✅ Complete", row tinted `#E8F5E8` | "Complete", text `#107C10` |
| `0042` (Annual Report only) | Missing 5 files | "⚠️ Missing 5 files", row tinted `#FFF4E5` | "Missing 5 files", text `#FF8C00` |
| `ClientE` (all six types, two GSTR-2B Reco) | Has duplicates | "⚠️ Has duplicates", row tinted `#FFF4E5` (same as a client with missing files) | "Has duplicates", text `#D83B01` |
| "Complete Only" | | ticks ClientA only | ticks ClientA only |

Neighbouring screen differences seen in the same runs (registered under category 16 as A-R138 / B-R134): A's rows start unticked and the "Missing" column lists the missing type names; B's rows start all ticked and "Missing" shows a count.

**What it changes.** Only how quickly staff can tell a client with duplicate files (for example two GSTR-2B reconciliations for the year) from a client with missing files. Which clients are complete, and which are ticked by "Complete Only", is the same.

**When.** Status logic since `edac6a3` (shared). A's tints are the shared Tk code (colours in `GUI_CONFIG`, the same values in both projects' `utils/constants.py`; B no longer reads them). B's colours came with `c258047` (07-02-2026) and `7e09e5c` (08-02-2026); no message explains the choice. Tagged [GST RULE] because the status rule is tagged; the difference is display only, and nothing here judges the rule.

#### 5.2 Scan statistics and where they are shown (A-R31 / B-R29)

**What both do.** The same statistics function (funcdiff `get_statistics` IDENTICAL).
**What differs.** The input: B's "variations" include the files rejected as too small, not Excel or unreadable (1.3 to 1.5; an entry whose details cannot be read at all, 1.5, is only logged and not counted), which A never sees, while B's "total files" counts only the files that passed those checks. The display: A writes one text box on Step 2 (A:gui/main_window.py:294-328); B writes a log on Step 1, an HTML summary (formatted text with small tables) on Step 2 and a status-bar line (B:gui/pyqt_main_window.py:250-269, 281-298, 308-312).

**Worked example (RAN, `display_probe.py`, the `src_discovery` folder of 1.2).**

| Shown to the user | A | B |
|---|---|---|
| Files found | "Total Excel Files Found: 6" | "Files found: 4" (status bar: "Scan complete: 1 clients, 4 files, 7.8 KB") |
| Parsed | "✅ Successfully Parsed: 3"; "📊 Parsing Success Rate: 50.0%" | "Parsed: 2 (50.0%)" |
| Unparsed | "❌ Unparsed/Variations: 3" | "Unparsed: 2" |
| Variations | "🔍 FILE VARIATIONS: 3 files" | "Variations: 5" (on the same Step 1 log and Step 2 summary as "Unparsed: 2", a few lines further down) |
| Clients | "Total Clients Identified: 1"; "⚠️ Clients with Missing Files: 1"; "📈 Completeness Rate: 0.0%" | "Clients: 1"; "Incomplete: 1"; "Complete: 0 (0.0%)" |
| GSTR-3B files | "• GSTR-3B Export: 3 files" | "GSTR-3B Export: 2" |
| Size | "💾 TOTAL DATA SIZE: 11.7 KB" | "Total size: 7.8 KB" |

**What it changes.** Only the counts the staff read after a scan. Technical point (clarity of messages): B shows "Unparsed: 2" and "Variations: 5" for the same scan without saying why they differ; A labels every client that is not Complete (including duplicate-only clients) "Clients with Missing Files".

**When.** The statistics function is shared (`edac6a3`). A's text box is the shared Tk code. B's display came with `c258047` / `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking", which added the statistics to the scan result), and the extra variations with `7e09e5c` (08-02-2026). Not a GST rule.

### Category 6: Folder layout and naming

This category decides the folder tree Step 3 builds in the target folder: a run folder `Annual Statement-<DDMMYY HHMM>` (Level 1), one client folder `<client>-<state code>` per client (Level 2), a `Version-<DDMMYY HHMM>` folder for each run (Level 3) and three category folders inside it (Level 4); and which run and version Step 4 later works on.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R32 | B-R31 | Run folder name `Annual Statement-<DDMMYY HHMM>`; a second Fresh run in the same minute reuses it | A:utils/constants.py:82; A:core/file_organizer.py:163-196 (worked out once per run, A-R111) | B:utils/constants.py:82; B:core/file_organizer.py:146-179 (worked out per client, same result) | same | RAN core_probe and step3_probe: `Annual Statement-211026 1030` in both. |
| A-R33 | B-R30 | Time stamps: `%d%m%y %H%M` for folders, the same with `_` for report and summary names; two separate clock readings per run | A:utils/helpers.py:25-35; A:core/file_organizer.py:43; A:gui/handlers/processing_handler.py:186 | B:utils/helpers.py:23-33; B:core/file_organizer.py:43; B:gui/controllers/processing_controller.py:44 | same | funcdiff: `get_timestamp`, `get_safe_timestamp` IDENTICAL. These are clock times, not tax periods. |
| A-R34 | B-R30 | `get_date_only()` (`%d%m%y`) - never called | A:utils/helpers.py:29-31 | B:utils/helpers.py:27-29 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller (only the re-export list in `utils/__init__.py`) |
| A-R35 | B-R32 | Client folder name [GST RULE] (the state-code part): abbreviations (2.2), "-<state code>", cut to the maximum length; if fewer than 6 characters would remain for the client, the whole name is cut and the state code is lost | A:utils/helpers.py:620-644 | B:utils/helpers.py:513-537 | same | funcdiff: `create_client_state_key` IDENTICAL. RAN core_probe: `ClientB PRIVATE LIMITED`/Maharashtra/35 → `ClientB Pvt Ltd-MH`; `ClientA Private Limited`/Delhi/15 → `ClientA Pvt -DL`; maximum 8 → `ClientA ` (code lost) - in both. |
| A-R36 | B-R32, B-R138 | The maximum client-folder length on Step 1 (default 35; 0 or less becomes 35 inside the organiser, A:core/file_organizer.py:42 = B:core/file_organizer.py:42) [GST RULE] (tagged here only because B-R32, the client-folder name whose state-code part a small value removes, is tagged in B's registry; A's registry does not tag A-R36, B's does not tag B-R138) | Spin box (a number box with up and down arrows) with arrows 15-100, always active; typed values not checked. A:gui/tabs/setup_tab.py:434-443 | Qt spin box held within 15-100; greyed out unless "Include client name" is ticked, but its value is used anyway. B:gui/views/setup_view.py:374-382; B:gui/pyqt_main_window.py:203 | **different** | 6.1 below (RAN) |
| A-R37 | B-R33 | `FOLDER_STRUCTURE['level2'] = '{client}-{state}'` - never used | A:utils/constants.py:83 | B:utils/constants.py:83 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R38 | B-R34 | A new `Version-<DDMMYY HHMM>` folder in every run and mode | A:utils/constants.py:84; A:core/file_organizer.py:102-108 | B:utils/constants.py:84; B:core/file_organizer.py:98-104 | same | funcdiff: `create_client_structure` IDENTICAL; RAN modes_probe. |
| A-R39 | B-R35 | Names of the three category folders | `GSTR3B ({client})`, `ITC ({client})`, `Sales ({client})`. A:utils/constants.py:85-95 | `GSTR-3B Exports ({client})`, `Other ITC related files ({client})`, `Sales related files ({client})`. B:utils/constants.py:85-89 | **different** | 6.2 below (RAN) |
| A-R40 | B-R35, B-R36 | Whether a client's name goes into its category folders: the global tick for all clients, else a per-client flag | Decision A:core/file_organizer.py:121-134; per-client flags set on Step 2, A:gui/handlers/client_handler.py:175-223 | Decision B:core/file_organizer.py:117-130 (identical); nothing in B ever sets a per-client flag (dictionary defined at B:models/app_state.py:83 and only passed on) | **different** | 6.3 below (RAN) |
| A-R41 | B-R37 | Which folder each file type goes into ⚠ [GST RULE] (A's registry marks ⚠, B's does not) | GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder. A:core/file_organizer.py:302-315 | The same mapping. B:core/file_organizer.py:285-298 | **different** only in the folder names of 6.2 | 6.4 below (RAN) |
| A-R42 | B-R39 | Step 4: which run folder and which version folders are offered | Run folder: the `Annual Statement-*` entry with the newest date modified (when the system last recorded a change). In it: client folders not starting with `_`; per client only its newest `Version-*` folder by date modified. A:power_query_extractor/gui/extractor_window.py:611-668 | Run folder: the highest name (sorted Z to A). In it: every client folder (including names starting `_`) that has at least one sub-folder holding an `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx`; every such sub-folder (in practice the Version folders; the name is not checked), highest name first (Z to A). B:gui/pyqt_main_window.py:419-495 | **different** | 6.5 below (RAN) |
| A-R62 (A files it under category 8) | B-R40 | Organisation text report `_Organization_Report_<client>_<state>.txt` in the Version folder, with whole-run running totals | A:core/file_organizer.py:317-379 | B:core/file_organizer.py:300-362 | same (only the folder names written inside differ, 6.2) | funcdiff: `create_organization_report` IDENTICAL. |
| (no A rule; the same unused method at A:core/file_organizer.py:381-396) | B-R41 | `cleanup_on_error` - never called, so a client that fails midway leaves its folders | A:core/file_organizer.py:381-396 (its helper `_is_folder_empty` :398-407) | B:core/file_organizer.py:364-379 (its helper `_is_folder_empty` :381-390) | same (dead in both) | funcdiff: `cleanup_on_error` IDENTICAL; RAN: a search of both copies finds no caller. |
| A-R43 (A files it under category 7) | B-R38 | Re-Run and Resume reuse the `Annual Statement-*` folder with the newest date modified | A:core/file_organizer.py:174-196 (newest at :184) | B:core/file_organizer.py:158-179 (newest at :167) | same | RAN modes_probe: with `Annual Statement-311225 1000` (modified 31-12-2025) and `Annual Statement-211026 0335` (modified 21-10-2026), a Re-run used `Annual Statement-211026 0335` in both. Note: B's own Step 4 would show the other folder (6.5). |

#### 6.1 The maximum client-folder length box (A-R36 / B-R32, B-R138) · [GST RULE]

**What A does.** A Tk spin box (a number box with up and down arrows) whose arrows run from 15 to 100 (A:gui/tabs/setup_tab.py:440-442), always active. A value typed into it is not checked: a typed value goes to Step 3 as it is (A:gui/handlers/processing_handler.py:194). An arrow click first brings an out-of-range value back into 15-100 (0 or 8 become 15, 500 becomes 100; RAN `scripts2/spin_probe.py` with the same box settings under a virtual screen), so the code's own "0 or less becomes 35" step after an arrow click (A:gui/tabs/setup_tab.py:434-438) never acts.
**What B does.** A Qt spin box with range 15-100 (B:gui/views/setup_view.py:374-377), which cannot hold a value outside the range. It is greyed out until "Include client name in brackets for Level 4 folders" is ticked (B:gui/views/setup_view.py:377, 382), but Start reads its value whether greyed or not (B:gui/pyqt_main_window.py:203).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`, `core_probe.py`).**

| Typed or stored value | A reads | B holds |
|---|---|---|
| `8` | 8 | 15 |
| `500` | 500 | 100 |
| `20.5` | 20 | (cannot be typed) |
| `abc` | reading fails (Tcl error, an error from the layer underneath Tk) | ignored, keeps the previous value (20) |
| a saved 0 brought back at start-up | 35 (READ, A:gui/main_window.py:386-387) | shown as 15 |
| box enabled at start | yes | no (greyed) |

Folder names with a maximum of 8 (only possible in A): `ClientA`/Delhi → `ClientA-`; `ClientA`/Maharashtra → `ClientA-` (the same folder for both states); `ClientB PRIVATE LIMITED` → `ClientB ` (ending in a space); `ClientD`/Tamil Nadu → `ClientD-`. With B's lowest value 15: `ClientA Private Limited`/Maharashtra → `ClientA Pvt -MH`. (RAN, the builder is identical in both, rule A-R35.)

**What it changes.** In A, a typed value of 8 or less removes the state code from every client folder name "<client>-<code>" that is longer than the value (9 or less for a three-letter made-up code of A-R22; a short name such as `0042-MH` with a maximum of 8 keeps its code, RAN core_probe), so one client's registrations in different states would share a client folder; a typed text such as `abc` makes the Step 3 run stop with "Fatal error" (READ: the value is read at A:gui/handlers/processing_handler.py:194, inside the run's outer protection only; the A audit RAN this and also found that every settings save then fails). In B the value always stays within 15-100. Technical point: B's box cannot produce the values that drop the state code; A's can. B's greyed-out look suggests the value is not used, while it is.

**When.** The arrows 15-100 are the shared Tk code since `edac6a3` (`47a1492:gui/tabs/setup_tab.py:271-272`). At the split the shared box also had a typing check that refused anything that was not a whole number from 10 to 100 (an empty box was allowed) (`47a1492:gui/tabs/setup_tab.py:252-260, 269, 274`, from `edac6a3`). A removed that check in `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"); the message does not mention it (the A writers listed it as "UNREGISTERED" in `A_cross_flags.md`). So "typed values unchecked" is A's own change after the split, not the shared code (READ). B's PyQt5 box came in `c258047` (07-02-2026) with range 5-100 and the greyed-out state; `7e09e5c` (08-02-2026) set the range to 15-100. No message gives a reason. Tagged [GST RULE] only through B-R32 (a small value removes the state code from the folder name; the folder-name rule itself, A-R35, is the same in both); the box is described only.

#### 6.2 Category folder names (A-R39 / B-R35)

**What A does.** `GSTR3B`, `ITC`, `Sales`, each followed by " (<cleaned client name>)" when the client's name is to be included (A:utils/constants.py:91-95).
**What B does.** `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`, with the same optional suffix (B:utils/constants.py:85-89).
**Same in both.** All three are created even when empty; the suffix is removed by text replacement when not wanted (A:core/file_organizer.py:117-139 = B:core/file_organizer.py:113-135).

**Worked example (RAN, `core_probe.py`, Fresh run at 21-10-2026 10:30).**

| | A | B |
|---|---|---|
| ClientA's GSTR-2B Reco copy | `…\ClientA-DL\Version-211026 1030\ITC\GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` | `…\ClientA-DL\Version-211026 1030\Other ITC related files\GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` |
| ClientA's GSTR-3B copy | `…\Version-211026 1030\GSTR3B\GSTR3B-ClientA-Delhi-Apr.xlsx` | `…\Version-211026 1030\GSTR-3B Exports\GSTR3B-ClientA-Delhi-Apr.xlsx` |
| ClientA's Sales copy | `…\Version-211026 1030\Sales\Sales-ClientA-Delhi-Apr25-Mar26.xlsx` | `…\Version-211026 1030\Sales related files\Sales-ClientA-Delhi-Apr25-Mar26.xlsx` |
| With the include-name tick | `ITC (ClientA)`, `GSTR3B (ClientA)`, `Sales (ClientA)` | `Other ITC related files (ClientA)`, `GSTR-3B Exports (ClientA)`, `Sales related files (ClientA)` |
| Path length | A's paths are 20 (ITC), 14 (Sales) and 9 (GSTR-3B) characters shorter | |

**What it changes.** Every path below a Version folder, and therefore the folder text written into the template Links cells (8.3). Staff opening a client's working folder see different folder names; folders made by older runs keep their old names. Excel and Power Query fail on full paths above about 218 characters ⚠ (Excel behaviour, may change), so the shorter names leave more room for long client names and share paths.

**When.** B's names are the shared originals (`edac6a3`; `47a1492:utils/constants.py:86-88`). A shortened them in `7347383` (26-07-2026, "Shorten category folder names to fit Excel's path limit"). The message reports a real run of 3030 output files on the network share whose worst path was 270 characters, with 1094 files (36 %) over Excel's ~218 limit; it models the short names together with a planned shorter share root at a worst case of 215 characters with none over, against 235 characters and 27 files over for the shorter root alone, and warns that "the margin is thin" because copied file names still carry the full client name. The code comment that accompanies the change (A:utils/constants.py:85-90) calls the 235 / 27 and 215 / 0 figures "Measured on a real run" and leaves out that they assume the shorter share root. Not a GST rule in either registry; its effect on the Links cells is 8.3.

#### 6.3 Per-client choice to include the client's name (A-R40 / B-R35, B-R36)

**What both do in the organiser.** If the global tick is on, every client's category folders carry the name; otherwise a per-client flag looked up under "<client>-<state code>" decides, default off (A:core/file_organizer.py:121-134 = B:core/file_organizer.py:117-130).
**What A does on screen.** Step 2 has a "Name" column; clicking it toggles the client's flag (A:gui/handlers/client_handler.py:175-209), with a warning for names over 10 characters; the global tick sets all flags (A:gui/handlers/client_handler.py:211-223).
**What B does on screen.** Nothing sets a per-client flag: the dictionary is created empty (B:models/app_state.py:83) and only passed to the organiser (B:gui/pyqt_main_window.py:370; B:gui/controllers/processing_controller.py:27-32, 50). Step 1 still says "Names >10 chars will show warning. This overrides individual client settings if checked." (B:gui/views/setup_view.py:385).

**Worked example (RAN, `core_probe.py`, `gui_probe_A.py`).** With the flag `{'ClientA-DL': True}` handed to the organiser, both projects name only ClientA's folders `… (ClientA)` and leave the other clients' plain (identical result). On A's Step 2, clicking ClientA's "Name" cell produced exactly that flag; in B no screen action can produce it (search RAN: no line writes the dictionary).

**What it changes.** In B the client's name can only be added to the category folders of all clients or none; in A it can be added for single clients (for example only for the two clients whose folders staff open often).

**When.** The per-client control is shared Tk code, but at the split it stored flags under the state NAME while the organiser looked them up under the state CODE, so per-client choices were silently ignored in both lines of work (`47a1492:gui/handlers/client_handler.py:178, 214`). A fixed the key in `33d0baf` (23-07-2026; the message says "per-client folder-name toggles were silently ignored"). B dropped the control with the Tk screens (`c258047`, `a62d3cc`, 07-02-2026). Not a GST rule.

#### 6.4 Which folder each file type goes into (A-R41 / B-R37) · [GST RULE] · ⚠

**What both do.** The same fixed map (funcdiff `_get_destination_folder` IDENTICAL): GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder itself; any other type → "No destination folder for type: <type>". Only the folder names differ (6.2).

**Worked example (RAN, `core_probe.py`).** Every one of ClientA's files (eight in A, seven in B, because B does not recognise the IMS " (1)" copy, 1.1) landed in the same category in both projects; for example `GSTR3B-ClientA-Delhi-Apr.xlsx` → A `…\Version-211026 1030\GSTR3B\`, B `…\Version-211026 1030\GSTR-3B Exports\`; `AnnualReport-ClientA-Delhi-2025-26.xlsx` → `…\Version-211026 1030\` in both.

**What it changes.** Nothing about which return goes with which other return; only the folder names that the Links cells point to (8.3).

**When.** The map since `edac6a3` (shared); the folder names as 6.2 (`7347383` in A). Tagged because the registry of A tags it (it would need changing if the set of return types changes ⚠); nothing here judges it.

#### 6.5 Step 4: which run folder and which version folders are offered (A-R42 / B-R39)

**What A does.** In the folder shown on Step 4, every entry named `Annual Statement-*` is a candidate and the one with the newest date modified is used. Inside it, every sub-folder not starting with `_` is a client; its newest `Version-*` folder by date modified is used; a client is listed if it has any `Version-*` folder, with the ITC and Sales ticks set only if such a report is present (names starting `itc_report_` / `sales_report_` and ending `.xlsx`, any letter case). A:power_query_extractor/gui/extractor_window.py:20-42, 611-668.
**What B does.** It reads the target folder last saved in the settings (B:gui/pyqt_main_window.py:421), sorts the `Annual Statement-*` folders by name from Z to A and uses the first (B:gui/pyqt_main_window.py:432-448). Inside it, it goes through every client folder (names starting `_` included) and, in each, every sub-folder that holds a file matching `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx` (the sub-folder's name is not checked; in practice these are the Version folders), sorted by name Z to A; a client folder with no such sub-folder is left out (B:gui/pyqt_main_window.py:453-495).

**Worked example (RAN, `step4_probe.py` and `step4_probe_newonly.py`; dates modified set on purpose).** Target folder:
- `Annual Statement-311225 1000` (run of 31-12-2025, last modified 31-12-2025), with ClientA's ITC and Sales reports;
- `Annual Statement-211026 0335` (run of 21-10-2026, last modified 02-11-2026) holding: `ClientA-DL` with `Version-211026 0335` and a re-run `Version-021126 1000` (both with reports); `ClientB Pvt Ltd-MH` whose only report was renamed by hand to `….XLSX`; `_ClientC-DL` with an ITC report; `ClientD-TN` whose Version folder holds no report.

| | A | B |
|---|---|---|
| Run folder used | `Annual Statement-211026 0335` (newest date modified) | `Annual Statement-311225 1000` (highest name: "31…" sorts after "21…"), so last December's run is offered |
| If only the October 2026 run exists: ClientA | one line: `Version-021126 1000` (newest by date modified), ITC and Sales ticked | two lines: `Version-211026 0335` first, then `Version-021126 1000` (name order, day first) |
| `_ClientC-DL` | left out | listed |
| `ClientD-TN` (Version folder without reports) | listed, both ticks off | left out |
| `ClientB Pvt Ltd-MH` (report named `.XLSX`) | listed with ITC ticked; the search for the original report then finds nothing on Linux ("No ITC report found"); on Windows it is found (HIGH, not run) | left out on Linux; on Windows listed (HIGH) |

**What it changes.** Which run's reports Step 4 refreshes and extracts, and therefore which run's figures go into the consolidated workbook. Day-first names do not sort by date, so B's choice of run folder is the run whose name has the highest day number (ties broken by month, then year), which is not the latest run whenever an earlier run fell on a later day of its month (as here: 31-12-2025 beats 21-10-2026); inside a run B offers every version, in the same day-first name order (a re-run and its predecessor side by side; here the older `Version-211026 0335` comes first because 21 sorts after 02). A's choice follows the last change on disk. A folder's date modified changes only when something is written directly inside it (RAN on Linux in `run/VERIFY-C5a/tmp/mt`: a lock file written in a Version folder changed that folder's date only, not its client or run folder's). So opening an older run's summary workbook `GST_Processing_Summary_…xlsx`, which sits directly in the run folder and makes Excel write a lock file next to it, can make that older run "newest"; opening a report inside an older Version folder can make that Version folder the "newest" of its client (the A audit ran this on Linux; that Windows updates folder dates the same way is MEDIUM). Within B, Step 3's Re-Run picks the run folder by date modified (row A-R43 / B-R38) while Step 4 picks it by name, so the two steps of B can work on different runs. Technical point: both methods can pick an older run; neither shows the user which date it used.

**When.** At the split the Tk extractor chose the run folder by date modified and the version by name (`47a1492:power_query_extractor/gui/extractor_window.py:517, 529`). A changed the version choice to date modified in `c1583d6` (23-07-2026, "Fix extractor cache path and version-folder selection"; the message says names "are 'Version-DDMMYY HHMM' and day-first text does not sort chronologically across month boundaries") and replaced the report searches with the prefix test in `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor", for speed). B wrote its own Step 4 list in `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking"), choosing both by name; no reason is given. Not tagged as a GST rule in either registry; it decides which data set is used, not a tax rule.

### Category 7: Copying and versioning of files

This category decides how source files are copied into the tree, what happens when a copy already exists, and what the three processing modes do.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R43 | B-R44 | The three processing modes (fresh, rerun, resume) and how they are described on screen | Logic A:core/file_organizer.py:60-63, 163-196, 276-287. Screen A:gui/tabs/setup_tab.py:401-412 (radio buttons) and 103-110 (hover help) | Logic B:core/file_organizer.py:56-59, 146-179, 259-270. Screen B:gui/views/setup_view.py:343-359 (radio buttons), 510-512 (help card) | **different** (screen texts only; logic the same) | 7.1 below (RAN) |
| A-R44 | B-R45 | `PROCESSING_MODES` name, description and icon | Name and description shown on Step 1 (A:gui/tabs/setup_tab.py:404, 410); icon never read (A:utils/constants.py:178, 183, 188) | None of the three is read (B:utils/constants.py:157-173) | **different** | 7.1 below (RAN) |
| A-R45 | B-R43 | A destination file that already exists: Fresh → backup then overwrite; Resume → "Skipped", "File already exists"; Re-run → overwrite without backup | A:core/file_organizer.py:276-287 | B:core/file_organizer.py:259-270 | same | funcdiff: `_organize_single_file` IDENTICAL; RAN modes_probe (7.1). |
| A-R46 | B-R43 | Backup name `<stem>_backup_<DDMMYY_HHMM><ext>` in the same folder | A:utils/helpers.py:217-234 | B:utils/helpers.py:187-204 | same | funcdiff IDENTICAL; RAN modes_probe: `GSTR3B-ClientA-Delhi-Apr_backup_211026_1030.xlsx` in both. |
| A-R47 | B-R42 | Copy only (never move or delete the source), sizes compared, a partial copy left on error | A:utils/helpers.py:172-215 | B:utils/helpers.py:142-185 | same | funcdiff: `safe_copy_file` IDENTICAL. |
| A-R48 | B-R57 (B files it under category 8; [GST RULE] in B's registry) | After a successful copy the client's file record takes the cleaned copy name (used by the Links cells); skipped or failed files keep the original name [GST RULE] (tagged because B's registry tags B-R57; A's does not tag A-R48) | A:core/file_organizer.py:233-246, 270-274 | B:core/file_organizer.py:216-229, 253-257 | same | funcdiff: `organize_files` IDENTICAL. |
| A-R49 | B-R46 | `get_unique_path` - never called | A:utils/helpers.py:451-470 | B:utils/helpers.py:344-363 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller (only the re-export list in `utils/__init__.py`) |
| A-R50 | B-R47 | `calculate_file_hash` - never called | A:utils/helpers.py:529-547 | B:utils/helpers.py:422-440 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller (only the re-export list in `utils/__init__.py`) |

#### 7.1 How the processing modes are described on screen (A-R43, A-R44 / B-R44, B-R45)

**What both do.** The same logic (funcdiff: `_organize_single_file`, `organize_files`, `create_client_structure` IDENTICAL; the run-folder choice differs only in being worked out once per run in A). Fresh makes a new run folder; Re-Run and Resume reuse the run folder with the newest date modified; every mode makes a new Version folder named by the current minute, so Resume skips files only when it is started in the same minute as the earlier run. A mode text other than exactly `fresh`, `rerun` or `resume` is refused.

**Worked example (RAN, `modes_probe.py`).** ClientA's two files, four runs in the same minute:

| Run | A | B |
|---|---|---|
| 1 Fresh | both files "Success" | the same |
| 2 Resume | both "Skipped", "File already exists" | the same |
| 3 Re-run | both "Success" (overwritten, no backup) | the same |
| 4 Fresh | both "Success", with `…_backup_211026_1030.xlsx` copies beside them | the same (in `GSTR-3B Exports` instead of `GSTR3B`) |
| Mode text `Fresh` (capital F) | "Invalid processing mode: Fresh" | the same |

**What differs: the texts on Step 1 (RAN, `gui_probe_A.py`, `gui_probe_B.py`; A's hover help, a tip shown when the mouse rests on a button, was not captured by the probe and is READ at A:gui/tabs/setup_tab.py:103-110, 409).**

| Mode | A (radio button + line below + hover help) | B (radio button) | B (help card on the same page) |
|---|---|---|---|
| Fresh | "🆕 Fresh Run" / "Create new folder structure (recommended for first time)" / "Create a brand new timestamped folder for this run. Use this normally - previous runs are left alone." | "Fresh - Create new folder structure (recommended for first time)" | "Fresh: New timestamped structure" |
| Re-Run | "🔄 Re-Run" / "Add new version folders to existing structure" / "Add a new version inside the most recent run's folder instead of starting a new one. Use when redoing work for the same batch." | "Re-Run - Add new version folders to existing structure" | "Re-Run: Overwrite existing" |
| Resume | "▶️ Resume" / "Continue previously interrupted processing" / "Continue a run that was interrupted, skipping files that are already in place." | "Resume - Continue previously interrupted processing" | "Resume: Skip already processed" |

**What it changes.** Only what staff expect when they pick a mode, for example when a run broke off half-way through the client list. Technical point (clarity of messages): B's help card says "Re-Run: Overwrite existing", while its own button and the code add a new Version folder; no text in either project says that Resume skips files only within the same minute.

**When.** A shows the constant's name and description since the shared Tk code (`47a1492:gui/tabs/setup_tab.py:212-221`, from `edac6a3`); A added the hover help in `322717f` (25-07-2026). B's first PyQt5 texts (`c258047`, 07-02-2026) were "Re-Run - Overwrite existing files" and "Resume - Skip already processed"; the present button and card texts came in `7e09e5c` (08-02-2026). Not a GST rule.

### Category 8: Report creation and template Links cells

This category decides how Step 3 makes each client's ITC report and Sales report: what must be chosen first, which template is copied, and what is written into the template's "Links" sheet (the sheet into which the program writes folder paths and file names for the template's Power Query queries to read; Power Query is Excel's data-import tool, whose queries live inside the workbook). It also covers the Step 3 summary workbook. The program never writes a tax figure into a report; the figures are produced later by the template's own queries.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R51 | B-R48, B-R133 | What must be chosen before Dry Run and Start [GST RULE] (A's registry tags it; B's does not) | Dry Run and Start: both templates set and existing, and a target typed. A:gui/handlers/file_handler.py:132-151; called at A:gui/handlers/processing_handler.py:40 (Dry Run) and :78 (Start) | Start: the same three checks with other wording, B:gui/pyqt_main_window.py:345-361, and again per client, B:gui/controllers/processing_controller.py:210-218. Dry Run: no check, B:gui/pyqt_main_window.py:326-343 | **different** (Dry Run) | 8.1 below (RAN) |
| A-R52 | B-R48, B-R51 | Which template file types are offered or accepted; the report is always named `.xlsx` [GST RULE] (tagged here only because B-R51, the report-name rule, is tagged in B's registry for the state code in the name; A-R52 and B-R48 are not tagged) | Browse filter `*.xltx *.xlsx`, A:gui/handlers/file_handler.py:31, 42; Step 1 checklist accepts `.xlsx .xltx .xltm .xlsm`, A:gui/utils/setup_validation.py:21, 62-63; Start accepts any existing file | Browse filter `*.xlsx *.xlsm` and the label "(.xlsx or .xlsm)", B:gui/views/setup_view.py:221, 612, 619; Start accepts any existing file | **different** | 8.2 below (RAN) |
| A-R53 | B-R49 | `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` - never used | A:utils/constants.py:197 | B:utils/constants.py:180 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R54 | B-R54 | ITC report Links cells ⚠ [GST RULE]: B2 GSTR-3B folder (always); B4/B5 Version folder and Annual Report name; B7/B8 ITC folder and GSTR-2B Reco name; B10/B11 ITC folder and IMS Reco name; B3, B6, B9 never written | Map A:utils/constants.py:102-114; values A:core/excel_handler.py:603-656 | Map B:utils/constants.py:96-108; values B:core/excel_handler.py:540-588 | **different** (only the folder text, through 6.2; and B11 through 1.1) | 8.3 below (RAN) |
| A-R55 | B-R55 | Sales report Links cells ⚠ [GST RULE]: B2/B3 Sales folder and Sales name; B5/B6 Version folder and Annual Report name; B8/B9 Sales folder and Sales Reco name; B4, B7 never written | Map A:utils/constants.py:115-125; values A:core/excel_handler.py:657-708 | Map B:utils/constants.py:109-119; values B:core/excel_handler.py:590-622 | **different** (only the folder text, through 6.2) | 8.3 below (RAN) |
| A-R56 | B-R56 | Form of the Links values [GST RULE]: full folder paths with `/` turned into `\`; file names without extension; written as text; an empty value is not written, so the template's own text stays | A:utils/helpers.py:418-422, 553-555; A:core/excel_handler.py:189-197 (Excel), 297-303 (openpyxl) | B:utils/helpers.py:311-315, 446-448; B:core/excel_handler.py:159-167 (Excel), 246-252 (openpyxl) | same | funcdiff: `clean_windows_path`, `extract_filename_without_extension` IDENTICAL; RAN core_probe: B3 keeps "PLACEHOLDER-B3" and the file names have no extension in both. |
| A-R57 | B-R56 | Which file of a type feeds a Links cell [GST RULE]: the first in sorted full-path order | Sort A:utils/helpers.py:361; first file A:core/excel_handler.py:629, 640, 651, 667, 683, 699 | Sort B:utils/helpers.py:262; first file B:core/excel_handler.py:561-562, 572-573, 583-584, 595-596, 606-607, 617-618 | same (the outcome for IMS "(1)" files differs through 1.1) | RAN core_probe: ClientD's two GSTR-2B files `…-FY2025-26 (1).xlsx` and `…-FY2025-26.xlsx` → B8 = `GSTR-2B-Reco-ClientD-Tamil_Nadu-FY2025-26_(1)` in both (a space sorts before a dot). |
| A-R58 | B-R53 | Finding the Links sheet ⚠ [GST RULE] (B's registry tags it; A's does not) | Excel path: `Links` in any letter case, else a sheet whose name contains "link", else a new empty `Links` sheet, and success either way, A:core/excel_handler.py:162-183. openpyxl path: the name `Links` in any letter case only (no "contains link" search, no new sheet; if absent, nothing is written and the report still counts as created); a Links cell inside a merged range (several cells joined into one) is skipped silently, A:core/excel_handler.py:285-303, 310 | Excel path identical, B:core/excel_handler.py:131-153. openpyxl path: the same any-case name search with the same "nothing written, still created" result when absent; a Links cell inside a merged range makes the report fail, B:core/excel_handler.py:233-252, 258 | same on the Excel path; same sheet search on the openpyxl path; **different** on the openpyxl path for a merged Links cell | 8.4 below (RAN; sheet names `LINKS`, `links` and `My Links` RAN in both with `scripts2/links_case_probe.py`: the first two are found and filled, the third is not used and the report still returns success) |
| A-R59 | B-R51, B-R61 | Report file names and how Step 4 recognises them [GST RULE] | `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx`, `Sales_Report_…`, A:gui/handlers/processing_handler.py:291-298, 355-356. Step 4 ticks by names starting `itc_report_` / `sales_report_` and ending `.xlsx` in any letter case (A:power_query_extractor/gui/extractor_window.py:20-42) but finds the original with the search `ITC_Report_*.xlsx` (A:power_query_extractor/core/report_processor.py:111-125) | The same names, B:gui/controllers/processing_controller.py:111-130, 220. Step 4 uses the search `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` for the list (B:gui/pyqt_main_window.py:461-462) and for the original (B:power_query_extractor/core/report_processor.py:72-89) | **different** (names only through 2.2; recognition) | 8.5 below (RAN) |
| A-R60 | B-R50, B-R60 | What one Step 3 run does per client [GST RULE] (A-R60 and B-R50 tagged): order folders → copies → ITC report → Sales report → text report; both reports tried for every client; a client is "Success" even if its reports fail | A:gui/handlers/processing_handler.py:182-453; template paths and per-client flags read from the screen during the run (:193, 301, 310, 332) | B:gui/controllers/processing_controller.py:42-181, 204-239; template paths and options fixed when Start is pressed (B:gui/pyqt_main_window.py:352, 363-371) | **different** (live reads; client order); success counting the same | 8.6 below (RAN) |
| A-R61 | B-R94, B-R95 (B files them under category 11) | Step 3 summary workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (sheets Summary, Client Status, File Mapping, Errors, Variations) | Variations sheet filled from the scan; a failure is shown as a log line. A:gui/handlers/processing_handler.py:209-226, 455-475; A:core/excel_handler.py:371-601 | Variations sheet always "No file variations found"; a failure shows nothing. B:gui/controllers/processing_controller.py:58-76, 183-193; B:core/excel_handler.py:308-538 | **different** | 8.7 below (RAN) |
| A-R62 | B-R40 | Organisation text report | see Category 6 | see Category 6 | same | funcdiff IDENTICAL |
| A-R63 | B-R58 | `get_relative_excel_path` - never called (Links values are always full paths) | A:utils/helpers.py:557-574 | B:utils/helpers.py:450-467 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller (only the re-export list in `utils/__init__.py`) |
| A-R48 | B-R57 | Copy names used in the Links cells [GST RULE] (B's registry tags B-R57) | see Category 7 | see Category 7 | same | funcdiff IDENTICAL |
| A-R98 (A files it under category 12) | B-R52 | Where the report is assembled: a local temporary copy, or the template copied straight to the final report name | Excel path: works on a copy in `%TEMP%\gst_excel_temp` and copies the finished report to the target once, A:core/excel_handler.py:147-155, 215 (removal of the temporary copy :254-260); openpyxl path the same, A:core/excel_handler.py:273-279, 306-308, 315-320 | Copies the template to the final report name first and fills it there, B:core/excel_handler.py:122-129 (Excel), 227-231 (openpyxl) | **different** | 8.8 below (RAN on the openpyxl path) |
| A-R99 (A files it under category 12) | B-R59 | A report counts as "created" if the output file exists | A:core/excel_handler.py:217-224 (Excel), 310 (openpyxl) | B:core/excel_handler.py:182-189 (Excel), 258 (openpyxl) | same | READ; RAN step3_probe: a Sales report with no Links cell written counted as created in both. |

#### 8.1 What must be chosen before a Dry Run (A-R51 / B-R48, B-R133) · [GST RULE]

**What A does.** Dry Run (a preview that lists what Start would do, without copying anything) and Start use the same check: both template paths set ("Please select both ITC and Sales template files"), a target typed ("Please select a target folder for organized files"), and both template files present ("ITC template file not found" / "Sales template file not found"). A:gui/handlers/file_handler.py:132-151.
**What B does.** Start checks the same three things with its own words ("Please select a valid ITC template." / "Please select a valid Sales template." / "Please select a target folder."; B:gui/pyqt_main_window.py:352-361). Dry Run checks only that a client is ticked (B:gui/pyqt_main_window.py:326-330).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`).** No templates chosen, three clients ticked, Dry Run pressed:

| | A | B |
|---|---|---|
| Result | error box "Please select both ITC and Sales template files"; no preview | the preview is printed: for each client "📄 Would organize: …" for every file, then "📊 Would create: ITC Report" and "💰 Would create: Sales Report", ending "🧪 ════════ DRY RUN COMPLETE ════════" and "💡 No files were actually moved or created" |

**What it changes.** Only the preview: in B a Dry Run can promise ITC and Sales reports that Start would then refuse to make until templates are chosen. Start behaves alike in both: neither checks the template's type or contents (8.2).

**When.** The shared Dry Run called the check (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`). B's PyQt5 Dry Run came without it in `c258047` (07-02-2026); no reason is given. Tagged [GST RULE] by A's registry (it gates whether returns are processed); described only.

#### 8.2 Template file types (A-R52 / B-R48, B-R51) · [GST RULE]

**What A does.** The Browse dialog offers `*.xltx *.xlsx`; the Step 1 checklist marks other types "Not an Excel template", but this only colours the readiness pill; Start accepts any existing file.
**What B does.** The Browse dialog offers `*.xlsx *.xlsm`; Start accepts any existing file; there is no checklist.
**Same in both.** The report is always named `.xlsx` whatever the template's type (A:gui/handlers/processing_handler.py:298, 356; B:gui/controllers/processing_controller.py:220), and the template's contents (Links sheet, queries) are not checked.

**Worked example (RAN, inline runs in `run/C5a` on `tmp/tplnames`).**

| ITC template chosen | A checklist | A Start | B Start |
|---|---|---|---|
| `ITC_template.xlsx` | ok | accepted | accepted |
| `ITC_template.xltx` | ok | accepted | accepted |
| `ITC_template.xlsm` | ok | accepted | accepted |
| `ITC_template.xltm` | ok | accepted | (not run; accepted by the same test) |
| `ITC_template.xlsb` | "Not an Excel template" | accepted | accepted |
| `ITC_template.txt` | "Not an Excel template" | accepted | accepted |
| none | "Not selected yet" | refused | refused ("Please select a valid ITC template.") |

**What it changes.** Which template files staff can pick with the mouse: an `.xltx` template in A, an `.xlsm` in B (both can be typed in either). An `.xlsm` or `.xltx` template gives a report named `.xlsx` whose inside is a macro workbook or a template; whether Excel then opens it normally is not sure (MEDIUM; how to check: run one client on Windows with such a template and open the report). Tagged [GST RULE] only through B-R51 (the report name, which carries the state code; the names themselves are the same in both, 8.5); the choice of template types is tagged in neither registry and is described only.

**When.** The shared Browse filter was `*.xltx *.xlsx` (`47a1492:gui/handlers/file_handler.py:31, 42`), kept by A; A added the checklist in `322717f` (25-07-2026). B's filter `*.xlsx *.xlsm` came in `c258047` (07-02-2026) and its label in `7e09e5c` (08-02-2026). No reason is given.

#### 8.3 The folder text in the Links cells (A-R54, A-R55 / B-R54, B-R55) · [GST RULE] · ⚠

**What both do.** The same cells ⚠ (template cell positions) from the same sources (funcdiff: `prepare_template_data` differs only in removed debug lines). What differs is the folder text, because of the category folder names (6.2), and for IMS the file chosen (1.1).

**Worked example (RAN, `core_probe.py`, a placeholder template whose cells B2-B11 hold "PLACEHOLDER-B…", openpyxl path; Linux paths appear with backslashes as `\root\…`, shortened here to `…`).** ClientA, Delhi:

| ITC report cell | A | B |
|---|---|---|
| B2 (GSTR-3B folder) | `…\Annual Statement-211026 1030\ClientA-DL\Version-211026 1030\GSTR3B` | `…\Annual Statement-211026 1030\ClientA-DL\Version-211026 1030\GSTR-3B Exports` |
| B3 | `PLACEHOLDER-B3` (template text kept) | `PLACEHOLDER-B3` |
| B4 (Version folder) | `…\ClientA-DL\Version-211026 1030` | the same |
| B5 (Annual Report) | `AnnualReport-ClientA-Delhi-2025-26` | the same |
| B7 (GSTR-2B folder) | `…\Version-211026 1030\ITC` | `…\Version-211026 1030\Other ITC related files` |
| B8 (GSTR-2B file) | `GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26` | the same |
| B10 (IMS folder) | `…\Version-211026 1030\ITC` | `…\Version-211026 1030\Other ITC related files` |
| B11 (IMS file) | `ImsReco-ClientA-Delhi-31032026_(1)` (1.1) | `ImsReco-ClientA-Delhi-31032026` |

| Sales report cell | A | B |
|---|---|---|
| B2 and B8 (Sales folder) | `…\Version-211026 1030\Sales` | `…\Version-211026 1030\Sales related files` |
| B3 (Sales file) | `Sales-ClientA-Delhi-Apr25-Mar26` | the same |
| B5 / B6 | `…\Version-211026 1030` / `AnnualReport-ClientA-Delhi-2025-26` | the same |
| B9 (Sales Reco file) | `SalesReco-ClientA-Delhi-FY2025-26` | the same |

For `ClientB PRIVATE LIMITED` the file-name cells differ as well, through 2.2 (B8 `GSTR-2B-Reco-ClientB_Pvt_Ltd-Maharashtra-FY2025-26` in A, `…ClientB_PRIVATE_LIMITED…` in B). For a client with only GSTR-3B and GSTR-2B files (ClientD), only ITC B2 and B7/B8 are written and the Sales report gets no cell, in both.

**What it changes.** The folder each template query is pointed at. If the templates' queries take the folder from these cells, both projects' reports find their files. If any query contains a fixed folder name such as "Other ITC related files", it works only with B's names (and with A's runs made before `7347383`). The templates are not in either repository, so this is NOT SURE. How to check: open each template in Excel, Data > Get Data > Launch Power Query Editor, and for each query open Advanced Editor; search for "GSTR-3B Exports", "Other ITC related files" and "Sales related files", and see whether the source step refers to the Links cells.

**When.** Cells and sources since `edac6a3` (shared; B changed no Links value after the split). A's folder text changed with the folder names in `7347383` (26-07-2026, "Shorten category folder names to fit Excel's path limit"; reasons in 6.2); the commit changed only `utils/constants.py`, and the A audit (Section 5, A-R54) asks the author the same template question. Which version to keep is the author's decision.

#### 8.4 A Links cell inside a merged cell range, openpyxl path (A-R58 / B-R53) · [GST RULE] · ⚠

**What both do on the Excel (COM) path** - the path used on office PCs with pywin32 - identical code (A:core/excel_handler.py:162-183 = B:core/excel_handler.py:131-153; READ only, Excel cannot run here).
**Same on the openpyxl path.** Both look only for a sheet named `Links` in any letter case; there is no "contains link" search and no new sheet, and if no such sheet exists nothing is written and the report still counts as created (A:core/excel_handler.py:289-310 = B:core/excel_handler.py:238-258; RAN `scripts2/links_case_probe.py`: `LINKS` and `links` filled in both, `My Links` left untouched with success in both).
**What differs on the openpyxl path.** Writing into a cell that lies inside a merged range (several cells joined into one on the sheet), but is not its first cell, raises an error in openpyxl. A catches every error for that cell and moves on (a bare `except: pass`, a catch-all that swallows any error without a trace, A:core/excel_handler.py:302-303). B catches only `KeyError` and `ValueError` (two named kinds of error, B:core/excel_handler.py:251); the merged-cell error is a different kind, so it ends the whole report.

**Worked example (RAN, `merged_probe.py`).**

| Template Links sheet | A | B |
|---|---|---|
| B2 is the first cell of a merged B2:C2 | B2 written, all other cells written, report "created" | the same |
| B2 lies inside a merged A2:B2 | B2 left empty, all other cells written (B4, B5, B11 …), report "created", no message | report fails; log file "Fallback method error: 'MergedCell' object attribute 'value' is read-only"; the unfilled template copy stays on disk under the report name, with B4, B5, B11 still showing the template's placeholders |

**What it changes.** Nothing for reports made through Excel (identical code). The openpyxl path runs in A on any PC where pywin32 cannot be loaded (A:core/excel_handler.py:20-28; normally off Windows, since A's start-up check requires pywin32 on Windows, rule A-R124); in B the window cannot start without pywin32 (rule B-R124), so B's openpyxl path cannot be reached through the app at all (HIGH, READ). The openpyxl path also loses the template's Power Query (rule A-R96 / B-R97, category 12).

**When.** Shared bare `except` from `edac6a3`. B narrowed it in `c258047` (07-02-2026); the message lists "Fixed all 26 bare except: clauses with specific exception types". Tagged [GST RULE] and ⚠ by B's registry (the sheet name `Links` belongs to the templates); described only.

#### 8.5 Report names and how Step 4 recognises them (A-R59 / B-R51, B-R61) · [GST RULE]

**What both do.** Name the reports `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` and `Sales_Report_…` in the client's Version folder; the cleaned client name differs through 2.2 (RAN step3_probe: `ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx` in A, `ITC_Report_ClientB_PRIVATE_LIMITED_MH_211026_1030.xlsx` in B; ClientA and ClientD identical). Refreshed copies (`…_Refreshed_…xlsx`) count as reports in both.
**What differs in Step 4.** A ticks a report when a file name, in lower case, starts `itc_report_` / `sales_report_` and ends `.xlsx`; it then looks for the original with the search `ITC_Report_*.xlsx`, which respects letter case on Linux. B uses that search for both the list and the original.

**Worked example (RAN, `step4_probe.py`, on Linux).** ClientB's only report renamed by hand to `ITC_Report_ClientB_Pvt_Ltd_MH_211026_0335.XLSX`:

| | A | B |
|---|---|---|
| Step 4 list | ClientB listed, ITC ticked | ClientB not listed |
| Processing that report | "No ITC report found" | (not offered) |

On Windows both find the file (HIGH, Python's search ignores letter case there; not run). The program itself always writes `.xlsx` in lower case, so only renamed files are affected.

**What it changes.** Only reports whose names were changed by hand; on the office's Windows PCs the two behave alike (HIGH). The name differences of 2.2 decide which file a staff member opens for `ClientB PRIVATE LIMITED`.

**When.** Names since `edac6a3` (shared). A's prefix test came in `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor"), for speed on the network share; its message calls the new helpers "proven equivalent to the original glob logic", which holds on Windows but not on Linux (RAN above). B's search came with its Step 4 page in `b0e7d0e` (09-02-2026). Described only.

#### 8.6 What one Step 3 run does per client (A-R60 / B-R50, B-R60) · [GST RULE]

**What both do.** For each ticked client: folders → copies → ITC report → Sales report → organisation text report; both reports are tried for every client whatever files it has; a client is counted "Success" unless an error escapes the client's steps, even if its reports failed; the look-up of the client's data sits outside the per-client protection (A:gui/handlers/processing_handler.py:232; B:gui/controllers/processing_controller.py:85).
**What differs (1): when the template paths are read.** A reads the ITC and Sales template paths and the per-client name flags from the Step 1/Step 2 screen while the run goes on (A:gui/handlers/processing_handler.py:193, 301, 310, 332). B collects all settings once when Start is pressed and hands that snapshot to the worker (B:gui/pyqt_main_window.py:352, 363-371; used at B:gui/controllers/processing_controller.py:119, 127).
**What differs (2): the order of clients.** A processes clients in the order of its Step 2 table, which is the order in which the scan first met each client in the alphabetically sorted file list (A:gui/handlers/client_handler.py:30, 165). B's table, and so its run, is sorted by client key (B:gui/views/validation_view.py:367), unless the user clicks a column heading: the table allows sorting by any column (B:gui/views/validation_view.py:324), and the run then follows the new order on screen, because the ticked keys are read row by row (B:gui/views/validation_view.py:412-421; READ, HIGH). (B's order is registered in B-R134, category 16: rows sorted by client key, a heading click re-sorts, processing follows the order on screen. A's order is in no registry rule - A-R60 only says the run follows the table order; the A writers flagged it as "UNREGISTERED" in `A_cross_flags.md`.)

**Worked example (RAN, `step3_probe.py`, `step3broken`, inline order run).**

| | A | B |
|---|---|---|
| Three clients (ClientA, ClientB PRIVATE LIMITED, ClientD); after ClientA is completed the ITC template on Step 1 is changed from "v1" to "v2" | ClientA's ITC report from v1; ClientB's and ClientD's from **v2** (Links A1 marker "ITC TEMPLATE v2") | all three ITC reports from **v1** |
| ClientD has only one GSTR-3B file | its Sales report is still made and counted "created" (no Links cell written) | the same |
| The ITC template is a damaged file | client "Success"; summary: Successful 1, ITC Reports Created 0, Report Creation Errors 1 | the same counts |
| Clients `Zeta Traders` (first file `AnnualReport-Zeta Traders-…`) and `Alpha Traders` (first file `GSTR3B-Alpha Traders-…`), both ticked | Zeta Traders first, then Alpha Traders | Alpha Traders first, then Zeta Traders |

**What it changes.** If the author changes a template on Step 1 while a batch of clients is running (for example after spotting that the previous year's template was selected), A's batch silently ends up with reports from two templates, split at the client being processed at that moment; B's whole batch uses the template that was chosen when Start was pressed. The client order decides which client's files win when two clients share a folder (2.2) and how the whole-run running totals in each organisation text report read (rule A-R62).

**When.** The live reads are the shared Tk code (`47a1492:gui/handlers/processing_handler.py:238, 247, 269`, from `edac6a3`); A's table order is the shared Tk code. B's snapshot and sorted table came with `c258047` (07-02-2026); no reason is given. Described only.

#### 8.7 The Step 3 summary workbook (A-R61 / B-R94, B-R95)

**What both do.** The same workbook code (funcdiff: `create_summary_report` and its five sheet functions IDENTICAL; the column-width helper `_autofit_columns` differs only in the kind of error it catches): sheets Summary, Client Status, File Mapping, Errors, Variations; written when at least one client was processed (A) or succeeded (B), which mean the same here because A counts both together (A:gui/handlers/processing_handler.py:416-417).
**What differs.** A hands over the scan's variation list (A:gui/handlers/processing_handler.py:224); B hands over an empty list (B:gui/controllers/processing_controller.py:73). If writing the workbook fails, A shows "❌ Failed to create summary report" or "❌ Summary Report Error: …" in the Step 3 log (A:gui/handlers/processing_handler.py:468-475; in practice the first, because A's `create_summary_report` also catches its own errors and returns False, A:core/excel_handler.py:407-409); B shows a line only when it succeeds (B:gui/controllers/processing_controller.py:189-193: its error branch at :192-193 cannot run, because B's `create_summary_report` catches every error itself and returns False, B:core/excel_handler.py:344-346; the B audit ran the failure case and saw nothing on screen, `audit/B/S3a.md`, feature B-F15, Case D).

**Worked example (RAN, `step3_probe.py`; the same three-client folder with `Claims register ClientA.xlsx` and ClientA's IMS "(1)" file).**

| Sheet | A | B |
|---|---|---|
| Summary: Total Files / Reports Generated | 10 / 6 | 9 / 6 |
| Client Status: ClientA "Files" | 7 | 6 |
| Variations | `Claims register ClientA.xlsx` · "No matching pattern found" · `ImsReco-ClientName-State-DDMMYYYY.xlsx` · 2004 · "Rename file" | "No file variations found" |

**What it changes.** The summary workbook is the only place where A lists the unused source files; B has no such place (1.6). The counts differ only through 1.1. Not a GST rule.

**When.** As 1.6: the real list at the split (`47a1492:gui/handlers/processing_handler.py:161`), the empty list in B since `c258047` (07-02-2026).

#### 8.8 Where the report is assembled (A-R98 / B-R52)

**What A does.** Copies the template to a local temporary folder (`%TEMP%\gst_excel_temp`), fills and saves it there, and copies the finished report to the client's Version folder once; the temporary copy is removed afterwards (Excel path A:core/excel_handler.py:147-155, 215, 254-260; openpyxl path :273-279, 306-308, 315-320).
**What B does.** Copies the template straight to the final report name, waits 0.5 s on the Excel path, and fills the copy in place (B:core/excel_handler.py:122-129, 227-231). B also creates the temporary folder and never uses it (B:core/excel_handler.py:52-53).

**Worked example (RAN on the openpyxl path, `core_probe.py` and `step3broken`; the Excel path is READ only).** The ITC template is a damaged file:

| | A | B |
|---|---|---|
| `_create_report_openpyxl` returns | False ("Fallback method error: File is not a zip file" in the log file) | False (the same log line) |
| File left under the report name | none | `ITC_Report_ClientA_DL_211026_1030.xlsx`, which cannot be opened (not a valid workbook) |
| Temporary folder afterwards | empty | (not used) |
| Step 4 later | nothing to list for ITC | the file matches `ITC_Report_*.xlsx`, so B's Step 4 lists it as ClientA's ITC report (B:gui/pyqt_main_window.py:461) |

**What it changes.** After a failed report, B leaves a broken or unfilled "report" in the client's Version folder under the normal name, which Step 4 then offers for refreshing; A leaves nothing there. Technical point (file safety): A's way never leaves a half-made report under the final name; B's does. A's commit also gives speed on the network share as the main reason.

**When.** B keeps the shared way (`edac6a3`). A changed it in `12d505f` (23-07-2026, "Build Excel reports on local disk, copy to network once"); the message explains that Excel's many small reads and writes are slow over a network share, and notes "On failure the destination is no longer left with a half-written template copy (minor improvement)" and that the Excel path "could not be executed in CI; validate one report on a Windows machine before bulk processing". Not a GST rule.

### Categories 1 to 8 at a glance

| Category | Rule pairs (rows) | same | different | GST-rule pairs that differ |
|---|---|---|---|---|
| 1 Input-file recognition | 17 | 10 | 7 (1.1-1.7) | 1.1 IMS "(N)" ⚠; 1.6 hint shown ⚠; 1.7 example names ⚠ |
| 2 Client identity and name cleaning | 4 | 1 | 3 (2.1, 2.2 covering A-R19 and A-R20) | 2.1 client key handed back by Step 2 |
| 3 State, GSTIN and PAN codes | 4 | 4 | 0 | - |
| 4 Financial year and tax period | 1 | 1 | 0 | - |
| 5 Expected files and completeness | 6 | 4 | 2 (5.1, 5.2) | 5.1 status display |
| 6 Folder layout and naming | 14 | 9 | 5 (6.1-6.5) | 6.1 length box (tagged only through B-R32: a small value drops the state code); 6.4 type → folder ⚠ (folder names only) |
| 7 Copying and versioning | 8 | 6 | 2 (7.1, covering A-R43 and A-R44) | - |
| 8 Report creation and Links cells | 14 (plus two pointer rows) | 5 | 9 (8.1-8.8, 8.3 covering A-R54 and A-R55) | 8.1 Dry Run check; 8.2 template types (tagged only through B-R51); 8.3 Links folder text ⚠; 8.4 merged Links cell ⚠; 8.5 report recognition; 8.6 live template reads and client order |
| **Total** | **68** | **40** | **28** | **14 pairs** (A-R2, A-R12, A-R16, A-R17, A-R29, A-R36, A-R41, A-R51, A-R52, A-R54, A-R55, A-R58, A-R59, A-R60 with their B partners; A-R36 and A-R52 are tagged only through their B partners B-R32 and B-R51) |

Every GST rule that is the same in both projects in these categories comes from the shared commits and was not changed in behaviour by B after the split (B's `c258047` and `7e09e5c` rewrote some of these functions, removing log-file debug lines and moving the size and signature checks of 1.3-1.4, without changing what these rules decide): the five other file-name patterns, the matching method and the handling of unrecognised names (A-R1, A-R3 to A-R7, A-R11), the state table, fallback codes and their uses (A-R21 to A-R23), the absence of GSTIN, PAN and period logic (A-R24, A-R25), the expected types, duplicate, completeness and "critical" rules (A-R26 to A-R28, A-R30), the client-folder name (A-R35), the copy rename (A-R48 / B-R57), and the form of the Links values and the "first file in sorted order" rule (A-R56, A-R57).

### Category 9: Power Query refresh

This category is Step 4's first half: for each ticked client, the program makes a copy of the ITC and Sales report (the "refreshed copy"), opens it in Excel, makes Excel re-run the report's Power Query (Excel's built-in data-import tool; "refresh" means re-running its queries so the report picks up the client's latest files) and decides whether the refresh worked. Both projects do this through COM (Windows' way for one program to drive another; here Python drives Excel), which needs pywin32 (the add-on Python package that provides COM).

**How the examples in categories 9 to 16 were run.** Every "RAN" below means: a copy of each project's code at its snapshot was run on the same made-up inputs (invented clients `ClientA-DL`, `ClientB-MH`, invented figures and paths) with `python -B` in `/root/merge-scratch/run/C5b/` (scripts `scripts/p9.py`, `p11a.py`, `p11b.py`, `p12.py`, `p13.py`, `p14.py`, `p15.sh`, `p16.py`; outputs in `out/`). This Linux machine has no Excel and no pywin32, so where Excel is involved a stand-in was used: `C5b/fakewin/` holds fake `pythoncom`, `win32com`, `win32gui`, `win32api`, `win32con` modules whose "Excel" reads and writes the made-up workbook with openpyxl (a Python library that reads and writes `.xlsx` files without Excel) and records every call the program makes. The stand-in proves the ORDER of calls, the waits and what the program does with the answers; it cannot prove what real Excel does, so everything Excel itself does stays READ. Windows screens were built off-screen (Tk is A's window toolkit and Qt is B's; `xvfb-run` and `QT_QPA_PLATFORM=offscreen` give each an invisible screen on this server: Tk under `xvfb-run`; Qt with `QT_QPA_PLATFORM=offscreen`). Three runs cited below as "C5b `p145`", "C5b `p166`" and the B double-scan of 13.4 (and the `B2:C2` merge of 12.2) were not kept as scripts; the check VERIFY-C5b re-ran each of them in `/root/merge-scratch/run/VERIFY-C5b/scripts/` (`v_145.py`, `v_btn.py`, `v_btnB.py`, `v_trunc.py`, `v_scan2.py`, `v_merge.py`) with the same results, except where a correction says otherwise. All the kept scripts above were re-run there too, with identical results apart from time stamps.

#### 9.1 Rule by rule

| Rule | A's version | B's version | Relation |
|---|---|---|---|
| Step 4 without pywin32 | A-R64: pywin32 imported inside a guard; each report fails with a clear message, the rest of the app runs. `A:power_query_extractor/core/report_processor.py:20-36, 66-77` | B-R124 (category 15): imported without a guard; the whole window cannot start. `B:power_query_extractor/core/report_processor.py:11-15` | **different** (9.2) |
| Which reports are processed | A-R65: one row per client; the ITC and Sales ticks start TICKED where the report exists; unticked → "Not selected for processing". `A:power_query_extractor/gui/extractor_window.py:693-694, 849-861`; `A:power_query_extractor/core/report_processor.py:86-104` | B-R61 (part) and B-R136 (ticks): one row per Version folder (the `Version-<DDMMYY HHMM>` folder that holds one Step 3 run's files for a client); every tick starts UNTICKED; unticked → the same "Not selected for processing". `B:power_query_extractor/core/report_processor.py:50-68`; `B:gui/views/extraction_view.py:291-390` | **different** (9.3) |
| Which file is the "original" report | A-R66: first file the folder listing returns for `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`. `A:power_query_extractor/core/report_processor.py:111-125` | B-R61: the same search and "first file" rule. `B:power_query_extractor/core/report_processor.py:72-89` | **same** - [GST RULE] in B's registry (the numbered rule list each audit keeps) only (9.8) |
| Name of the refreshed copy | A-R67: `<stem><suffix>.xlsx`, `{timestamp}` → `DDMMYY_HHMM`; default `_Refreshed_{timestamp}` typed in 3 places. `A:power_query_extractor/core/report_processor.py:127-130, 162-164` | B-R62: identical code; default typed in 5 places. `B:power_query_extractor/core/report_processor.py:91-94, 118-120` | **same** behaviour (9.9) |
| Recognising a refreshed copy | A-R68: derived from the CONFIGURED suffix, plus the old marker `_refreshed_`. `A:power_query_extractor/core/refresh_naming.py:14-66` | B-R63 (Skip) and B-R136 (status): fixed texts `<stem>_Refreshed_*.xlsx` and `*_Refreshed*`. `B:power_query_extractor/core/report_processor.py:100-102`; `B:gui/pyqt_main_window.py:466-479` | **different** (9.4; status text in 16.12) |
| Skip refresh | A-R69: newest recognised copy reused as success; the tick writes the shared option at once and is never saved. `A:power_query_extractor/core/report_processor.py:132-159`; `A:power_query_extractor/gui/extractor_window.py:483-492` | B-R63: the same reuse logic with B's fixed search; the tick is written into the shared option at each Start and saved in settings. `B:power_query_extractor/core/report_processor.py:96-115`; `B:gui/controllers/extraction_controller.py:53-55` | **different** (9.4) - [GST RULE] by this comparison |
| `auto_refresh_if_missing` | A-R70: never read. `A:power_query_extractor/config/cell_mappings.py:250` | B-R64: never read (same file). | **same** (dead) |
| How the refresh is triggered | A-R71: `Dispatch` (start or connect to Excel through COM), visible maximised Excel, `AppActivate` (bring the Excel window to the front), keys Alt+A, R, A ⚠, always save. `A:power_query_extractor/core/report_processor.py:209-415` | B-R66, B-R67, B-R69: identical steps. `B:power_query_extractor/core/report_processor.py:165-366` | **same** (9.5) |
| Waits, retries and the wait box | A-R72: same fixed waits; box 5-60, typed values not checked. `A:power_query_extractor/gui/extractor_window.py:269-287, 864` | B-R68 (waits), B-R65 (box): same waits; box 5-120, enforced by Qt; value saved in settings. `B:gui/views/extraction_view.py:126-128` | **different** (box only) (9.6) |
| When a refresh counts as successful | A-R73: no error dialog AND `Info!BB2` ⚠ filled, not 0, not an error text. `A:power_query_extractor/core/report_processor.py:638-685` | B-R70, B-R71: identical test. `B:power_query_extractor/core/report_processor.py:589-636` | **same** - [GST RULE] ⚠ (9.7) |
| "Could not open - continue without refresh" branch | Same unreachable branch in A, `A:power_query_extractor/core/report_processor.py:167-181` (A's audit lists it under gap A-G74, "Code that looks like protection but never acts"; it cannot run because `_refresh_power_query_simple` catches every error itself) | B-R72: unreachable. `B:power_query_extractor/core/report_processor.py:122-137` | **same** (dead) |
| Step 4 options (wait, suffix, Skip) | A-R72 (box), A-R119 (category 14: never stored) | B-R65: box 5-120, values read at Start and stored in `settings.json` | **different** (9.6 and 14.6) |

#### 9.2 Step 4 on a PC without pywin32 (A-R64 against B-R124)

- **A.** The five pywin32 modules are imported inside a guard (`A:power_query_extractor/core/report_processor.py:20-36`). Without them, a warning is logged and every report gets the status "pywin32 is not installed for this Python, so Power Query cannot be refreshed. Install it with: pip install pywin32" (`:66-77`; pip is Python's package installer); Steps 1 to 3 still work.
- **B.** The same five modules are imported without a guard (`B:power_query_extractor/core/report_processor.py:11-15`), and that module is loaded as soon as the main window's code is imported (`B:gui/pyqt_main_window.py:25` imports from the controllers package, whose file `B:gui/controllers/__init__.py:4` imports the extraction controller - line 27 would do so too - which imports the refresh module at `B:gui/controllers/extraction_controller.py:10`), so nothing at all opens. [Corrected during review - originally "`B:gui/pyqt_main_window.py:27` imports the extraction controller"; audit B's self-review (its correction 29, B-R124) found that line 25 is reached first; READ both files.]
- **Example (RAN, `p9.py nostub` - "nostub" means without the stand-in pywin32 - and `p15.sh` case 2B).** The same made-up Version folder with an ITC and a Sales report, on this PC without pywin32:

| | A | B |
|---|---|---|
| Loading the refresh module | works; `WIN32_AVAILABLE = False`; log "pywin32 not available (No module named 'pythoncom'). Power Query refresh is disabled…" | `ModuleNotFoundError: No module named 'pythoncom'` |
| Processing the client | ITC and Sales both `success False`, error "pywin32 is not installed for this Python, so Power Query cannot be refreshed…"; the reports on disk untouched | not reachable |
| Starting the program (`main.py`) | starts (after the install offer, 15.3) | console: "Missing module: No module named 'pythoncom'" / "Please install required packages:" / "pip install PyQt5 openpyxl"; exit code 1 |

- **What it changes in a client's annual working.** On an office Windows PC without pywin32 neither program gets as far as Step 4: A's start-up check refuses to start (after offering to install it, 15.2 and 15.3), B cannot open its window and its advice ("pip install PyQt5 openpyxl") does not fix the cause. A's guard matters where A does start without a working pywin32: on a non-Windows PC, or on Windows when pywin32 is present on disk but cannot be loaded (A's check only looks for the module file, `A:utils/requirements_check.py:64-70`; A audit RAN XL E2, simulated). There A still sorts the season's GST files and builds the ITC/Sales reports with openpyxl (their Power Query is lost, 12.2) and Step 4 shows every client as failed with a clear message. Technical point only.
- **When it came in.** A's guard: `c463256`, 26-07-2026, "Fix: app failed to start when pywin32 was missing". B's unguarded import is the shared code of `edac6a3` (31-08-2025); it became a start-up requirement in `c258047` (07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"), when the main window began to import the extractor.

#### 9.3 Which reports are sent to the refresh (A-R65 against B-R61 / B-R136)

- **Same.** An unticked report gets `{success: False, error: 'Not selected for processing'}` and no data in both (RAN `p9.py` (a): identical results).
- **Different.** A lists one row per client folder (only its newest Version folder, A-R42 in category 6) with the ticks already ON where the report exists. B lists every Version folder that holds a report, with every tick OFF, and sends one request per ticked Version folder named `<client folder> / <version folder>`.
- **Example (RAN, `p11a.py` / `p11b.py`).** Made-up run folder `Annual Statement-311225 1000` with `ClientA-DL` (two Version folders: `Version-301225 0900` with an ITC report only; `Version-311225 1000` with ITC and Sales) and `ClientB-MH` (one Version folder, ITC report plus an earlier refreshed ITC copy):

| | A (after Scan) | B (after the page fills) |
|---|---|---|
| Rows | `ClientA-DL` (Version-311225 1000), `ClientB-MH` | `ClientA-DL` → `Version-311225 1000`, `Version-301225 0900`; `ClientB-MH` → `Version-311225 1000` |
| Ticks at start | ClientA: ITC ✓ Sales ✓; ClientB: ITC ✓ (Sales cannot be ticked) | all off; "Start" with nothing ticked would be refused |
| Sent after the user's minimum action | A: press Start → 2 requests (ClientB ITC only; ClientA ITC + Sales) | B: press Select All, then Start → 3 requests (`ClientA-DL / Version-311225 1000` ITC + Sales; `ClientA-DL / Version-301225 0900` ITC; `ClientB-MH / Version-311225 1000` ITC) |

- **What it changes.** In B the older Version folder of ClientA is refreshed too if the user presses Select All, which costs another Excel session (about 39 s at the default wait, 9.6) and puts a second ClientA row into every sheet of the consolidated workbook (Step 4's single workbook that collects every client's extracted figures, category 11; 11.3). In A only the newest Version folder can be refreshed. Technical/presentation point; which Version folder's figures should feed the season's working is the author's call (the choice of folder itself is category 6, A-R42 / B-R39).
- **When it came in (READ).** The ITC and Sales ticks, started ON where the report exists, are the shared `bd176bf` (30-10-2025; `bd176bf:power_query_extractor/gui/extractor_window.py:575-576`); A kept them. B's `c258047` (07-02-2026) still started the ticks ON per client; B's per-Version-folder tree with every tick OFF came with `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking").

#### 9.4 Recognising refreshed copies and "Skip refresh" (A-R68, A-R69 against B-R63) · [GST RULE] (treated as one here)

Neither registry tags these rules. This comparison tags Skip refresh **[GST RULE]**, "when in doubt": it decides which file the figures in the consolidated workbook are read from - a fresh refresh, an older refreshed copy, or the unrefreshed original. Below is only what each project does.

- **A.** The literal pieces of the CONFIGURED suffix around `{timestamp}` are the markers (`_PQ_{timestamp}` → `_pq_`); a file counts as a refreshed copy if, in lower case, it starts with the original's name, ends `.xlsx` and contains all markers, OR contains the old marker `_refreshed_` (`A:power_query_extractor/core/refresh_naming.py:23-66`). Skip refresh reuses the newest such copy (`A:power_query_extractor/core/report_processor.py:132-159`).
- **B.** Skip refresh looks only for `<original name>_Refreshed_*.xlsx`, whatever the suffix box says (`B:power_query_extractor/core/report_processor.py:100-102`).
- **Both.** The reused copy is reported as success "Skipped - Using existing refreshed file" without any check that its own refresh had worked (shared gap A-G2 / B-G7); with no copy found, the report is refreshed normally.
- **Example (RAN, `p9.py` (d2)).** ITC report `ITC_Report_ClientA_DL_311225_1000.xlsx`, Skip refresh ticked, refresh itself replaced by a recorder. The probe sorted the folder listing by name, as a Windows NTFS disk (the usual Windows disk format) lists it (MEDIUM) (on this Linux machine's own order the copy came first in S2 and S5 and both projects then took the copy as the "original" - the shared copy-of-a-copy defect A-G15 / B-G8, section 7):

| Case | Version folder holds (made-up) | Suffix box | A | B |
|---|---|---|---|---|
| S1 | original + `…_1000_Refreshed_311225_1100.xlsx` | `_Refreshed_{timestamp}` (default) | reuses the `_Refreshed_` copy, no Excel | same |
| S2 | original + `…_1000_PQ_311225_1100.xlsx` | `_PQ_{timestamp}` | reuses the `_PQ_` copy, no Excel | does not see it; refreshes again into `…_1000_PQ_<now>.xlsx` |
| S3 | original only | `_R{timestamp}` | treats the ORIGINAL report itself as a refreshed copy (`_r` is inside `_report_`), reuses it and reports success; nothing is refreshed | refreshes into `…_1000_R<now>.xlsx` |
| S4 | original + `…_1000_Refreshed_311225_1100.xlsx` | `_PQ_{timestamp}` | reuses the old `_Refreshed_` copy (old marker) | reuses the same copy (fixed search) |
| S5 | original + `…_1000_Refreshed.xlsx` | `_Refreshed` (no `{timestamp}`) | reuses it | does not see it (no `_` after "Refreshed"); refreshes again and overwrites `…_1000_Refreshed.xlsx` |

- **What it changes in a client's annual working.** Suppose the staff refreshed 40 clients' ITC reports last evening with the suffix `_PQ_{timestamp}` and today tick Skip refresh only to rebuild the consolidated workbook. A reuses yesterday's copies in seconds; B refreshes all 40 again (about 40 × 39 s ≈ 26 minutes of Excel at the default wait) and the figures then come from today's refresh. With the suffix `_R{timestamp}`, A would put the figures of the never-refreshed Step 3 reports under a ✓ in the consolidated workbook; B would refresh. Which behaviour is wanted is the author's decision; the over-match itself (`_R`, `_` or a client name containing "Refreshed") is recorded as part of gap A-G16.
- **The tick itself.** A writes the tick into the shared option the moment it is clicked (RAN `p11a.py`: option `True` right after the tick; `A:power_query_extractor/gui/extractor_window.py:483-492`) and never saves it (A-R119). B writes it at every Start (`B:gui/controllers/extraction_controller.py:53-55`) and saves it as `pq_skip_refresh` (B-R116).
- **When each version came in.** Shared: Skip refresh with the fixed `<stem>_Refreshed_*.xlsx` search is already in the first commit, `edac6a3` (31-08-2025, "Initial commit"; `edac6a3:power_query_extractor/core/report_processor.py:76-93`, READ); the editable suffix box that the fixed search ignores came with `bd176bf` (30-10-2025, "Enhance report processing UI with configurable wait time and file suffix options; add ITC/Sales report selection functionality"); B kept both unchanged. A: `2aaa2b1` (25-07-2026, "Recognise refreshed files by the CONFIGURED suffix, not a hardcoded one"); its message gives the reason: with a customised suffix "'skip refresh' found nothing and re-refreshed every report instead of reusing the existing files - the opposite of what the option is for", and adds that files made under the old default are still accepted "so changing the setting does not make the whole existing history read as 'Never refreshed'". B's write of the option at each Start: `c258047` (07-02-2026).

#### 9.5 How the refresh is done (A-R71 against B-R66, B-R67, B-R69) - same

The two refresh functions differ only by an unreachable `return` that B deleted and by the kind of error a few `except` lines catch (`git diff A/main B/main -- power_query_extractor/core/report_processor.py` - git's line-by-line comparison of the two versions - READ). RAN with the fake Excel (`p9.py` (e)): the recorded call sequence is identical in both - `CoInitialize` (COM's start-up call for one line of work; `CoUninitialize` is its shut-down), `Dispatch("Excel.Application")`, `Visible = True`, `DisplayAlerts = False`, `WindowState = -4137` (maximised), open the copy, `Dispatch("WScript.Shell")` (Windows' scripting helper, used here to send key presses), `AppActivate`, keys `%a`, `r`, `a` (Alt+A, R, A = Data → Refresh All in an English Excel ⚠: ribbon key tips depend on the Excel version and display language), `Save`, `Close(SaveChanges=True)`, `Quit`, `CoUninitialize`; then a second, hidden Excel for reading. With a made-up error dialog on screen, both clicked its "OK" button and both still saved the workbook. Nothing in either project uses `RefreshAll`, `BackgroundQuery = False` or `CalculateUntilAsyncQueriesDone` (READ). Shared since `edac6a3`.

#### 9.6 Waits and the "wait time" box (A-R72 against B-R68, B-R65)

- **Same waits.** RAN with the fake Excel and a counting `sleep` (`p9.py` (e)): 39 s per refreshed report at the default wait of 10 s and 139 s at 60 s, in both (= 19 + w + max(10, w) seconds, before Excel's own start-up and opening time).
- **Different box.**

| | A | B |
|---|---|---|
| Range of the arrows | 5 to 60, hint "seconds (Min: 5, Max: 60)" (`A:power_query_extractor/gui/extractor_window.py:270-287`) | 5 to 120 (`B:gui/views/extraction_view.py:126-128`) |
| Typed `0` | accepted: 0 (RAN `p11a.py`) | refused: on Enter the box goes back to its previous value (10 here) (RAN VERIFY-C5b `v_spin.py`, keys typed into the box). A value set by the program, e.g. from the settings file, is clamped instead: `setValue(0)` → 5 (RAN `p11b.py`) |
| Typed `200` | accepted: 200 | the third digit is refused, the box shows 20 (RAN `v_spin.py`); `setValue(200)` → 120 (RAN `p11b.py`) |
| Typed `12.7` | read as 12 | the "." is refused, the box shows 12 (whole numbers only; RAN `v_spin.py`) |
| Typed `abc` or nothing | `TclError` (Tk's error message type) when Start reads it - the click does nothing (A audit RAN PQ gui_probe; RAN `p11a.py`) | letters refused; an emptied box goes back to its previous value (RAN `v_spin.py`) |
| Remembered next time | no, back to 10 (A-R119) | yes, `pq_wait_time` (B-R116) |

- **What it changes.** By the formula above: in A a typed 0 makes each report wait 29 s in all, a typed 200 makes it 419 s; B stays between 34 s (5) and 259 s (120). For a firm refreshing 2 reports × 150 clients at the end of the annual return season, the default 39 s per report is about 3¼ hours of Excel driving in either project. Technical point. History: the 5-60 box is the shared `bd176bf` (30-10-2025); B's 5-120 is `c258047` (07-02-2026).

#### 9.7 When a refresh counts as successful (A-R73 against B-R70, B-R71) · [GST RULE] · ⚠ - same

- **What both do.** Success = no error dialog was found AND a sheet named `info` (any letter case) exists AND its cell `BB2` ⚠ (a template cell position that can change with the template) is not empty, not blank, not `0` / `0.0`, and not one of `#REF!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#ERROR!`. The function is identical in both (`A:power_query_extractor/core/report_processor.py:638-685` = `B:power_query_extractor/core/report_processor.py:589-636`).
- **Worked example through both (RAN, `p9.py` (f), identical results):**

| `Info!BB2` holds (made-up) | Result in A and B |
|---|---|
| empty | fail - "Cell BB2 is empty" |
| `''` or spaces | fail - "Cell BB2 is blank" |
| `0`, `0.0`, `'0'`, `'0.0'` | fail - "Cell BB2 contains 0" |
| `'00'` | pass - "Validated - Client Name: 00" |
| `#N/A`, `#n/a`, `#VALUE!` | fail - "Cell BB2 contains error: …" |
| `ClientA` (the client being processed) | pass |
| `ClientB` (another client's name) | pass - BB2 is never compared with the client |
| `12.5`, `-2146826246`, `True`, `False` | pass |
| no sheet named info | fail - "Info sheet not found" |

  In both, a failed test still saves the workbook and still reads every figure (RAN `p9.py` (e): BB2 empty → status failed, J15 and AA1 read), and the consolidated workbook then shows "Missing" for every figure of that report (11.4).
- **When it came in.** Same code in both since `edac6a3` (31-08-2025, "Initial commit"); neither project changed it. Nothing in either repository writes BB2; the code's comment and log call it the client name (READ). Whether `Info!BB2` is still the cell the current templates fill is for the author to confirm ⚠.

#### 9.8 The registry difference on "which file is the original" (A-R66 / B-R61)

Same code in both (RAN `p9.py` (c): same listing, same original taken). B's registry tags it [GST RULE] ("when in doubt: which file the figures come from"); A's registry does not. Following the "when in doubt" rule, this comparison treats it as a [GST RULE]; there is nothing to choose between the projects. Shared since `edac6a3` (31-08-2025, "Initial commit"); the ITC/Sales choice since `bd176bf` (30-10-2025). On this Linux machine the search listed an older refreshed copy before the original in some folders, and both projects then took that copy as the original and refreshed a copy of it (shared gap A-G15 / B-G8; RAN VERIFY-C5b `v_copyfirst.py`, Skip refresh off, copy listed first: both refreshed `…_1000_Refreshed_311225_1100_Refreshed_<now>.xlsx`; on a Windows NTFS disk the original normally lists first, MEDIUM).

#### 9.9 Copy name (A-R67 / B-R62) - same

RAN `p9.py` (b), identical in both: `_Refreshed_{timestamp}` → `…_Refreshed_<DDMMYY_HHMM>.xlsx`; `_PQ_{timestamp}` → `…_PQ_<…>.xlsx`; `{timestamp}` → the stamp glued to the name; `_Refreshed` → the same name every run; `_Refreshed_{TIMESTAMP}` (capitals) → the text `{TIMESTAMP}` stays in the name; an empty suffix → "… are the same file" and the report fails. Only the number of places where the default text is typed differs (A 3: `A:power_query_extractor/core/report_processor.py:51`, `A:power_query_extractor/gui/extractor_window.py:297`, `A:power_query_extractor/core/refresh_naming.py:14`; B 5: `B:power_query_extractor/core/report_processor.py:28`, `B:models/app_state.py:22`, `B:gui/controllers/extraction_controller.py:24, 139`, `B:gui/views/extraction_view.py:136`) - a code-tidiness point, no behaviour difference.

### Category 10: Extraction cell mappings

After the refresh, a second, hidden Excel opens the refreshed copy and reads fixed cells. The list of cells is the extraction mapping (a table in the code saying "take this cell of this sheet of this report and put it in this column of this sheet of the consolidated workbook"). It lives in `power_query_extractor/config/cell_mappings.py`, which is **byte-for-byte the same file in both projects** (RAN: `git rev-parse` gives blob - git's fingerprint of a file's exact content - `598495f` for both). Every row is a **[GST RULE]**, and every return name, table number (GSTR-9 Tables 6J, 8D), sheet name and cell address is **⚠** (it follows the template version of the time).

#### 10.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| "GSTR-3B Difference" ← ITC `Diff GSTR-3B` ⚠ J15, K15, L15 | A-R74 | B-R73 | **same** [GST RULE] ⚠ |
| "GSTR-1 Difference" ← Sales `Diff - GSTR-1` ⚠ N1, O1, P1, Q1 | A-R75 | B-R74 | **same** [GST RULE] ⚠ |
| "Turnover GSTR-1" ← Sales `Diff - GSTR-1` ⚠ I1 | A-R76 | B-R75 | **same** [GST RULE] ⚠ |
| "Reco Diff" ← ITC `Reco` ⚠ C37, D37, E37 | A-R77 | B-R76 | **same** [GST RULE] ⚠ |
| "Table 6J Match" ← ITC `T6 of R9` ⚠ K29 | A-R78 | B-R77 | **same** [GST RULE] ⚠ |
| "Table 8D Limits" ← ITC `T8 of R9` ⚠ G6 | A-R79 | B-R78 | **same** [GST RULE] ⚠ |
| "IMS Reco" ← ITC `IMS_Reco` ⚠ AA1, AD1, AG1 (shows J15/K15/L15 instead, 10.3) | A-R80 | B-R79 | **same** [GST RULE] ⚠ |
| "Extra In Purchase" ← ITC `ExtraInPurchase` ⚠ O1, P1, Q1 (same mismatch) | A-R81 | B-R80 | **same** [GST RULE] ⚠ |
| "Extra In GSTR2B" ← ITC `ExtraInGSTR2B` ⚠ S1, T1, U1 (same mismatch) | A-R82 | B-R81 | **same** [GST RULE] ⚠ |
| How mappings become cell reads (exact, case-sensitive sheet names; each value stored twice) | A-R83 (`A:power_query_extractor/core/report_processor.py:527-636`) | B-R82 (`B:power_query_extractor/core/report_processor.py:478-587`) | **same** [GST RULE] (only the kind of error an `except` catches differs) |
| Consolidated look-up: first field CONTAINING the column title, else the exact field | A-R88 (category 11, `A:power_query_extractor/core/data_consolidator.py:137-150`) | B-R83 (`B:power_query_extractor/core/data_consolidator.py:125-138`) | **same** [GST RULE] (10.3) |
| `FIELD_DISPLAY_NAMES` read only by dead code | A-R84 | B-R85 | **same** (dead) |
| Calculated columns (subtract / add / multiply / divide; 'Error', 'Div/0', 'Unknown'), unused | A-R85 | B-R84 | **same** [GST RULE] (dormant); B narrowed one bare `except:` to `except (TypeError, ValueError)` in `c258047`, same results |

There is no difference in this category, so no choice between the projects arises. The decisions the author still has to take are about the mapping itself (which cell is right), not about A or B.

#### 10.2 Worked example: the 22 cells through both (RAN, `p9.py` (g))

A made-up ITC report with a different invented number in every mapped cell (J15 = 101, K15 = 102, L15 = 103, C37 = 201 …, AA1 = 501, AD1 = 502, AG1 = 503, O1 = 601 …, S1 = 701 …, K29 = "Yes", G6 = 401.5) and a made-up Sales report (N1 = 801 … Q1 = 804, I1 = 900000.25) were read by each project's own reading code through the fake Excel. The 22 values read, the stored keys (each value twice: under the field name, e.g. `IMS_Reco_IGST`, and under the bare cell, e.g. `AA1`) and the whole consolidated workbook came out **identical** in A and B. With the sheet renamed `diff gstr-3b` (lower case), both stored `None` for the three fields of J15/K15/L15 and did not store the bare cell keys (sheet names are matched exactly; `A:power_query_extractor/core/report_processor.py:570-576`) while AA1 was still read as 501; and a result with J15 = `None` put AA1 = 501 into the "IMS Reco" IGST column in both (10.3, row "None").

#### 10.3 The 9-column mismatch, in both (A-R80 to A-R82, A-R88 / B-R79 to B-R81, B-R83) · [GST RULE] · ⚠

Technical description only. In both projects the consolidated workbook fills the IGST, CGST and SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" from `Diff GSTR-3B` J15, K15, L15 whenever those cells were read with a value, because the one-word titles "IGST", "CGST", "SGST" are found inside the field names `GSTR-3B_Difference_Diff_IGST` etc. first. RAN (same results in both):

| `Diff GSTR-3B`!J15 read as | "GSTR-3B Difference" Diff IGST shows | "IMS Reco" IGST shows | "Extra In Purchase" IGST shows |
|---|---|---|---|
| 101 | 101 | **101** (configured AA1 = 501 not shown) | **101** (configured O1 = 601 not shown) |
| `None` (empty cell) | Missing | 501 (AA1) | 601 (O1) |
| `''` (empty text) | Missing | **Missing** | **Missing** |
| 0 | 0 | **0** | **0** |

For a client's GSTR-9 working this means: the three IMS Reco tax-head figures and the six "extra" figures in the consolidated workbook are, today, the GSTR-3B difference figures in both projects. Which cell each column should hold is the author's call. History: the look-up is from `edac6a3` (31-08-2025); the affected one-word columns came with `47a1492` (30-10-2025, "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets"), the last shared commit, so the mismatch is in A and B alike (gap A-G1 / B-G1).

**When each mapping came in.** "GSTR-3B Difference", "GSTR-1 Difference" and "Turnover GSTR-1": `edac6a3` (31-08-2025). "Reco Diff", "Table 6J Match", "Table 8D Limits", "IMS Reco", "Extra In Purchase", "Extra In GSTR2B": `47a1492` (30-10-2025). Neither project changed the file after the split.

### Category 11: Consolidated output layout

The consolidated workbook (`PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, one workbook per Step 4 run with a "Summary" sheet and one sheet per output group) is built by `DataConsolidator.create_report`, which is the same in both projects apart from B's deletion of a leftover header write that was overwritten anyway (`git diff … -- power_query_extractor/core/data_consolidator.py`, READ; RAN `p9.py` (g): identical workbooks from identical results). The differences come from what each project SENDS to it and what it does around it. B's registry also files the Step 3 summary workbook (`GST_Processing_Summary_…xlsx`, B-R94, B-R95) here; A files it under category 8 (A-R61), so it is compared here as well.

#### 11.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| File name, folder, when written, errors | A-R86: always written; into the run folder chosen by the LAST Step 4 scan; an error loses the results and shows an error box. `A:power_query_extractor/gui/extractor_window.py:941-966` | B-R86: written only if at least one result; into the run folder shown when Start was pressed; an error becomes one log line and the Extraction_Log is still written. `B:gui/controllers/extraction_controller.py:86-101`. In both, an error raised while openpyxl is already writing the file (for example a date with a time zone) leaves a half-written workbook that cannot be opened under the normal name (Section 9, M4, RAN both; audit B's B-R86 as corrected in its self-review). [Corrected during review - added from audit B's final text.] | **different** (11.2) |
| Sheet layout; column A label | A-R87: one row per client, column A = client folder name | B-R87: one row per Version folder, column A = `<client folder> / <version folder>` | **different** (11.3) |
| Which extracted value goes into a column | A-R88 | B-R83 (category 10) | **same** [GST RULE] (10.3) |
| Failed refresh → every figure "Missing" | A-R89 | B-R88 | **same** [GST RULE] (11.4) |
| `#,##0.00` for non-zero numbers (display only) | A-R90 | B-R89 | **same** [GST RULE] (11.4) |
| Summary sheet | A-R91: "Total Clients" = clients sent | B-R90: "Total Clients" = Version folders sent | **same code, different count** (11.3) |
| Description lines (dormant); leftover header write | A-R92 (leftover still there, overwritten) | B-R91 (leftover deleted in `c258047`) | **same** output |
| `show_timestamp` never read | A-R93 | B-R92 | **same** (dead) |
| `_write_section`, `_create_details_sheet` never called | A-R94, A-R95 | B-R93 | **same** (dead) |
| Step 3 summary workbook: when written, failure feedback | A-R61 (category 8): when at least one client was processed - A's `processed` counter counts exactly the successful clients (`A:gui/handlers/processing_handler.py:416-417`), so the condition is the same as B's; failure → log line "❌ Failed to create summary report". `A:gui/handlers/processing_handler.py:455-475` | B-R94: when at least one client succeeded; failure → nothing on screen. `B:gui/controllers/processing_controller.py:183-193` | **different** (failure feedback only) (11.5) |
| Step 3 summary workbook: content | A-R61: five sheets; "Variations" lists the unrecognised files | B-R95: same five sheets; "Variations" always "No file variations found" | **different** (11.5) |

#### 11.2 Where and when the consolidated workbook is written (A-R86 against B-R86)

- **Example (RAN, `p11a.py` / `p11b.py`, refresh replaced by the same canned results with invented figures):**

| Situation | A | B |
|---|---|---|
| Normal run | `target1/Annual Statement-311225 1000/PQ_Extraction_Report_<now>.xlsx`, then the question "Processing complete! Open the consolidated report?" | the same workbook in the same folder, plus `Extraction_Log_<now>.txt` beside it; no question |
| The user browses Step 4 to a second target and scans it while the run is going | the workbook goes to the SECOND target's run folder (`target2/Annual Statement-010126 0900/…`) although it holds target1's clients | the workbook stays in `target1/Annual Statement-311225 1000` (the page shows the new folder, the run keeps its own) |
| A value the workbook writer cannot store (here a made-up two-cell result) | error box "Processing failed: Cannot convert (1, 2) to Excel"; no workbook; the figures exist only in the log (READ: the extraction writes every figure to the program log, A-R122; in this run the refresh and extraction were replaced by canned results) | log line "Report error: Cannot convert (1, 2) to Excel"; no workbook; `Extraction_Log_<now>.txt` written |
| Stop pressed before the first Version folder | (A's Step 4 has no Stop) | "Extraction stopped by user", "Extraction complete: 0/3 clients processed"; no workbook; Extraction_Log written |

- **What it changes.** During the annual return season a staff member who checks another client batch on Step 4 while A is refreshing would find the consolidated workbook filed under the other batch's run folder; in B it stays with its own batch. When one cell cannot be written, B at least leaves a text log of every figure in the run folder; A leaves the figures only in the program log. Technical points.
- **When it came in.** A: the folder of the last scan since `edac6a3` (`47a1492:power_query_extractor/gui/extractor_window.py:808`, same line in the shared Tk extractor). B: `c258047` (07-02-2026) wrote it into the target folder itself and only "if results"; `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking") moved it back into the run folder and added the Extraction_Log.

#### 11.3 One row per client or per Version folder; "Total Clients" (A-R87, A-R91 against B-R87, B-R90)

- **Example (RAN, same tree as 9.3, ticks as each project starts plus Select All in B):**

| | A | B |
|---|---|---|
| Rows of "Turnover GSTR-1" | `ClientB-MH` ✗ Missing; `ClientA-DL` ✓ 250000 | `ClientA-DL / Version-311225 1000` ✓ 250000; `ClientA-DL / Version-301225 0900` ✗ Missing; `ClientB-MH / Version-311225 1000` ✗ Missing |
| Summary "Total Clients" | 2 | 3 |
| "ITC Reports Processed" / "Sales Reports Processed" | 2/2 and 1/2 | 3/3 and 1/3 |
| Row order | as the file system lists the client folders (no sorting; ClientB came first here) | client folders by name; Version folders by name from Z to A (`B:gui/pyqt_main_window.py:453, 457`) - the highest `DDMMYY HHMM` name first, which is not always the newest because the day comes first (as in B-R39) |

- **What it changes.** A client's annual working from B's workbook can contain two or more rows for the same client (one per Version folder ticked), labelled with the folder names; A's has one row per client, labelled with the client folder name (e.g. `ClientA-DL`). B's "Total Clients" counts Version folders, not clients. Layout point for the author; the cells and figures in each row are the same rules (category 10).
- **When it came in.** The consolidator is shared (`edac6a3`). A's per-client list is the shared Tk extractor's; B's per-version tree and labels: `b0e7d0e` (09-02-2026).

#### 11.4 "Missing" and number display (A-R89, A-R90 against B-R88, B-R89) · [GST RULE] - same

RAN `p9.py` (g), identical in both: a report whose refresh status is not success shows "Missing" in every column although its figures were read (made-up ClientB-MH row); an empty value shows "Missing"; 0 stays 0; non-zero numbers get the display format `#,##0.00` (two decimals, thousands separators) while the stored value is not rounded; text stays text. B's audit also RAN that a text starting with `=` is stored as a live formula; the same line is in A (READ, identical code). Shared since `edac6a3` (31-08-2025, "Initial commit"); neither project changed these lines after the split.

#### 11.5 The Step 3 summary workbook (A-R61 against B-R94, B-R95)

- **Same.** Five sheets (Summary, Client Status, File Mapping, Errors, Variations), same layout; the writing code is identical (MAP: `ExcelHandler.create_summary_report` IDENTICAL).
- **Different.** What is passed in and what happens on failure. A passes the scan's real list of problem files (`A:gui/handlers/processing_handler.py:224`); B passes an empty list (`B:gui/controllers/processing_controller.py:73`).
- **Example (RAN, `p16.py`: a real Step 3 run on the same made-up source folder, openpyxl path, with one unrecognised file `Purchase register ClientB.xlsx`):**

| Sheet "Variations" | A | B |
|---|---|---|
| Row 2 | `Purchase register ClientB.xlsx` · "No matching pattern found" · "Check file naming convention" · 4859 · "Rename file" | "No file variations found" |

  The "Summary" sheets were identical (Total Clients 2, Successful 2, Total Files 8, Reports Generated 4, Processing Mode fresh).
- **What it changes.** After organising a season's files, A's summary workbook tells the staff which input file was not picked up for which reason; B's says there were none although the scan found one (B shows only a count on screen). The hint text in that sheet is itself a [GST RULE] (A-R12 / B-R12, category 1; MAP 2.1 item 2), so this difference also decides whether the hint is ever seen.
- **When it came in.** The workbook: `edac6a3`. B's empty list: `c258047` (07-02-2026); at the split the real list was passed (`47a1492:gui/handlers/processing_handler.py:161`, READ). A summary that fails: A logs "❌ Failed to create summary report"; B shows nothing because `create_summary_report` catches its own errors and B's worker skips only the success line (B's audit RAN, XL E4).

### Category 12: Excel handling and safety

This category covers how each project starts, uses and closes Excel, and what happens to files when something goes wrong. The author's Standard 12 ("Excel through COM", `rebuild_pipeline_FINAL.md:844-878`) is the yardstick in Section 8; here only A and B are set side by side. Both projects share almost all of this code; the one real difference is where Step 3 builds a report while Excel works on it.

#### 12.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Excel or openpyxl for Step 3 reports | A-R96: Excel only on Windows with pywin32, else openpyxl (Power Query dropped); reachable wherever the pywin32 import fails - normally off Windows, since on Windows A's start-up check refuses to start without pywin32 unless the module file is present but cannot be loaded (12.2, 15.2). `A:core/excel_handler.py:19-28, 89-102` | B-R97: the same condition and fallback, but unreachable through the app (no window without pywin32). `B:core/excel_handler.py:19-28, 83-93` | **different** (reachability) (12.2) |
| Step 3 Excel session | A-R97: `GetActiveObject` (the user's Excel) else `Dispatch`; hide; save twice; always `Quit`. `A:core/excel_handler.py:104-260` | B-R96, B-R110: identical. `B:core/excel_handler.py:95-217` | **same** (12.4) |
| Where the report is built | A-R98: on a local copy in `%TEMP%\gst_excel_temp` (`%TEMP%` is Windows' folder for temporary files), copied to the Version folder once at the end. `A:core/excel_handler.py:135-155, 212-215, 254-260` | B-R52 (category 8), B-R105: template copied straight to the report's final name and edited there; `gst_excel_temp` created and never used. `B:core/excel_handler.py:52-53, 122-129` | **different** (12.3) |
| A report counts as created if the file exists | A-R99 | B-R59 (category 8) | **same** |
| "Already open" check by full path | A-R100 (shared helper `A:utils/excel_com.py:19-29` since `ba2a753`) | B-R98, B-R99 (the same loop typed twice) | **same** behaviour |
| Step 3 open with one 8.3 short-path retry (the old 8-character DOS form of a long path) | A-R101 | B-R98 | **same** |
| Step 4 open with long-path prefix (`\\?\`, which tells Windows to accept paths beyond its usual length limit) and fallbacks | A-R102 | B-R99 | **same** |
| Step 4 Excel sessions (two `Dispatch` per report; hidden reading session) | A-R103 | B-R101, B-R66, B-R110 | **same** (RAN 9.5) |
| 218 characters treated as Excel's path limit ⚠ | A-R104 | B-R100 | **same** |
| Error-dialog search and clicking ⚠ (window words and classes assume an English Windows/Excel) | A-R105 | B-R102, B-R103 | **same** (RAN 9.5: "OK" clicked in both) |
| openpyxl load with `keep_links=True` (keep links to other workbooks), macros not kept | A-R106 | B-R97 | **same** |
| `EXCEL_SAFETY` stored, never read | A-R107 | B-R106, B-R107, B-R108, B-R109 | **same** (dead) |
| All Python warnings switched off | A-R108 | B-R104 | **same** (RAN `p12.py`: first warning filter `('ignore', None, Warning, None, 0)` in both) |
| `CoInitialize` / `CoUninitialize` inside each Excel function | part of A-R97 and A-R103 | B-R110 | **same** |

#### 12.2 The openpyxl fallback: reachable in A, not in B (A-R96 against B-R97)

- RAN `p12.py`: on this PC `USE_WIN32COM` is `False` in both, and the same openpyxl path builds the report in both. In A that path is what a user actually gets on a PC where pywin32 cannot be loaded (Windows refuses to start A without pywin32, 15.2, so in normal office use it is a non-Windows or broken-pywin32 case). In B the path cannot be reached because B's window does not open without pywin32 (RAN `p15.sh` 2B).
- **A side effect found while running (category 8 rule, wording correction).** On this path a Links cell (a cell of the template's Links sheet, into which the program writes folder paths and file names for the report's queries) that sits INSIDE a merged range but is not its top-left cell (made-up template with `A2:B2` merged, mapped cell B2): A skips that cell silently and reports success (B2 empty, the other cells written); B fails the whole report ("Fallback method error: 'MergedCell' object attribute 'value' is read-only") and leaves the template under the report's name. When the merge STARTS at the mapped cell (`B2:C2`), both write it and succeed (RAN, openpyxl of both virtual environments; re-run VERIFY-C5b `v_merge.py`: same results, and the `A2:B2` case again A True with B2 empty, B False). This sharpens MAP row A-R58 / B-R53 and B-R97 ("a merged target cell makes the report fail").
- **When it came in.** Fallback: `edac6a3`. B's narrower error handling that turns the merged cell into a failure: `c258047` (07-02-2026); its message lists "Fixed all 26 bare except: clauses with specific exception types".

#### 12.3 Where Step 3 builds a report (A-R98 against B-R52, B-R105)

- **A.** Copies the template to a local staging file (a temporary local copy that is worked on and then copied to its final place) `%TEMP%\gst_excel_temp\<report name>`, lets Excel work there, copies the finished file to the Version folder once and deletes the staging file.
- **B.** Copies the template straight to `…\Version-…\ITC_Report_<client>_<code>_<DDMMYY_HHMM>.xlsx` and lets Excel open and save that file in place; it also creates `gst_excel_temp` at every run and never uses it.
- **Example (RAN, `p12.py`).** Made-up ITC template; each project's own Step 3 code; Excel = the fake Excel in the COM cases:

| Case | A | B |
|---|---|---|
| Normal (fake Excel) | returns True; report at the final name with B2 = the GSTR-3B folder text; staging folder empty afterwards | returns True; the same report; `gst_excel_temp` exists, empty |
| Excel cannot open the file (made-up error "-2147352567 … Open method of Workbooks class failed" on every attempt) | returns False; **no file** at the report name | returns False; **the unfilled template sits under the report name** (B2 empty, template text in B3) |
| openpyxl path, template that is not a workbook | returns False; no file at the report name | returns False; a byte-for-byte copy of the bad template under the report name |

- **What it changes.** In B a failed report leaves a file that looks like a finished ITC report (right name, right folder) but has empty Links cells; Step 4 later lists it as a report to refresh (HIGH: its name matches the search `ITC_Report_*.xlsx`), and its refresh would run on the template's own Links text (MEDIUM; not runnable here). In A the Version folder simply has no report for that client, which the log names as failed. On a network share A also writes each report in one sequential copy instead of Excel's many small reads and writes (A's code comment, `A:core/excel_handler.py:135-146`, which also states that the report's data sources are absolute paths in the Links sheet, so the copy is byte-identical). Technical point (file safety).
- **When it came in.** A: `12d505f` (23-07-2026, "Build Excel reports on local disk, copy to network once"); at the split the template was copied straight to the output path (`core/excel_handler.py:135@47a1492`), which is what B still does. The `gst_excel_temp` folder itself: `edac6a3`.

#### 12.4 The shared Excel session (A-R97, A-R103 against B-R96, B-R101, B-R110) - same

RAN with the fake Excel (`p12.py`): identical call sequences in both. With an Excel "already open by the user" (the fake answered `GetActiveObject`), both used it, set `Visible = False`, `DisplayAlerts = False`, `ScreenUpdating = False` on it, saved, set alerts and screen updating back on, and then called `Quit` on the USER'S Excel ("excel.Quit USER-OWN-EXCEL"); `Visible` was never set back. Without a running Excel both started one with `Dispatch` and quit it. Each report waited 1.0 s of fixed pauses (12.1 / category 13). What real Excel then does with the user's own unsaved workbooks is not testable here (MEDIUM: a hidden "save changes?" prompt or lost work). Same in both since `edac6a3`; it is a shared gap (taking over, hiding and quitting the user's own Excel: A-G54 / B-G63; clean-up that can skip `Quit`: A-G28 / B-G38; the comparison with Standard 12 is Section 8), not an A-or-B choice.

### Category 13: Performance and operational constants

#### 13.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| How the source folder is read in Step 1 | A-R109: one directory listing, its cached details reused. `A:utils/helpers.py:309-365` | B-R1, B-R3 (category 1): three wildcard searches, then a size query and a read per file. `B:utils/helpers.py:249-266`; `B:core/file_parser.py:171-215` | **different** (13.2) |
| Signature reads (reading a file's first 8 bytes to confirm it is a real Excel file) in parallel | A-R110: up to 16 at once from 4 files on. `A:utils/helpers.py:267-272, 346-361` | — (one after another) | **only A** (13.2) |
| Run folder worked out once per run | A-R111. `A:core/file_organizer.py:45-47, 150-161` | — (recomputed for every client, `B:core/file_organizer.py:146-179`) | **only A** (13.3) |
| Step 3 fixed pauses (0.5 s after copy, 0.5 s after close) | A-R112. `A:core/excel_handler.py:155, 210` | B-R114. `B:core/excel_handler.py:125, 180` | **same** (RAN `p12.py`: 1.0 s per report in both) |
| Setup checklist re-check 350 ms after typing | A-R113. `A:gui/tabs/setup_tab.py:506-534` | — (no checklist) | **only A** (16.5) |
| Background jobs | A-R135, A-R143: Step 1 scan on the window's own thread (a thread is one line of work; the window's own thread is the one that draws the screen, so the window freezes while it is busy); Step 3 and Step 4 in daemon threads (killed instantly when the program ends); Step 4 can be started twice | B-R111: scan, processing and extraction each in its own Qt thread; a second job of the same kind refused silently | **different** (13.4) |
| Stop | A-R135: Step 3 only, between clients | B-R112: Step 3 between clients and Step 4 between Version folders | **different** (13.4) |
| Progress and time-left display | A-R141 (category 16): bar split into copying and reports, "📊 Progress: X.X%" | B-R113: bar moves between clients only; ends "Complete \| Xm Ys"; scan progress per file | **different display, same tracker** (13.5) |

#### 13.2 Reading the source folder (A-R109, A-R110 against B-R1, B-R3)

- **Example (RAN, `p13.py`).** A made-up source folder with 24 recognisable files of 4 clients. Every file-system call on that folder was counted, and each was given an invented 20 ms delay as a stand-in for an office network share (the real delay of the share is not known here):

| | A | B |
|---|---|---|
| Directory listings | 1 | 3 |
| Size/type queries (`stat`) | 2 | 26 |
| File opens (signature reads) | 24, in a pool of reader threads | 24, one after another |
| Time with 20 ms per call | 0.11 s | 1.08 s |
| Clients found | 4 | 4 |

- **What it changes.** For a firm whose GST files for the annual return season sit on a share reached over a VPN (a private network link into the office from outside), A's scan is the faster one; A's commit message (`4f84ba9`) reports "40.78s wall clock for 2508 files" on the real share, against "a ~677s cold sequential projection" (about 11 minutes; a projection from measured per-file times, not a timed run). B's scan runs in a background thread, so its window stays usable while it scans (13.4); A's window freezes until the scan ends. Technical point.
- **When it came in.** A: `a603c02` (23-07-2026, "Speed up network file scanning (behavior-preserving)"), `56aaad9` (25-07-2026, "Reuse directory-listing metadata: 81% fewer round-trips per file"), `4f84ba9` (25-07-2026, "Read Excel signatures concurrently: 11 min -> 41s on the real share"). B: the shared three-search scan of `edac6a3`, kept.

#### 13.3 Run folder once per run (A-R111)

- **Example (RAN, `p13.py`).** Mode "Re-run" (which reuses the newest existing `Annual Statement-…` folder), two clients; between client 1 and client 2 a newer run folder `Annual Statement-311225 1015` appears on the share (for example a colleague's run):

| | A | B |
|---|---|---|
| ClientA | `Annual Statement-311225 1000\ClientA-DL` | `Annual Statement-311225 1000\ClientA-DL` |
| ClientB | `Annual Statement-311225 1000\ClientB-DL` | **`Annual Statement-311225 1015\ClientB-DL`** |

- **What it changes.** In B one run can split its clients over two run folders, so its summary workbook (written into the first client's run folder) and Step 4 (which shows one run folder) miss the clients that went elsewhere; A keeps the whole run in one folder and lists the target folder once instead of once per client. Technical point. History: A `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor"); B: the shared per-client look-up of `edac6a3`.

#### 13.4 Background jobs and Stop (A-R135, A-R143 against B-R111, B-R112)

- RAN (C5b, Qt off-screen; re-run VERIFY-C5b `v_scan2.py`): two scans started at once in B → the second was refused with only the log-file line "Scan already in progress" (`B:gui/controllers/scan_controller.py:59-61`); one result came back. RAN `p11b.py` run 4: Stop before the first Version folder → "Extraction stopped by user", no consolidated workbook, Extraction_Log written. A: the scan runs on the window's thread (`A:gui/handlers/file_handler.py:71-72`, READ), Step 3 and Step 4 run in daemon threads (`A:gui/handlers/processing_handler.py:129-134`; `A:power_query_extractor/gui/extractor_window.py:878-883`), Step 4 has no Stop and its Start button is switched back on by any Step 4 scan, even during a run (A audit RAN, PQ gui_probe).
- **What it changes.** In B a staff member can stop a long Step 4 run between Version folders; in A the only way is to close the program, which kills the daemon thread in the middle of a client (Excel may be left open, MEDIUM). Technical point. History: A's Step 3 re-entrancy guard (a check that ignores a second Start while a run is going) `33d0baf` (23-07-2026); B's thread pattern `c258047` (07-02-2026).

#### 13.5 Progress display (A-R141 against B-R113)

The time-left tracker and its time texts are the same code (MAP: `ProgressTracker.*` IDENTICAL; RAN `p13.py`: `format_duration` gives "0s", "59s", "1m 1s", "59m 59s", "1h 1m", "25h 1m" in both). Only the display differs: A moves the bar inside each client (first half copying, second half reports) and shows "📊 Progress: X.X%" and "Client i of n: <client> | ETA: …" (ETA = estimated time left); B moves the bar only between clients, shows "Client i/N: <client> | ETA: …" and ends with "Complete | Xm Ys". Presentation only.

### Category 14: Settings, cache, logs and run history

#### 14.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Where settings live, file format | A-R114: one file `~/.gst_organizer_cache.json`, one line of JSON (a plain-text settings format). `A:gui/utils/cache_manager.py:14-48` | B-R115, B-R116: folder `~/.gst_organizer/` with `settings.json` (indented), `session_state.json`, `run_history.json`. `B:models/cache_manager.py:16-62` | **different** (14.2) |
| Which keys | A-R115: `source_folder`, `itc_template`, `sales_template`, `target_folder`, `processing_mode`, `workflow_mode`, `include_client_name`, `dark_mode`, `client_name_max_length`, `recent_folders`. `A:gui/main_window.py:358-389` | B-R116: 11 keys incl. `include_client_name_in_folders`, `pq_wait_time`, `pq_suffix_pattern`, `pq_skip_refresh`. `B:models/app_state.py:10-46` | **different** (14.2) |
| When settings are saved | A-R116: after a Browse, the workflow radio, the length arrows, the global tick, dark mode - NOT a typed path, NOT on closing | B-R116: at every scan, after processing, dark mode, File > Save Settings, on closing - NOT at Start | **different** (14.3) |
| Recent folders (never read) | A-R117: `recent_folders` in the settings file | B-R117: `recent_source_folders` in `session_state.json` (plus `active_tab`, `window_size`, also never read) | **different place, both unused** (14.3) |
| One-time import of A's settings file | — | B-R118: imports `~/.gst_organizer_cache.json` once | **only B** (14.4) |
| How the standalone extractor finds its folder | A-R118: home settings file first, then three local files. `A:power_query_extractor/extractor_main.py:36-63` | B-R128 (category 15): the old entry point cannot even be imported | **different** (14.5) |
| Step 4 options stored? | A-R119: no (10 s, `_Refreshed_{timestamp}`, Skip off at every launch) | B-R65, B-R116: yes | **different** (14.6) |
| Main log file `gst_organizer.log` in the current folder | A-R120. `A:main.py:20-27` | B-R121. `B:main.py:15-23` | **same** (RAN `p15.sh`: created in the folder the program was started from, in both) |
| Extractor log file | A-R121: standalone extractor writes `pq_extractor.log` | B-R127, B-R128: B's launcher sets up no log file | **different** (14.5) |
| What is logged | A-R122: client names, Links values (first 50 characters), every extracted figure; `_log` always "info" | B-R121, B-R123: the same lines and the same eight text cut-offs (50 / last 50 / 100 / 200 characters; READ, same lines in both files) | **same** |
| Run history and extraction log | A-R123: none | B-R119, B-R120, B-R122: `run_history.json`, `Extraction_Log_<YYYYMMDD_HHMM>.txt` per Step 4 run | **only B** (14.7) |

#### 14.2 One settings file or a settings folder (A-R114, A-R115 against B-R115, B-R116)

RAN `p14.py`: A writes one line of JSON to `~/.gst_organizer_cache.json`; B writes `~/.gst_organizer/settings.json` (13 lines, indented, UTF-8 - the standard text encoding) plus `session_state.json`. Both rewrite the whole file in place (a crash in the middle of a write leaves a broken file; both then start with defaults). Neither follows Standard 13's `%LOCALAPPDATA%\[App Name]\` (Section 8). History: A's file is the shared `edac6a3` design; B's folder `c258047` (07-02-2026).

#### 14.3 The same settings file through both programs (A-R114 to A-R117 against B-R116 to B-R118)

- **Example (RAN, `p14.py`, real windows, fresh scratch home).** The file A writes, with made-up paths: source `/srv/share/GST/Inputs-FY2425`, mode `resume`, workflow `full`, include-name ON, a typed max length `0`, recent folders `[…FY2425, …FY2324]`. Then the user types a new source `/srv/share/GST/Inputs-FY2526` into the box and closes the window.

| Step | A | B |
|---|---|---|
| Window opens: max length box shows | 35 (A replaces 0 or less by 35) | 15 (B's import copies 0; Qt's box keeps 15-100) |
| Files after opening | unchanged | `.gst_organizer/settings.json` written (`client_name_max_length` 0, `include_client_name_in_folders` true, the Step 4 keys at their defaults), `session_state.json` with only `…FY2425` (the old `recent_folders` list is NOT carried over), marker `.migrated` |
| After typing a new source and closing | settings file unchanged - the typed path is lost | `settings.json`: source `…FY2526`, max length **15**; `session_state.json`: `[…FY2526, …FY2425]`, `active_tab` 0, `window_size` [1400, 900] |
| A save as made by Browse (A) | `source_folder` `…FY2526`, max length 35, `recent_folders` `[…FY2526, …FY2425, …FY2324]` | — |

- **What it changes.** A staff member moving from A to B keeps the four paths, the processing mode and the include-name tick, loses the workflow choice and the recent list, and a typed 0 becomes 15 characters in B (35 in A) - so client folder names are cut at a different length (category 6). In A, a path typed (not browsed) before closing is forgotten at the next start. Technical points.
- **When it came in.** A's save points: shared Tk code plus `510f7b4` (26-07-2026, `workflow_mode`). B's settings, close-time save and import: `c258047` (07-02-2026).

#### 14.4 B's import of A's settings file (B-R118) - only B

Shown in 14.3. The import is the only place where B reads the file A writes today; it runs once (marker `.migrated`) and leaves the old file in place (RAN). Two limits of "once", from audit B's final text: the marker is written even when `settings.json` could not be saved, because the save only logs its error; and deleting the marker while the old file still exists makes the next start import the old settings again, over the current ones (READ `B:models/cache_manager.py:100-101, 119, 127, 135-142`; RAN in audit B, B-F2 bugs 3 and 5, B-G24). [Corrected during review - added from audit B's self-review (its corrections 42-43).] A has nothing to import.

#### 14.5 Standalone extractor and its log (A-R118, A-R121 against B-R127, B-R128)

RAN `p15.sh`: A's `launch_extractor.py` (case 4A) found the target in `~/.gst_organizer_cache.json` ("Found target folder in cache: /srv/share/GST/Out-FY2425"), opened the extractor alone and wrote `pq_extractor.log` in the current folder. B's `launch_extractor.py` (case 4B) printed "Failed to launch: No module named 'pythoncom'" (with pywin32 it opens the full window on Step 4 and sets up no log file at all, READ `B:launch_extractor.py:1-28`). B's old `power_query_extractor/extractor_main.py` (case 5B) cannot be imported: `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'` (deleted by `a62d3cc`, 07-02-2026). A: home settings file first since `c1583d6` (23-07-2026, "Fix extractor cache path and version-folder selection"). Technical point.

#### 14.6 Step 4 options remembered or not (A-R119 against B-R65, B-R116)

RAN: B's second start (`p14.py`) showed the suffix `_PQ_{timestamp}` typed in the first session; A's panel starts at `_Refreshed_{timestamp}`, wait 10 and Skip off every time (`A:power_query_extractor/gui/extractor_window.py:270, 297, 315`; RAN `p11a.py` default 10). Combined with 9.4: in A, copies made yesterday with a custom suffix are, after a restart, recognised only if they also contain `_refreshed_` (RAN C5b `p145`, re-run VERIFY-C5b `v_145.py`: with the default suffix in the box A showed "Never refreshed" for a `_PQ_` copy). History: B `c258047` (07-02-2026).

#### 14.7 Run history and the extraction log (A-R123 against B-R119, B-R120, B-R122) - only B

- RAN (`p14.py`, B's own `RunHistory`): a processing record is kept under the scan key `ClientA-DL` (it feeds Step 2's "Last Run" column); the Step 4 result for the same client is kept under `ClientA-DL / Version-311225 1000`, a separate record that no screen reads, with `pq_refreshed: true` although the made-up refresh had FAILED; after 61 runs, 50 run entries were kept and `total_runs` was 61. RAN `p11b.py`: `Extraction_Log_<YYYYMMDD_HHMM>.txt` begins "PQ Extraction Log - …", 60 "=", then every Step 4 screen line.
- **What it changes.** B keeps a permanent record in the user profile of every extracted figure, path and BB2 text for every client (client data stored outside the client folders, relevant to confidentiality); A keeps no history, only the workbooks and the log files. Technical/confidentiality point. History: B `c258047` (07-02-2026) and `b0e7d0e` (09-02-2026, Extraction_Log).

### Category 15: Start-up and dependencies

#### 15.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Which packages are required and checked | A-R124: openpyxl essential; pywin32 essential on Windows only; checked at start with a report. `A:utils/requirements_check.py:17, 37-84`; `A:main.py:68-106` | B-R124: nothing checked; pywin32 needed on every system just to open the window | **different** (15.2) |
| Offer to install missing packages and restart | A-R125. `A:utils/dependency_installer.py:26-158` | — | **only A** (15.3) |
| `requirements.txt` ⚠ | A-R126: `openpyxl>=3.1.2`, `pywin32>=305; sys_platform == 'win32'`, `lxml>=4.9.0` (an optional, faster XML reader), `pytest>=7.0.0` (the tool that runs the automated tests) | B-R124: `PyQt5>=5.15.0`, `openpyxl>=3.1.2`, `lxml>=4.9.0`; pywin32 only as the comment `# pywin32>=306` | **different** (15.4) |
| Console handling | A-R127: console output always re-wrapped as UTF-8 (no console → crash); banner and report; waits for Enter on exit paths. `A:main.py:17, 45-73` | B-R125: re-wrapped only if a console exists; prints only on failure. `B:main.py:11-13, 32-59` | **different** (15.5) |
| Preparing the program folder | A-R128: program folder first on the search path (the folders Python looks in for program code); creates `core`, `utils`, `gui`, `logs`, `temp` and placeholder `__init__.py` files (small files that mark a folder as Python code). `A:main.py:29-43` | B-R125, B-R127: search path only; nothing created | **different** (15.5) |
| Step 4 loading guard card | A-R129. `A:gui/tabs/extract_tab.py:21-30, 84-105` | — | **only A** |
| Unused version labels | A-R130: `3.0.0` in core, gui, utils; `1.0.0` in power_query_extractor | B-R129 (category 16): `4.0.0` in gui, `3.0.0` in core and utils, `1.0.0` in power_query_extractor (RAN grep) | **different** (labels only) |
| Window start-up order; Step 4 filled from disk at start | — (A builds Step 4 empty; no scan at start-up, A-R143, A-R144) | B-R126: fills Step 4 by reading the saved target folder on the window's thread. `B:gui/pyqt_main_window.py:35-68, 174-190` | **only B** (16.13) |
| Launcher | A's `launch_extractor.py` opens the extractor alone (A-R128, A-F29) | B-R127: opens the full window on Step 4; no log set-up | **different** (14.5) |
| Old extractor entry point | A's works (A-R118, A-R121) | B-R128: cannot be imported | **different** (14.5) |

#### 15.2 Requirements check (A-R124 against B-R124)

- **Example (RAN, A's own check on this Linux PC and with Windows simulated by loading the check while `sys.platform = 'win32'` - re-run VERIFY-C5b `v_req.py`, same texts; B by starting `main.py`, `p15.sh` 2B):**

| | A on Linux | A on Windows (simulated) | B |
|---|---|---|---|
| pywin32 missing | "[not found] pywin32 …" - optional, the program continues | "[MISSING] pywin32 …" - essential: error, wait for Enter, exit 1 (after the install offer) | "Missing module: No module named 'pythoncom'" / "pip install PyQt5 openpyxl", exit 1 |

- **What it changes.** On an office Windows PC without pywin32, A refuses to start and says exactly which package is missing and why ("Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them."), so no report can silently lose its Power Query; B also refuses to start, with advice that does not name pywin32. Technical point. History: A `f363946` (26-07-2026, "Check dependencies at startup and report them clearly") and `be6418c` (26-07-2026, "Require pywin32 on Windows so reports cannot lose Power Query"); B `c258047`.

#### 15.3 Offer to install (A-R125) - only A

RAN: A's command is `[<this python>, -m, pip, install, -r, <program folder>/requirements.txt]` with a 600-second limit; the WHOLE file is installed (pytest included). On this Linux PC the offer appears at every start because pywin32 stays "missing" (A audit RAN START; READ `A:main.py:78-88`; `p15.sh` 2A ran until its 15-second limit, which fits a wait at the offer, but its console text was never flushed, so that run alone does not show the dialog). History: `a5ad487` (26-07-2026, "Offer to install missing dependencies, then restart").

#### 15.4 `requirements.txt` ⚠ (A-R126 against B-R124)

READ (both files). Minimum versions only in both; package versions change over time ⚠. A lists pywin32 as a real requirement on Windows (`064c634`, 23-07-2026) and pytest (`6c1e1c6`, 23-07-2026); B lists PyQt5 and leaves pywin32 as a comment (`c258047`). Standard 13 asks for exact versions (Section 8).

#### 15.5 Console and program folder (A-R127, A-R128 against B-R125, B-R127)

- **Example (RAN, `p15.sh`, fresh copies of each program, standard output closed as under `pythonw.exe`, the Windows Python that runs without a console window):**

| | A | B |
|---|---|---|
| Start with no console | dies at once: `AttributeError: 'NoneType' object has no attribute 'buffer'` (`A:main.py:17`), before any log file | passes the console step (`B:main.py:12` checks first); on this PC stops later at the missing pywin32, writing "Import Error: No module named 'pythoncom'" to `gst_organizer.log` |
| Folders created in the program folder on a normal start | `logs`, `temp` (never used) | none |

- **What it changes.** If the office later starts the tool from a desktop shortcut without a console (Standard 13's launcher uses `pythonw.exe`), A cannot start at all and B starts but shows its start-up errors nowhere on screen (only in `gst_organizer.log`, RAN `p15.sh` 1B). A on a read-only network program folder fails the first time `logs`/`temp` do not exist yet (A audit, MEDIUM). Technical point. History: A's unconditional wrapper is the shared `edac6a3` line; B added the check in `c258047`.

### Category 16: User-interface behaviour rules

This category holds the screens' fixed values: window, tabs, buttons, tables, messages and help texts. Two items here are GST rules: the Dry Run gate (what must be chosen before a preview of the reports, 16.4) and the file names shown to the user as the expected input names (16.11). The Step 2 table (16.7) also shows a third, the client status (A-R29 / B-R27, MAP 2.1 item 5), whose rule belongs to category 5. Everything else is presentation or technical behaviour.

#### 16.1 Rule by rule

| Rule | A | B | Relation |
|---|---|---|---|
| Main window | A-R131: "GST File Organizer v3.0", 1200×800, minimum 1000×700, tabs with emoji. `A:gui/main_window.py:57-77` | B-R129: "GST File Organizer v4.0", 1400×900, minimum 1200×800, File and View menus. `B:gui/pyqt_main_window.py:70-118` | **different** (16.2) |
| Automatic tab switches | A-R132 | B-R140 (with coloured tab dots) | **different** (16.3) |
| Setup checklist | A-R133. `A:gui/utils/setup_validation.py:13-115` | — | **only A** (16.5) |
| Scan and Dry Run gates | A-R134 (Dry Run uses A-R51's template gate) | B-R133 (Dry Run needs only a ticked client; also the processing and extraction gates) | **different** (16.4) - Dry Run gate [GST RULE] per MAP 2.1 item 7 |
| Run control (what is greyed out) | A-R135: Start greyed, Stop active; Step 2/3 buttons react to the mouse even when greyed, not to the keyboard | B-R138: Start/Stop swap; everything else stays active; a disabled Stop looks enabled | **different** (16.6) |
| Long client-name warnings | A-R136 | — (B's help note still promises one, B-R141) | **only A** |
| Per-client include-name toggle on Step 2 | A-R137 | — (B-R36, category 6: nothing sets it) | **only A** |
| Step 2 client table | A-R138 | B-R134, B-R27 (status colours and "Complete Only", category 5) | **different** (16.7) - the status display is [GST RULE] per MAP 2.1 item 5 |
| Dry-run text | A-R139 | B-R135 | **different** (16.4) |
| Step 3 log and completion | A-R140 | B-R137 (HTML logs - HTML is web-page markup), B-R139 (feedback) | **different** (16.8) |
| Progress and time left | A-R141 | B-R113 (category 13) | **different display** (13.5) |
| Export of ticked clients | A-R142 | — (B's "EXPORT LIST" button is connected to nothing, B-F10) | **only A** |
| Step 4 panel behaviour | A-R143: no Stop; question "Open the consolidated report?" | B-R138, B-R139: Stop; no end message, status "Extraction complete" | **different** (11.2, 13.4) |
| Step 4 folder | A-R144: follows Step 1's target as typed; no scan at start | B-R39 (category 6), B-R126: the last SAVED target, read from disk at start-up, after processing, after extraction, on Refresh List | **different** (16.13) |
| "Last refresh" status text | A-R145 | B-R136 | **different** (16.12) |
| Mouse-wheel rules | A-R146 | — (Qt's own scrolling) | **only A** |
| Tooltips (help text on hover) | A-R147 (9 places, RAN grep) | — (0 places, RAN grep `setToolTip`) | **only A** |
| Dark mode | A-R148: records and restores each widget's colours | B-R144: one style sheet (a single block of colour and font settings for the whole window); Steps 1-2 stay light | **different** (16.9) |
| GUI colour table | A-R149: used | B-R130: never read; B-R132: B's own colour table and `#info` style never used | **different use** |
| `GUI_CONFIG['fonts']` | A-R150 (dead) | B-R130 (dead) | **same** (dead) |
| `GUI_CONFIG['colors']['white']` | A-R151 (dead) | B-R130 (dead) | **same** (dead) |
| Message texts in use | A-R152: four `MESSAGES` texts used | B-R131: none used (texts written inline) | **different** (16.4) |
| Other message texts | A-R153 (dead) | B-R131 (dead) | **same** (dead) |
| Workflow choice ("organise only" / "full pipeline") | A-R154 | — | **only A** (16.10) |
| Title strip and status bar | A-R155: blue strip, Dark Mode tick, "v3.0 \| Production Ready" | B-R129: plain status bar "Ready" | **different** (16.2) |
| Version label | A-R156: "v3.0" | B-R129: "v4.0" | **different** (16.2) |
| Shortened paths in the status bar (`truncate_path`) | A-R157: used after Browse (`A:gui/handlers/file_handler.py:24, 54`) | B-R143: the same function, never called | **different use** (RAN C5b `p166`, script not kept; re-run VERIFY-C5b `v_trunc.py` with the made-up 118-character path `\\fileserver\GSTShare\Annual Statement-311225 1000\ClientA Made-up Name-DL\Version-311225 1000\Other ITC related files`: both give "...de-up Name-DL\Version-311225 1000\Other ITC related files" - on Linux a Windows-style path counts as one part, so only "..." + the last 57 characters remain (on Windows the path would be split into parts, MEDIUM); the function is identical text in both; only A shows it) |
| Help texts with the expected file names ⚠ | A-R16 (category 1), A-R147: the card reads the patterns' own example names | B-R141: hand-typed card | **different** - [GST RULE] ⚠ (16.11) |
| Size and time displays | A-R31 (sizes), A-R145 (`%Y-%m-%d %H:%M`) | B-R142 (`dd/mm/yy HH:MM`; history `YYYY-MM-DD HH:MM:SS`) | **different** (time formats; `format_size` identical) |
| Theme leftovers | — | B-R132 (dead) | **only B** (dead) |

#### 16.2 Window, title and version (A-R131, A-R155, A-R156 against B-R129)

RAN `p16.py`: A - title "GST File Organizer v3.0", size 1200×800, minimum 1000×700, tabs "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing", "🔄 Step 4: Extract". B - title "GST File Organizer v4.0", 1400×900, minimum 1200×800, tabs "Step 1: Setup", "Step 2: Validation", "Step 3: Processing", "Step 4: Extract && Refresh" (shown as "&"), menus File and View. Neither label tells the user which version is newer: A ("v3.0") has 34 commits of its own after the split, B ("v4.0") 5. Presentation only.

#### 16.3 Automatic tab switches (A-R132 against B-R140)

RAN `p16.py`: after a scan both switch to Step 2; Dry Run switches both to Step 3; after a real Step 3 run A (workflow "organise only") stays on Step 3 and logs "ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.", while B always moves to Step 4 (index 3) and marks Step 3 green and Step 4 orange. Presentation only.

#### 16.4 Gates, messages and the dry run (A-R134, A-R139, A-R152 against B-R133, B-R135, B-R131)

- **Example (RAN, `p16.py`, same made-up source folder: ClientA with all six files, ClientB with two):**

| Action | A | B |
|---|---|---|
| Scan with an empty source box | error box "Error": "Please select a source folder containing GST files" | warning box: "Please select a source folder first." |
| Scan with a folder that does not exist | error box: "Source folder does not exist" | warning box: "Folder not found: <path>" |
| Scan with a FILE as source | passes the gate, then error box "Scan Error": "Not a directory: <path>" | warning box: "Folder not found: <path>" |
| Dry Run, no templates chosen | error box: "Please select both ITC and Sales template files"; no preview | preview printed: each client's part ends "Would create: ITC Report" / "Would create: Sales Report", then "DRY RUN COMPLETE" and "No files were actually moved or created" |
| Start, no templates chosen | the same error box as Dry Run (READ: the same gate function, `A:gui/handlers/processing_handler.py:78`) | warning box: "Please select a valid ITC template." |
| Dry Run text | "[HH:MM:SS]" before each line; "🎯 Mode: fresh", "👥 Selected: 2 clients"; indented lines | no time, no mode, no count; leading spaces lost (HTML log, 16.8) |
| Step 4 Start with nothing ticked (READ) | warning "No Selection": "Please select at least one report to process" (`A:power_query_extractor/gui/extractor_window.py:859-861`); (Step 4 Scan with an empty folder box: "No Folder" / "Please select a folder first", `:613-615`) | warning: "No versions selected for extraction."; no run folder shown: "No Level 1 folder available."; ticked Version folders gone: "Could not find version folders for the selected items. Make sure clients have been processed first (Step 3)." (`B:gui/pyqt_main_window.py:506-538`) |

- **[GST RULE] for the Dry Run gate only** (MAP tags A-R51 / B-R48 / B-R133 as a GST rule, because it decides what must be chosen before a preview of the reports is shown). What each does is in the table; no recommendation. History: the gate is in the shared Tk code (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`); B's PyQt Dry Run without it came with `c258047` (07-02-2026). The rest of this sub-section is presentation only.
- **Correction (A's audit, A-R152).** A's audit says the four `MESSAGES` texts are "shown as `Warning` boxes before Dry Run and Start". The code shows "Please select a source folder containing GST files" (before a SCAN), "Please select both ITC and Sales template files" and "Please select a target folder for organized files" in ERROR boxes titled "Error" (`A:gui/handlers/file_handler.py:123, 135, 139`); only "Please select at least one client to process" is a "Warning" box (`A:gui/handlers/processing_handler.py:45, 83`). RAN above.

#### 16.5 Setup checklist and its re-check (A-R133, A-R113) - only A

RAN (A's `evaluate_setup` on made-up paths): source = a FILE, ITC template = a FOLDER named `tpl_folder.xlsx`, Sales template missing, target blank → source ✓, ITC ✓, Sales ✕ "File not found", target ○ "Not selected yet"; Scan allowed, Step 3 not; pill "2 of 4 items set"; still needed for Step 3: Sales template, Target folder. (The check is existence only - a file passes as a folder and a folder as a template.) B shows a fixed 7-line instruction card and checks paths only when SCAN or START is pressed (16.4). History: `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"); its message: "Each required input reports its own state as you pick it, and the buttons enable themselves - no more clicking Scan to discover something is missing".

#### 16.6 Run control (A-R135 against B-R138)

- **A.** Greys Start and enables Stop during Step 3, but its Step 2 and Step 3 buttons are bound to a mouse press instead of a button command (`A:gui/tabs/processing_tab.py:74, 87`; `A:gui/tabs/validation_tab.py:118, 130, 142, 190`). RAN (C5b `p166`, real A window under xvfb; re-run VERIFY-C5b `v_btn.py`, same result): with Start greyed as during a run, a mouse press on it still called the start handler; with Start enabled, the Space key and the button's own `invoke()` called nothing. (A's re-entrancy guard, `33d0baf`, then ignores a second Start while a run is active.)
- **B.** Swaps Start and Stop and leaves everything else active. RAN (C5b `p166`, real B window off-screen, run state switched on; re-run VERIFY-C5b `v_btnB.py`, same result): Step 3 Start disabled, Stop enabled; Step 2's START, DRY RUN, Select All, Clear All, Complete Only and EXPORT LIST all still enabled. Its red Stop style has no disabled look, so a disabled Stop looks enabled (B audit RAN MW e16; `B:gui/theme/theme_manager.py:104-111`).
- **What it changes.** In A a keyboard user cannot press Start (RAN) or the Step 2 buttons (READ: wired the same way); in B a second click on Step 2's START during a run is silently refused (13.4). Technical point.

#### 16.7 Step 2 client table (A-R138 against B-R134)

- **Example (RAN, `p16.py`):**

| | A | B |
|---|---|---|
| Columns | ✓, 🏢 Client Name, 🌏 State, 📊 Status, 📁 Files, ❌ Missing, ➕ Extra, 📁 Name | Select, Client, State, Status, Files, Missing, Extra, Folder Name, Last Run |
| ClientA row | ☐ · ClientA · Delhi · "✅ Complete" · 6 · None · 0 · ☐ No | [x] · ClientA · Delhi · Complete · 6 · 0 · 0 · ClientA-DL · (empty) |
| ClientB row | ☐ · ClientB · Maharashtra · "⚠️ Missing 4 files" · 2 · "Annual Report, GSTR-2B Reco, IMS Reco, Sales Reco" · 0 · ☐ No | [x] · ClientB · Maharashtra · Missing 4 files · 2 · **4** · 0 · ClientB-MH · (empty) |
| Rows ticked after a scan | none (Select All ticks both) | all, complete or not |

- **What it changes.** In A the staff see WHICH return files a client lacks (e.g. the IMS Reco and GSTR-2B Reco files for ClientB) and must tick clients on purpose; in B they see only the number 4 and every client - complete or not - is processed unless unticked. The status texts themselves come from the same category-5 logic in both ("Complete", "Missing <n> files", "Has duplicates"; same results in the table above). How the status is shown and used differs, and MAP 2.1 item 5 tags this pair (A-R29 / B-R27) [GST RULE]; described only: A tints the whole row #E8F5E8 for "Complete" and #FFF4E5 otherwise (`A:gui/handlers/client_handler.py:52-59`) and "Complete Only" ticks rows whose status text CONTAINS "Complete" (`A:gui/handlers/client_handler.py:152`); B colours the row's text #107C10 / #FF8C00 / #D83B01 (`B:gui/views/validation_view.py:387-393`) and "Complete Only" ticks rows whose status EQUALS "Complete" (`B:gui/views/validation_view.py:446`) (READ). With the status texts both projects produce today, the two "Complete Only" tests pick the same rows (HIGH; not run). Which display is wanted is the author's call. History: A's table is the shared Tk code; B's `c258047` / `7e09e5c`.

#### 16.8 Logs and completion feedback (A-R140 against B-R137, B-R139)

- **Example (RAN).** The message `    Links B11 = <ImsReco-ClientA-Delhi-31032025> & done` written to B's Step 3 log comes out as "Links B11 = & done": B inserts each message as HTML (a coloured `<span>`), so the leading spaces and the text between `<` and `>` vanish (`B:gui/views/processing_view.py:142-157`); B's audit RAN that 2,000 lines take about 49 s to show. A writes the same text unchanged with a "[HH:MM:SS]" stamp into a Tk text box (`A:gui/utils/ui_helpers.py:16-33`). After a real Step 3 run (`p16.py`): both showed "Processing Complete" / "Successfully processed 2/2 clients … Files organized in: <target>"; A then opened the run folder (on Linux with `xdg-open "<target>/Annual Statement-<now>"` - Linux's "open with the default program" - and on Windows with Explorer, `A:gui/handlers/processing_handler.py:508, 512`); B opens Explorer on Windows only and moves to Step 4.
- **What it changes.** A file name or note containing `<…>` in B's on-screen log is shown incomplete. Technical point.

#### 16.9 Dark mode (A-R148 against B-R144)

READ, with the audits' runs (A START dm_test; B MW e8): A records each widget's colours and restores them; B applies one dark style sheet, while Steps 1 and 2 keep their own light colours and some labels become dark on dark. B's light style also makes three Step 1 Browse buttons white on white (B audit RAN VW t5_pixels). Presentation only. History: A `aeb83cf`, `c89c1d4`, `725fbaf`, `169c7b5` (all 26-07-2026); B `c258047`, `7e09e5c`.

#### 16.10 Workflow choice (A-R154) - only A

A offers "organise only" (default) or "full pipeline"; "full" opens Step 4 after Step 3, fills its folder and scans it, without starting the refresh (`A:gui/handlers/processing_handler.py:136-175`; A audit RAN GUI exp1). B has no choice and behaves like A's "full" every time: after every Step 3 run it rebuilds the Step 4 list from disk (`B:gui/pyqt_main_window.py:384`) and moves to Step 4 (`:401`; RAN 16.3), and, like A, never starts the refresh by itself. History: `510f7b4` (26-07-2026, "Add a workflow choice: organise-only or full pipeline").

#### 16.11 Help texts with the expected file names (A-R16, A-R147 against B-R141) · [GST RULE] · ⚠

Described only; the patterns themselves are category 1.

| | A | B |
|---|---|---|
| Source of the text | each pattern's own example name, read from `FILE_PATTERNS` (`A:gui/tabs/setup_tab.py:477-503`) | a hand-typed list (`B:gui/views/setup_view.py:488-499`) |
| IMS Reco line ⚠ | "ImsReco-ClientName-State-DDMMYYYY.xlsx" | "ImsReco-Client-State-DDMMYYYY" |
| Mode lines | tooltips on the radio buttons (texts `A:gui/tabs/setup_tab.py:103-110`, attached at `:409`) | "Re-Run: Overwrite existing", "Resume: Skip already processed" (contradicting the mode buttons, B-R44) |
| Promise of a long-name warning and per-client settings | warning exists (A-R136) | note "Names >10 chars will show warning. This overrides individual client settings if checked." - neither exists in B |

Same six types in the same order in both; A's lines include `.xlsx`, B's do not; neither mentions `.xls`, the ban on hyphens inside a client or state name, or the "(N)" forms (READ). **When it came in.** A: `322717f` (25-07-2026); its code comment says the card is "read from FILE_PATTERNS so this card can never drift from the regexes the parser actually uses". B: `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections"). Older history (READ): the short hand-typed list with "ImsReco-Client-State-DDMMYYYY" and the "Names >10 chars will show warning" note were already in the shared Tk Step 1 of the first commit `edac6a3` (31-08-2025; `47a1492:gui/tabs/setup_tab.py:364` and `:238`); B's `c258047` (07-02-2026) first replaced it with a typed list that included `.xlsx` ("ImsReco-ClientName-State-DDMMYYYY.xlsx", `c258047:gui/views/setup_view.py:169`), and `7e09e5c` went back to the short form; A replaced the shared hand-typed list with the pattern-driven card in `322717f`. Which wording the staff should see is the author's call.

#### 16.12 "Last refresh" status on Step 4 (A-R145 against B-R136)

- **Example (RAN, `p11a.py`, `p11b.py` and C5b `p145`; rows 3 to 5 re-run VERIFY-C5b `v_145.py`):**

| Version folder holds (made-up) | Suffix box | A shows | B shows |
|---|---|---|---|
| ClientA: ITC + Sales reports, no copy | default | "Never refreshed" | "Pending" (orange) |
| ClientB: ITC report + `…_Refreshed_311225_1100.xlsx` | default | "ITC: 2025-12-31 04:00 \| Sales: Never" | "Refreshed (31/12/25 04:00)" (green) |
| ClientA: ITC report + `…_PQ_311225_1100.xlsx` | `_Refreshed_{timestamp}` | "Never refreshed" | "Pending" |
| the same | `_PQ_{timestamp}` | "ITC: 2025-12-31 04:00 \| Sales: Never" | "Pending" |
| the same | `_R{timestamp}` | "ITC: <date of the original report> \| Sales: <date of the original report>" - the never-refreshed originals count as refreshed because `_r` is inside `_report_` (gap A-G16) | "Pending" (the box is not consulted) |

- In both the time is the copy's date modified, so a copy left by a FAILED refresh also shows as refreshed (shared gap A-G2 / B-G7). A shows one time per report type; B one status per Version folder, where one refreshed ITC copy marks the whole folder.
- **What it changes.** A staff member checking which clients' reports are already refreshed for the GSTR-9 working sees per-report information in A and a single flag per Version folder in B; with a custom suffix that does not contain "_Refreshed", B never shows "Refreshed". Technical point. History: shared `bd176bf` (per report type, fixed `ITC_Report_*_Refreshed_*.xlsx`, `47a1492:power_query_extractor/gui/extractor_window.py:645, 652`); A `2aaa2b1` (configured suffix) and `2acffed` (25-07-2026, recomputed after a run); B `b0e7d0e` (09-02-2026, per Version folder, `*_Refreshed*`).

#### 16.13 Step 4's folder and start-up fill (A-R144 against B-R39, B-R126)

READ, with runs: A's Step 4 box follows Step 1's target as typed until the user puts another folder there, and Step 4 stays empty until Scan is pressed (`A:gui/tabs/extract_tab.py:76-131`; `bfcb24a`, 26-07-2026). B reads the last SAVED target and fills Step 4 from disk when the window opens, after processing, after an extraction that ends normally or is stopped (not after one that ends in an error; `B:gui/pyqt_main_window.py:549-565`) [Corrected during review - originally "after extraction"; Section 4, 4.7 item 2] and on Refresh List, on the window's own thread (`B:gui/pyqt_main_window.py:171-190, 419-504`; RAN `p11b.py` through that function), so a slow or unreachable network target delays B's start-up. Technical point; which run folder is chosen is category 6.

## 6. Screens and workflow

This section walks through both old tools the way the author would use them in the annual-return season: start the program, point it at the folder that holds a batch of clients' GST files, check which clients are complete, organise the files and build each client's ITC and Sales reports, then refresh those reports' Power Query (Excel's data-import tool, which re-reads the client's files into the report) and collect the figures into one consolidated workbook. For each project it gives the clicks in order (6.2 for A, 6.3 for B) and what each screen shows (6.4). Then it lists every difference the author would notice in one table (6.5), and runs the same input mistakes through both projects (6.6). Screen items that are GST rules are collected in 6.7 and only described. 6.8 maps every screen finding of the two audits to the other project (with an index of the feature map, 6.8.3), 6.9 compares the screens with the author's Standards 12 and 13 (facts only), 6.10 lists corrections, and 6.11 lists questions for the author.

### 6.0 How to read this section

**The two projects.** A = `goyal-sidhant/AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026): a Tkinter window (Tkinter, "Tk" for short, is the window toolkit built into Python) that calls itself "GST File Organizer v3.0". B = `goyal-sidhant/AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026): a PyQt5 window (PyQt5, "Qt" for short, is a separate add-on window toolkit) that calls itself "v4.0". The names mislead: A is not the older project and B is not the newer or better one. Both share their first 12 commits (a commit is one saved, dated change to the code, with a short message) and split after `47a1492` (30-10-2025). All of B's screen code was written in its own five commits (07-02-2026 to 09-02-2026); A rebuilt Step 1 and added Step 4 in July 2026.

**Evidence marks.**

- `A:path:line` and `B:path:line` point into A at `bf2f1dd` and B at `b0e7d0e`; `path:line@commit` points into older code.
- **READ** = established by reading the code of both projects.
- **RAN (C6)** = run for this section (details below). **RAN (A audit §10.x)** / **RAN (B audit §10.x)** = run by that audit's Section 10 and re-read here. The A audit is `/home/user/Annual_Statement_2.0/.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv1.md`; the B audit's Section 10 is `/root/merge-scratch/audit/B/S10.md`.
- **HIGH / MEDIUM / LOW** = the confidence of an inference that was not run.
- **D1 … D66** number the rows of the difference table in 6.5; **G1 … G9** the GST-rule items in 6.7; other sections can cite them as "§6 D15".

**GST rules.** A "GST rule" is anything that decides a tax figure; which return, table, sheet or cell a figure belongs to; a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. When in doubt, a point is treated as one. Such items are tagged **[GST RULE]** and only described: what each project does, one worked example through both, and when each version came in. No version is recommended. **⚠** marks values that change over time (file-name conventions, return and table names, template cell positions, state codes).

**"Works better" remarks** appear only for technical behaviour: crashes, safety of Excel and files, speed, and how clear the screens and messages are. They never judge a GST rule.

**How the runs were done (RAN (C6)).** Everything ran in `/root/merge-scratch/run/C6/` with `python -B` (no compiled files written), from fresh copies of both snapshots (`git archive` - a clean export of a project's code as it stood at its snapshot - of `A/main` and `B/main`, checked identical to `/root/merge-scratch/A-copy` and `/root/merge-scratch/B-copy`). The home folder was pointed into the run folder, so each program's settings file was written there.

- **One placeholder data set for both** (`scripts/mkdata.py`; invented names, no client data): one source folder of 20 files - ClientA (Maharashtra, all six file types, two GSTR-3B months), ClientB (Gujarat, no IMS Reco and no Sales Reco file), ClientC (Kerala, all six types with two Sales files), one file with an unrecognised name (`random_notes.xlsx`) and one file too small to be a real workbook; plus two placeholder templates that have a "Links" sheet.
- **A** (`scripts/tourA.py`, `scripts/probeA.py`): the real window on Xvfb (a screen that exists only in memory: `xvfb-run -a`), Python 3.14 with Tk 9.0, **without** pywin32 (the Python add-on that lets Python drive Excel; see the word list below) - A starts without it off Windows.
- **B** (`scripts/tourB.py`, `scripts/probeB.py`): the real window off-screen (`QT_QPA_PLATFORM=offscreen`, Qt draws into memory), PyQt5 5.15.11 on the Qt 5.15.19 libraries, with five small **stand-in pywin32 modules** in `/root/merge-scratch/run/C6/stubs/` (B will not open its window at all without pywin32; the stand-in Excel always fails with "stub: no Excel COM on Linux").
- **The same path through both**: empty start → look at all four tabs → type the four paths → Scan → tick all clients → Dry Run → Start → finish Step 3 → Step 4 → tick everything → start the refresh → finish → switch dark mode on and look at all four tabs. The probes then ran the same input mistakes through both (6.6).
- **Message boxes** were replaced by recorders (small stand-ins that write down the title and text instead of waiting for a click), so their wording is RAN but their look is not.
- **Outputs**: `out/tourA.json`, `out/tourB.json`, `out/probeA.json`, `out/probeB.json`; screenshots `shots/A01_…png` to `A19_…png` and `shots/B01_…png` to `B18_…png`.
- **Re-run for the check of this section (RAN (VERIFY-C6))**: all four scripts, the B start without stand-ins, and three small extra probes (`damagedA.py`, `measureA.py`, and a no-console import of each `main.py`, which runs its first lines - the console and log set-up - but not the window) were run again in `/root/merge-scratch/run/VERIFY-C6/` on fresh `git archive` copies of both snapshots (checked identical to the A and B copies). Every RAN (C6) fact below came out the same; only the clock times in file and folder names differ (for example `…_231026_1754` / `…_1755` instead of `…_1731` / `…_1732`). Where the re-run changed or added a fact, it is marked RAN (VERIFY-C6).

**Limits.** On Linux both programs build the Step 3 reports with openpyxl (a Python library that writes `.xlsx` files without Excel) instead of Excel, so Steps 1-3 ran end to end, while Step 4 could only show its failure path (A: "pywin32 is not installed…"; B: the stand-in Excel fails). Excel windows, Windows Explorer, the Windows file pickers and the Windows look of message boxes could only be READ. The Windows fonts ("Segoe UI", "Consolas") are replaced by look-alikes; in A's screenshots the emoji in tab titles and buttons are missing (Tk on this machine drew nothing for them), while B's logs show small emoji pictures. Positions, sizes, colours, which controls exist and which are active are reliable.

**Words used in this section** (first use).

| Word | Meaning |
|---|---|
| Window, tab, sub-tab | A window is the framed area on screen. A tab is one of the clickable pages at the top ("Step 1: Setup" …); a sub-tab is a tab inside a tab (A's Step 4 has two). |
| Radio buttons | A group of round buttons of which exactly one can be chosen (for example the three processing modes). |
| Dialog, message box, modal | A dialog is a small window that asks something. A message box is a standard small dialog with an icon, a text and buttons such as OK or Yes/No. Modal = the rest of the program waits until it is answered. |
| File picker | The standard Windows "Open" / "Save" / "Choose folder" dialog. |
| Status bar | The thin strip along the bottom of the window that shows a one-line message. |
| Tooltip | A small yellow pop-up hint that appears when the mouse rests on a control. |
| Console | The black text window from which `python main.py` is typed. |
| pywin32, COM | pywin32 is the Python add-on that lets Python drive Microsoft Excel through COM (Windows' way for one program to remote-control another). |
| Power Query, refresh | Excel's data-import tool. The ITC and Sales templates hold queries; "refresh" re-runs them so the report re-reads the client's files. |
| Alt+A, R, A | The keys both programs press inside Excel to start Data > Refresh All (⚠ Excel's ribbon shortcuts can change between Excel versions). |
| pip | Python's own installer for add-on packages. |
| openpyxl | A Python library that writes `.xlsx` files without Excel. |
| Links sheet | The sheet in each template into which Step 3 writes folder paths and file names for the queries to read. |
| Run folder, client folder, Version folder | `Annual Statement-<DDMMYY HHMM>` (one per processing run), `<client>-<state code>` inside it, and `Version-<DDMMYY HHMM>` inside that. |
| Consolidated workbook | `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, written by Step 4 into the run folder, with the figures read from every refreshed report. |
| Worker thread, "freezes" | A worker thread is a second line of work running beside the window so the window keeps responding. When work runs on the window's own thread instead, the window "freezes" (does not redraw or react) until it ends. |
| Dark mode | A switch that redraws the window in dark colours. |
| Splitter | An invisible divider between two side-by-side panels that the user can drag. |
| Tab dot | B's small coloured circle in front of a tab title. |

### 6.1 At a glance

- **Same four steps, different screens.** Both have one main window with four tabs: Setup, Validation, Processing, and Extract (B: "Extract & Refresh"). A has a blue title strip with a Dark Mode tick; B has a File/View menu bar and coloured dots on the tabs (D5, D7).
- **Start-up.** A checks its packages in the console and offers to install what is missing. B checks nothing and, without pywin32, opens no window at all; its console advice would not fix it (D1; RAN (C6) for B).
- **Step 1.** A has one long scrolling page with a live checklist that shows ✓ / ○ / ✕ for each path as it is typed, and a choice "organise only" or "full pipeline". B has two columns of coloured cards with a static help card and a dark SCAN LOG. In B three of the four "Browse" buttons cannot be seen (white on white) (D12-D15; RAN (C6) screenshots A01, B01). In both, the Scan button is below the visible part of the page at the opening size (D22, RAN (C6)).
- **Step 2.** After a scan, A's clients start unticked and its "Missing" column names the missing return files. B's clients start ticked and its "Missing" column shows only a number. A has a per-client "📁 Name" toggle and a working "EXPORT LIST"; B has "Folder Name" and "Last Run" columns, header sorting, and an "EXPORT LIST" button that does nothing (D29-D35; RAN (C6)).
- **Step 3.** Both show the same "Processing Complete" box. A's log is time-stamped and very detailed, including "[DEBUG]" lines in the Sales part (8 per client whose Sales report is made; the code holds 14 different ones); B's log is shorter, has no time stamps and loses indentation. A's progress bar is pale grey on grey in light mode; B's is green with a percentage (D39-D44; RAN (C6)).
- **After Step 3.** A stays on Step 3 ("organise only") unless "Full pipeline" was chosen. B always moves to Step 4 with the list already filled. Neither starts the Power Query refresh by itself (D44, D14; RAN (C6)).
- **Step 4.** A embeds the old extractor: a folder box, a Scan button, sub-tabs "Setup" and "Processing", one row per client, no Stop button, and at the end the question "Open the consolidated report?". B has one page: a client → version tree that is already filled, a Stop button, no end question, and a "Refreshed (date)" status in green even when every refresh failed (D48-D59; RAN (C6)).
- **Failure reasons in Step 4.** A writes the reason per failed report on screen ("⚠️ ITC Report - <reason>"). B writes only "Result: ITC: FAIL, Sales: FAIL"; the reason goes to the program's log file (D56; RAN (C6)).
- **Dark mode.** A recolours every screen, with some unreadable spots. B darkens only the frame, Step 3 and Step 4; Steps 1-2 stay light, and the Step 3/4 status lines become dark grey on black (D62-D63; RAN (C6) screenshots A16-A19, B15-B18).
- **Locking and closing.** Neither locks the other buttons while a job runs. A asks "Stop and exit?" when the window is closed during Step 3; B never asks (D10, D47).
- **In both**, Step 4 opens Excel visibly, types Alt+A, R, A into it, and prints every extracted figure in the on-screen log (D60, D61; READ, same code).

### 6.2 Project A, click by click: from start-up to the consolidated workbook

This is the path the code expects, with the default workflow "organise only" (READ), confirmed on screen where marked RAN (C6).

| # | What the author does | What he sees, and what happens next | Evidence |
|---|---|---|---|
| 1 | Opens a console in the program folder and types `python main.py` (or `py main.py`, or double-clicks `main.py`). There is no `.bat` launcher or shortcut. | The console prints a line of "=", "🚀 GST File Organizer v3.0 - Production Ready", "Checking requirements..." and one line per package ("[OK     ] openpyxl …"; for pywin32 "[MISSING]" on Windows, "[not found]" elsewhere). If a package is missing: a Yes/No box "Install missing packages?"; "Yes" runs pip in an "Installing…" window and restarts the program; "No" continues only if nothing essential is missing. (pip is Python's own package installer.) On Windows pywin32 is essential, so without it the program says so and ends with "Press Enter to exit...". The console stays open behind the window and keeps receiving log lines. | READ `A:main.py:57-134`, `A:utils/requirements_check.py:93-133`, `A:utils/dependency_installer.py:88-158`; RAN (A audit §10.3) |
| 2 | (nothing) | The main window "GST File Organizer v3.0" opens at 1200×800, centred on the screen, on "📁 Step 1: Setup". Across the top is a blue title strip "🏢 GST FILE ORGANIZER & REPORT GENERATOR" / "Organize files and generate Excel reports automatically", with a "🌙 Dark Mode" tick at the top right. Below it are four tabs, and along the bottom a charcoal status bar "💡 Step 1: Setup … v3.0 \| Production Ready". The last session's four paths, workflow, processing mode, include-name tick, length and dark mode come back from `~/.gst_organizer_cache.json`. Step 2 is empty (scan results are not remembered). The Step 3 log already holds three lines, including "⚠️ IMPORTANT: Templates must have 'Links' sheet for reports to work!". | RAN (C6, shot A01: 1200×800, status "💡 Step 1: Setup", the three log lines); READ `A:gui/main_window.py:57-77, 131-149, 373-398` |
| 3 | Reads the CHECKLIST card at the top of Step 1. | Four tinted rows: Source folder, ITC template, Sales template, Target folder. Each shows "○ Not selected yet" (grey), "✓ <path or file name>" (green), or "✕ Folder not found" / "File not found" / "Not an Excel template" (red). A pill at the top right reads "0 of 4 items set" … "Ready - all 4 items set". The rows are re-checked 350 ms after any change. | RAN (C6, A01: "0 of 4 items set"; A05, after typing the paths while the target folder did not yet exist: three ✓, the target row "✕ Folder not found", pill "3 of 4 items set"); READ `A:gui/tabs/setup_tab.py:282-314, 506-519, 578-608` |
| 4 | In "WHAT SHOULD THIS RUN DO?" leaves "📁 Organise files & create reports" chosen, or picks "🔄 Full pipeline". | The choice is saved at once. The footnote says: "Step 4 stays available either way - this only decides whether Step 3 continues into it automatically." | READ `A:gui/tabs/setup_tab.py:316-351`; `A:utils/constants.py:163-172` |
| 5 | Scrolls down to "📂 FOLDERS & TEMPLATES" and clicks the four "Browse…" buttons, or types or pastes the paths. | The Windows pickers "Select Source Folder", "Select ITC Template" (file type "Excel Template" = `*.xltx *.xlsx` only), "Select Sales Template" and "Select Target Folder" open. After each choice the status bar shows "Source: …", "ITC template selected", "Sales template selected" or "Target: …", and the settings are saved. A typed path is saved only when something else saves. A red line under a box repeats a ✕ problem. | READ `A:gui/handlers/file_handler.py:19-55`, `A:gui/tabs/setup_tab.py:353-391`; RAN (C6, A06b: red "Folder not found" under the target box) |
| 6 | In "⚙️ PROCESSING MODE" keeps "🆕 Fresh Run" (or chooses "🔄 Re-Run" / "▶️ Resume"). Optionally ticks "Include client name in Level 4 folder names" and sets "Max client folder name length" with the arrows. | Choosing a mode does not save by itself (the mode is written only at the next save of any other kind); the tick and the length arrows save at once. | READ `A:gui/tabs/setup_tab.py:393-451`, `A:gui/main_window.py:358-371`; A-R116 |
| 7 | Scrolls further and clicks "🔍  Scan Files". It is greyed out until the Source row shows ✓. | The window does not react while the source folder is read, because the scan runs on the window's own thread. If at least one client is recognised, the window jumps to "✅ Step 2: Validation" and the status bar reads "💡 Step 2: Validation". If nothing is recognised, it stays on Step 1 and, on the screen in view, only the status bar says "💡 No valid GST files found" (the Step 3 log, which is not in view, also gets "🔍 Starting file scan..." and "⚠️ No valid GST files found"). | RAN (C6: A01 Scan "disabled" with no source; A07 the jump; probeA P3 an empty folder: no box, still Step 1, "💡 No valid GST files found"); READ `A:gui/handlers/file_handler.py:57-95` (log lines `:64`, `:91`) |
| 8 | Reads "📊 SCAN SUMMARY" and the client table on Step 2. | A plain-text summary: "Total Excel Files Found: 19", "Successfully Parsed: 18", "Unparsed/Variations: 1", … "🔍 FILE VARIATIONS: 1 files", "💾 TOTAL DATA SIZE: 102.0 KB". The table rows are all unticked and tinted pale green (complete) or pale amber (any other status): ClientA "✅ Complete"; ClientB "⚠️ Missing 2 files" with Missing "IMS Reco, Sales Reco"; ClientC "⚠️ Has duplicates" with Extra 1. At 1200×800 the Missing, Extra and "📁 Name" columns need the horizontal scroll bar. | RAN (C6, A07) |
| 9 | Ticks clients: clicks the "✓" column, uses ↑/↓ and Space, or clicks "☑️ Select All", "☐ Clear All" or "✅ Complete Only". Optionally clicks a client's "📁 Name" cell. | The status bar shows "Selected all 3 clients" (or similar). The Name cell may raise "Long Client Name Warning" or "Global Setting Disabled". | RAN (C6, A08: "💡 Selected all 3 clients"); READ `A:gui/handlers/client_handler.py:67-209` |
| 10 | Optional: clicks "🧪 DRY RUN". | Both template paths must be filled in and point to files that exist, a target must be filled in, and at least one client must be ticked; otherwise an "Error" box appears (for no ticked client: a "Warning" box). The window jumps to Step 3 and writes a preview into the dark log, every line time-stamped. | RAN (C6, A09; probeA P5); READ `A:gui/handlers/processing_handler.py:38-46`, `A:gui/handlers/file_handler.py:132-151` |
| 11 | Optional: clicks "📋 EXPORT LIST". | A save dialog; a text file of the ticked clients; the box "Export Complete"; the file then opens in Notepad (Windows). | READ `A:gui/handlers/client_handler.py:225-257`; RAN (A audit §10.6.6) |
| 12 | Clicks "🚀 START" on Step 2 (or "🚀 START PROCESSING" on Step 3). | The checks run in this order: templates and target ("Error" boxes), at least one ticked client ("Warning"), and long names (a Yes/No box, only when the include-name tick is on). The window then jumps to Step 3: START is greyed, STOP becomes active, and the log shows "🚀 ════════ PROCESSING STARTED ════════". | RAN (C6: during the run the tab was Step 3, START "disabled", STOP "normal"; probeA P6-P9); READ `A:gui/handlers/processing_handler.py:69-134` |
| 13 | Watches Step 3. | A progress bar (pale grey on grey in light mode), "📊 Progress: x%", and "Client i of n: <client> \| ETA: …". The log shows, per client, the three folder levels, one line per copied file, the ITC report lines, the Sales report lines (with "[DEBUG]" lines: 8 for a client whose Sales report is made, at most 9 on the error paths; the code holds 14 different ones), and "✅ Completed <client>". At the end come the summary workbook lines and the totals. | RAN (C6, A10 log: "Level 2: ClientA-MH", "📁 GSTR3B-ClientA-Maharashtra-Apr.xlsx → GSTR-3B Export folder", "[DEBUG] Template data keys: […]", "✅ Processed: 3/3 clients", "📊 Reports created: 6"; RAN (VERIFY-C6): 24 "[DEBUG]" lines for the three clients, 8 each); READ `A:gui/handlers/processing_handler.py:331-411` (the 14 "[DEBUG]" texts at `:333, 342, 353, 362, 363, 367, 371, 374, 391, 393, 398, 403, 404, 411`) |
| 14 | Clicks OK on the box "Processing Complete". | The box says "Successfully processed 3/3 clients", a blank line, "Files organized in:" and the target folder. After OK the new run folder opens in Windows Explorer, and the log ends "📂 Opened output folder" and "ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query." The window stays on Step 3. | RAN (C6, A10: box text, sent from the worker thread "Thread-1 (process_files_thread)"; the two log lines) |
| 15 | Clicks the "🔄 Step 4: Extract" tab. | An explanation strip ("🔄  Step 4 — Refresh Power Query & extract data …"), then sub-tabs "Setup" (open) and "Processing". The "Folder Path" box already holds the Step 1 target; the client list is empty; "🚀 Start Processing" is greyed. | RAN (C6, A11) |
| 16 | Clicks "🔍 Scan" (or "📁 Browse", which scans at once after the choice). | The newest `Annual Statement-*` folder inside that folder, by date modified, is used. One row per client folder appears ("ClientC-KL", "ClientA-MH", "ClientB-GJ": the order the folder listing returned), with the ITC and Sales ticks pre-ticked where the report exists and "Never refreshed". Start becomes active. "Found 3 clients" appears only on the "Processing" sub-tab. | RAN (C6, A12); READ `A:power_query_extractor/gui/extractor_window.py:611-668` |
| 17 | Optionally changes ticks ("✓ Select All", "✗ Deselect All", "ITC Only", "Sales Only"), "Power Query Wait Time" (5-60), "Refreshed File Suffix" or "Skip Refresh". | None of these is saved. | READ `A:power_query_extractor/gui/extractor_window.py:249-403`; A-R119 |
| 18 | Clicks "🚀 Start Processing". | The "Processing" sub-tab opens, the log is cleared and Start is greyed. There is no Stop button. On Windows with pywin32, for each ticked report Excel opens visible and maximised and the program types Alt+A, R, A into it; after each client the log shows "✅ ITC Report - Refreshed successfully" or "⚠️ ITC Report - <reason>". Without pywin32 (this machine) each client showed "❌ pywin32 is not installed for this Python, …" and two amber "⚠️ … Report - …" lines. | RAN (C6, A13-A14: failure path); READ `A:power_query_extractor/gui/extractor_window.py:845-966`, `A:power_query_extractor/core/report_processor.py:209-415` |
| 19 | Answers the box "Complete" - "Processing complete!" / "Open the consolidated report?" (Yes / No). | "Yes" opens `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` (saved in the `Annual Statement-…` folder) in Excel. The progress label goes back to "Ready", the bar empties, and the "Last Refresh Status" texts are re-read. | RAN (C6: box recorded; log "📄 Report saved: PQ_Extraction_Report_20261023_1731.xlsx", "📂 Location: …/Annual Statement-231026 1731"); READ `A:power_query_extractor/gui/extractor_window.py:941-986` |
| 20 | Closes the window (× in the title bar). | Only while Step 3 runs: the box "Processing Active" - "Processing is still running. Stop and exit?" (Yes / No). Otherwise the window closes at once: nothing is asked while Step 4 runs, and closing does not save the settings. | READ `A:gui/main_window.py:344-355` |

**Variants in A.**

| Variant | What changes | Evidence |
|---|---|---|
| "🔄 Full pipeline" chosen in step 4 | After step 14 the program itself switches to Step 4, fills its folder box, scans, and logs "✅ Step 4 ready - review the client list and press Process to refresh." The refresh still waits for "🚀 Start Processing" (step 18). | RAN (C6, probeA P12: tab 4, three rows, Start active, that log line) |
| Refresh only (reports made earlier) | Start the program, click the Step 4 tab, Browse or Scan, then steps 17-19. | READ |
| Standalone extractor `python launch_extractor.py` | A separate window "Power Query Extractor" (1000×700) with a blue banner and a grey status bar. It takes the target folder from the settings file and scans it at once. It runs no start-up checks, asks nothing on close, and writes its own `pq_extractor.log`. | READ `A:launch_extractor.py:12-27`, `A:power_query_extractor/gui/extractor_window.py:1032-1049`; RAN (A audit §10.9) |
| A later day | The settings come back, but Step 2 is empty: Scan again before ticking clients. | RAN (A audit §10.13.3) |

### 6.3 Project B, click by click: from start-up to the consolidated workbook

| # | What the author does | What he sees, and what happens next | Evidence |
|---|---|---|---|
| 1 | Opens a console in the program folder and types `python main.py`. There is no `.bat` launcher or shortcut. | Without pywin32 no window opens. The console shows "… - root - ERROR - Import Error: No module named 'pythoncom'", "Missing module: No module named 'pythoncom'", "Please install required packages:" and "pip install PyQt5 openpyxl"; the program ends with exit code 1 (the number a program hands back when it stops; 1 = failed). With pywin32 the window opens; there is no requirements report, and log lines also go to the console. | RAN (C6, `cwdB_nostub`: exactly these lines, exit 1); READ `B:main.py:11-59` |
| 2 | (nothing) | The window "GST File Organizer v4.0" opens at 1400×900 (minimum 1200×800), always on "Step 1: Setup". There is a light menu bar "File", "View"; tabs "Step 1: Setup", "Step 2: Validation", "Step 3: Processing", "Step 4: Extract & Refresh"; and a status bar "Ready". From `~/.gst_organizer/settings.json` come the four paths, the mode, the include-name tick, the max length, the three Step 4 options and dark mode. Before the window appears, the Step 4 list is built from the saved target folder on disk (with no saved target: "No target folder configured."). | RAN (C6, B01, B04); READ `B:gui/pyqt_main_window.py:35-118, 174-190` |
| 3 | Types or pastes the source folder path. The "Browse" button beside the box is invisible; clicking the blank space just right of the box opens "Select Source Folder". | Nothing is checked yet. | RAN (C6, screenshot B01: blank white space where the Source, ITC and Sales "Browse" buttons are); READ `B:gui/views/setup_view.py:200-208, 605-608`; cause `B:gui/widgets/colored_section.py:62` (B audit §10.4) |
| 4 | Types or pastes the ITC and Sales template paths (the pickers sit behind the same invisible buttons, file type "Excel Files (`*.xlsx *.xlsm`)"). | Nothing is checked yet. | READ `B:gui/views/setup_view.py:226-256, 610-622` |
| 5 | Clicks the red "BROWSE" beside the target box, or types the path. | The folder picker "Select Target Folder" opens. | READ `B:gui/views/setup_view.py:297-324, 624-627` |
| 6 | Chooses a mode radio button. Optionally ticks "Include client name in brackets for Level 4 folders", which un-greys "Max length". | Nothing is saved yet. | READ `B:gui/views/setup_view.py:334-388` |
| 7 | Scrolls the left column down and clicks "SCAN FILES". At 1400×900 it is below the visible part. | The settings are saved. Both help cards fold to "> INSTRUCTIONS" and "> EXPECTED FILE NAMES / MODES". The SCAN LOG fills: "Scanning: …", "Scan started...", "Finding Excel files...", "Found 20 files, processing...", "[1/20] …" … "[20/20] random_notes.xlsx", then a "SCAN RESULTS" block. The status bar shows "Scanning... P% (i/N)", then "Scan complete: 3 clients, 19 files, 102.0 KB". The window jumps to Step 2; the Step 1 tab gets a green dot and Step 2 an orange one; "RE-SCAN" becomes active. An empty folder also jumps to Step 2, with "0 clients (0 complete)". | RAN (C6: B06 scroll facts, B07, B07b; probeB P3); READ `B:gui/pyqt_main_window.py:218-316` |
| 8 | Reads Step 2. | A formatted summary: "Scan Summary" (Total files 19, Parsed 18 (94.7%), Unparsed 1, Total size 102.0 KB), "Clients" (… Variations: 2) and "File Types". The table rows are all ticked, with coloured text: ClientA green "Complete"; ClientB orange "Missing 2 files" with Missing "2"; ClientC red-orange "Has duplicates" with Extra "1". Folder Name shows "ClientA-MH", "ClientB-GJ", "ClientC-KL"; Last Run is empty. The count reads "3 clients (1 complete)". | RAN (C6, B07) |
| 9 | Unticks the clients not wanted, or uses "Clear All", "Complete Only" or "Select All". Clicking a column header sorts the table. | The processing order follows the order on screen. | READ `B:gui/views/validation_view.py:321-447`; B audit §10.5 |
| 10 | Optional: clicks "DRY RUN". | The window jumps to Step 3 and writes the preview, without time stamps or indentation. Templates and target are not checked. | RAN (C6, B09; probeB P5) |
| 11 | Optional: clicks "EXPORT LIST". | Nothing happens. | RAN (C6, probeB P14: 0 connected actions, no box, no file) |
| 12 | Clicks "START" (or Step 3's "Start Processing"). | The checks run in this order: "No clients selected.", then "Please select a valid ITC template." / "Please select a valid Sales template." / "Please select a target folder." (all titled "Warning"). The window then jumps to Step 3: Start Processing is greyed and Stop becomes active. | RAN (C6: tab Step 3, Start disabled, Stop enabled; probeB P6-P9); READ `B:gui/pyqt_main_window.py:345-371` |
| 13 | Watches Step 3. | A green band "FILE PROCESSING"; the status line "Client i/N: <client> \| ETA: …"; a green bar with "n%". The log has no time stamps: "🏢 Processing ClientA - Maharashtra", "✓ Created folder structure", "✓ Organized 7 files", "📊 Creating ITC report...", "Output: ITC_Report_ClientA_MH_231026_1732.xlsx", "✅ Success: …", the same for Sales, "✅ Completed ClientA", …, "✅ Summary report created", "🎉 Processing complete: 3/3 clients". | RAN (C6, B10b) |
| 14 | Clicks OK on "Processing Complete". | The same text as A's box. Behind the box, Step 3 already has a green dot and Step 4 an orange one, the Step 4 list has been rebuilt, and the status bar reads "Settings saved" (it overwrote "Processing complete"). After OK the window moves to Step 4 and, on Windows only, Explorer opens the run folder. | RAN (C6, B10: tab 4, dots `#107c10 #ff8c00 #107c10 #ff8c00`, status "Settings saved" while the box was up); READ `B:gui/pyqt_main_window.py:373-410` |
| 15 | On Step 4: reads the list and ticks reports. | A light-blue strip "Source: Annual Statement-231026 1732" with a small "Refresh List" button. A tree with one bold row per client folder ("1 version(s)") and a child row "Version-231026 1732" with empty ITC and Sales squares and "Pending" in orange. The buttons are "ITC Only" (blue), "Sales Only" (orange), "Select All" (green) and "Deselect All" (plain). The options are "Refresh wait time (seconds):" (5-120), "File suffix:" and "Skip refresh (use existing refreshed files)". | RAN (C6, B11-B12: after "Select All" all six squares ticked) |
| 16 | Clicks "Start Extraction" (teal). | Start is greyed and Stop becomes active. The status line reads "Processing 1/3: ClientA-MH / Version-231026 1732"; the bar stays at 0% during the first version; the log lines are blue. Per version the log ends with "Result: ITC: FAIL, Sales: FAIL" in red; the reason ("stub: no Excel COM on Linux") went only to `gst_organizer.log` (12 ERROR lines). On Windows: Excel opens visible and maximised and Alt+A, R, A are typed into it. | RAN (C6, B13-B14; `cwdB/gst_organizer.log`); READ `B:gui/controllers/extraction_controller.py:55-100` |
| 17 | (end of the run) | No box appears. The status bar and the status line say "Extraction complete" and the bar shows 100%; the Step 4 dot turns green. The list is rebuilt with every square empty and "Refreshed (23/10/26 17:32)" in green - although every refresh failed - because an unrefreshed "_Refreshed_" copy was left beside each report. The log ends "Report saved: PQ_Extraction_Report_20261023_1733.xlsx" and "Extraction complete: 3/3 clients processed". The consolidated workbook and `Extraction_Log_20261023_1733.txt` are in the run folder; no button opens them. | RAN (C6, B14; the folder listing shows `ITC_Report_ClientA_MH_231026_1732_Refreshed_231026_1732.xlsx`); READ `B:gui/pyqt_main_window.py:549-561` |
| 18 | Closes the window: File > Exit or ×. | The settings and a session file are saved. Nothing is asked, even during a run. | READ `B:gui/pyqt_main_window.py:107-109, 601-608` |

**Variants in B.**

| Variant | What changes | Evidence |
|---|---|---|
| Refresh only, on another day | `python main.py`, then click the Step 4 tab; the list was already built at start-up from the last saved target. Or use `python launch_extractor.py`, which opens the same window on Step 4 and writes no program log at all. | READ `B:launch_extractor.py:11-24`; RAN (B audit §10.2) |
| Save without scanning | File > Save Settings; status "Settings saved" for 3 seconds. | READ `B:gui/pyqt_main_window.py:101-103, 210-214` |
| After renaming files | "RE-SCAN" (the same action as SCAN FILES), then continue at step 8. | READ `B:gui/pyqt_main_window.py:124-125` |
| First start after using A | B imports A's settings file `~/.gst_organizer_cache.json` once, so A's last paths appear in B. | READ `B:models/cache_manager.py:93-131`; MAP §1 B-F2 |

### 6.4 What each screen shows

#### 6.4.1 Start-up

- **A** shows a console report before any window opens (6.2 step 1). If a package is missing it offers to install it with pip and restart (a Yes/No box, then an "Installing…" window of 760×380 with a scrolling text and no button). The console is required: A's first line re-wraps the console output and fails without one (RAN, A audit §10.3).
- **A** also creates empty `logs` and `temp` folders inside the program folder at every start (when they are not there yet), and writes the program log `gst_organizer.log` into the folder it was started from (RAN (C6): both folders appeared in the run copy after one start; RAN (VERIFY-C6): the same in a fresh copy, with the log file in the start folder; `A:main.py:20-27, 29-43`). B creates no folders there but writes the same log file the same way (`B:main.py:16-29`).
- **B** shows nothing until the window opens. If pywin32 or anything else is missing, the console gets one log line ("… - root - ERROR - Import Error: …") and three short printed lines, and the program ends (RAN (C6); RAN (VERIFY-C6), the same lines and exit code 1). Unlike A, B's first line checks that a console exists (`B:main.py:12-13`), so B's first lines also pass with no console (RAN (VERIFY-C6): importing each `main.py` with no console, which runs those lines but not the window - A stops with "'NoneType' object has no attribute 'buffer'", B passes; the B audit's history section, S12BF H-B1, also ran a simulated start without a console). B's advice names PyQt5 and openpyxl, not pywin32 (`B:main.py:51-53`).

#### 6.4.2 The window frame

| Part | A (RAN (C6), A01) | B (RAN (C6), B01) |
|---|---|---|
| Title | "GST File Organizer v3.0" (`A:utils/constants.py:132-137`) | "GST File Organizer v4.0" (`B:gui/pyqt_main_window.py:71`) |
| Size | 1200×800, minimum 1000×700, centred | 1400×900, minimum 1200×800, placed by Windows |
| Top | Blue title strip, 70 px: big white title, pale-blue subtitle, "🌙 Dark Mode" tick | Menu bar "File" (Save Settings, Exit) and "View" (Dark Mode) |
| Tabs | Grey-beige tabs "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing", "🔄 Step 4: Extract" | Grey tabs, the selected one white with a blue underline: "Step 1: Setup" … "Step 4: Extract & Refresh"; a green or orange dot appears in front of a tab after a step has run |
| Bottom | Charcoal status bar: "💡 <message>" left, "v3.0 \| Production Ready" right | Plain grey status bar with a left-aligned message ("Ready") |

#### 6.4.3 Step 1 - Setup

- **A (RAN (C6), shots A01, A05, A06b, A06).** One column that scrolls as a whole: 1,521 px tall in a 640-px view at 1200×800, so about the top 40 % shows at first. It holds seven blocks in this order:
  1. a header card ("📁  Step 1 — Setup", pill "N of 4 items set");
  2. the CHECKLIST card;
  3. "WHAT SHOULD THIS RUN DO?" (two radio buttons and a footnote);
  4. a blue-banded "📂 FOLDERS & TEMPLATES" box with four labelled path boxes, each with a grey "Browse…" button;
  5. a teal-banded "⚙️ PROCESSING MODE" box: three radio buttons with a grey description line each, the tick "Include client name in Level 4 folder names" with its grey note "Overrides the per-client setting on the next tab.", and "Max client folder name length: [35] (default 35 · names over 35 are shortened)";
  6. an actions card with grey "Scan Files" and "Re-scan" buttons over a coloured strip ("Select a valid source folder to scan" in red, "Ready to scan · still needed for Step 3: Target folder" in amber, or "Ready to scan, and to process afterwards" in green);
  7. a purple "📝  EXPECTED FILE NAMES" card listing six types and patterns (6.7 G1).

  Tooltips appear on every control after the mouse rests 450 ms (A audit §10.5.11). The page mixes two looks: thin-bordered cards with a coloured left stripe, and older boxes with a thick black border and a coloured band.
- **B (RAN (C6), shots B01, B05, B06, B07b).** Two columns separated by a draggable splitter.
  - The left column scrolls: 1,070 px of content in an 815-px view, with SCAN FILES at 989 px. Top to bottom it holds:
    1. a white banner with the title "Welcome to GST File Organizer!" and two lines of welcome text; the title and the two-line text each sit inside their own blue rounded box;
    2. a green-banded card "SOURCE FOLDER" with one text box and a blank space where its Browse button is;
    3. a blue-banded "EXCEL TEMPLATES" card ("Select your Excel template files (.xlsx or .xlsm)") with two captioned boxes and two blank spaces;
    4. a red-orange-banded "TARGET FOLDER - IMPORTANT!" card with a pink strip "This is where your organized files will be created!", a text box with a thick red outline and a solid red "BROWSE" button;
    5. a teal-banded "PROCESSING MODE" card: three radio buttons; the tick "Include client name in brackets for Level 4 folders"; "Max length: [35]" (greyed while the tick is empty); and an orange italic note "Names >10 chars will show warning. This overrides individual client settings if checked.";
    6. a green-outlined box "READY TO SCAN?" with a solid green "SCAN FILES" and a grey "RE-SCAN" (grey because it is switched off until the first scan; orange once switched on, `B:gui/views/setup_view.py:439-456`).
  - The right column holds a light-blue card "v  INSTRUCTIONS" (seven numbered lines), a lilac card "v  EXPECTED FILE NAMES / MODES" (six patterns, a line, and three mode lines), and a near-black "SCAN LOG" that fills the rest.
  - There are no tooltips.

#### 6.4.4 Step 2 - Validation

- **A (RAN (C6), shot A07).**
  - The left column is a fixed 450 px wide. It holds:
    - an orange "📋 VALIDATION INSTRUCTIONS" box that folds with ▼/▶, with five orange bullets;
    - a blue-banded "📊 SCAN SUMMARY" box with a plain-text report in a fixed-width font; the box is a fixed 200 px high (`A:gui/tabs/validation_tab.py:83-85`), so only part of the report fits and the rest needs scrolling (RAN (VERIFY-C6), `measureA.py`: 15 of its 25 lines with this machine's look-alike font; the A audit §10.6 counted 13 - the number depends on the font, so on Windows it is not sure; how to check: run a scan on an office PC and count the lines visible in the box);
    - "🎯 READY TO PROCESS?" with three flat buttons: orange "🧪 DRY RUN", teal "📋 EXPORT LIST", green "🚀 START".
  - The right column holds a green-banded "👥 CLIENT SELECTION" box with the hint "⌨️ Use ↑↓ arrows + SPACE to select, or click with mouse", three buttons (green "☑️ Select All", charcoal "☐ Clear All", blue "✅ Complete Only") and the table. The table rows are tinted, the first row is highlighted slate-blue, and the columns past "Files" need the horizontal scroll bar.
- **B (RAN (C6), shot B07).**
  - A draggable splitter; the left column holds:
    - an orange "VALIDATION INSTRUCTIONS" box whose title and five lines each sit in their own rounded orange box (it does not fold);
    - a blue-banded "SCAN SUMMARY" card with a formatted report that has blue, green and orange headings (the "File Types" part needs the box's scroll bar);
    - "READY TO PROCESS?" with three solid buttons: orange "DRY RUN", teal "EXPORT LIST", green "START".
  - The right column holds a green panel "CLIENT SELECTION" with a light-blue hint "Use checkboxes to select clients, or use the buttons below", three buttons ("Select All", "Clear All", "Complete Only"), the count "3 clients (1 complete)" in dark blue, and the table. The table has a dark-blue header "Selec | Client | State | Status | Files | Missing | Extra | Folder Name | Last Ru" (the first and last headings are cut off), ticked squares, coloured row text and a 44-px horizontal scroll (B audit §10.5).

#### 6.4.5 Step 3 - Processing

- **A (RAN (C6), shots A09, A10).**
  - The top card has a blue band "📊 PROCESSING PROGRESS", a full-width progress bar, "⏳ Ready to start..." (later "📊 Progress: 100.0%"), a blue operation line ("Processing complete \| Total time: 0m 0s"), a green "🚀 START PROCESSING" and an orange-red "⏹ STOP" whose text is greyed while no run is active.
  - The bottom card has a green band "📝 PROCESSING LOG" over a near-black log with time-stamped coloured lines.
  - In light mode the bar's filled part is pale grey on a grey trough, so a full bar looks almost like an empty one.
- **B (RAN (C6), shots B09, B10b).**
  - A full-width green band "FILE PROCESSING", the status line ("Ready to process", later "Complete \| 0m 0s"), a green bar with "0%" … "100%" printed in it, a green "Start Processing", a red "Stop" (just as bright while inactive) and "Clear Log" at the far right.
  - Below, a near-black log fills the rest. It has no time stamps, and the dry-run preview stays above the run's lines.

#### 6.4.6 Step 4 - Extract

- **A (RAN (C6), shots A11-A15).**
  - A white explanation strip with a blue heading and two grey lines (cut off at the 1000-px minimum width, A audit §10.8.7).
  - A small notebook with sub-tabs "Setup" and "Processing".
  - Setup, left column:
    - "Target Folder" ("Folder Path:", box, grey "📁 Browse", green "🔍 Scan");
    - "Processing Options" ("Power Query Wait Time: [10] seconds (Min: 5, Max: 60)", "Refreshed File Suffix: [_Refreshed_{timestamp}]", "Use {timestamp} for datetime", the two-line tick "Skip Refresh / (Use existing refreshed files)");
    - a wide blue "🚀 Start Processing" bar.
  - Setup, right column: "Select Clients & Reports" with four small buttons (teal "✓ Select All", grey "✗ Deselect All", blue "ITC Only", blue "Sales Only"), the headings "Client Name   ITC   Sales   Last Refresh Status", and one row per client. The ticks sit to the right of their headings.
  - Processing: "Processing Progress" ("Ready to process", a bar) and a light "Processing Log" with time-stamped lines (info black, success green, warning amber, error red).
- **B (RAN (C6), shots B04, B11-B14).**
  - One page: a full-width teal band "EXTRACT & REFRESH".
  - A teal-outlined "Extraction Options" box (wait time, file suffix, skip refresh) beside a wider "Clients" box (strip, Refresh List, four coloured buttons, the tree "Name \| ITC \| Sales \| Refresh Status").
  - Under them: a teal "Start Extraction", a red "Stop", "Clear Log" at the far right, a teal progress bar with "%", the status line ("Ready for extraction"), and a near-black log in the bottom third.

#### 6.4.7 Dialogs, message boxes and file pickers

- **A**:
  - 25 message-box calls in the code (wording RAN for all 25 in the A audit §10.11; ten of them were met again in 6.6, RAN (C6)).
  - Six file and folder pickers (`A:gui/handlers/file_handler.py:21, 29, 40, 51`; `A:gui/handlers/client_handler.py:233`; `A:power_query_extractor/gui/extractor_window.py:603`).
  - Two start-up dialogs ("Install missing packages?", "Installation failed") and the "Installing…" window.
  - The Step 3 "Processing Complete" box is raised from the worker thread; until it is answered, Start stays greyed and the folder does not open (RAN (C6): thread "Thread-1 (process_files_thread)").
- **B**:
  - 14 message-box calls, all in `B:gui/pyqt_main_window.py` (13 different texts), all with a single OK button (READ, grep; B audit §10.8).
  - Four pickers (`B:gui/views/setup_view.py:606, 611, 618, 625`).
  - No Yes/No question anywhere.
- **Both**: no "About" box, no settings dialog, no drop-down lists. Pickers never start in the folder already chosen, except A's Step 4 "📁 Browse" (A audit §10.11.3; B audit §10.8).

#### 6.4.8 Dark mode

- **A (RAN (C6), shots A16-A19).** Switched on with the tick in the title strip.
  - Everything turns charcoal `#2b2b2b`: the blue title strip and the coloured bands are gone, and every coloured button becomes the same grey `#404040`.
  - Kept or changed: the Step 1 checklist keeps its own dark colours (dark green rows, bright-green "Ready - all 4 items set"); the Step 2 rows become dark green and dark brown; the Step 3 log turns mid-grey and the bar turns blue `#0078D4`.
  - Thin light-grey outlines stay round the Step 1 cards.
  - Known defects, RAN in the A audit §10.10.3: Step 2 rows become white-on-pale after a scan in dark mode; Step 4 rows built while dark stay white; Step 4 "info" log lines are black on dark grey.
- **B (RAN (C6), shots B15-B18).** Switched on with View > Dark Mode.
  - Dark: the menu bar, the tab bar, the status bar, the Step 3 and Step 4 page backgrounds and the Step 4 group boxes.
  - Light: Step 1 and Step 2 stay white and pastel (only their scroll bars turn dark; in Step 2 every other table row turns dark).
  - Unreadable: the Step 3 status line ("Complete \| 0m 0s") and the Step 4 status line ("Extraction complete") become dark grey on black; the Step 4 "Skip refresh" square and the tree's squares are hard to see.

### 6.5 Every difference the author would notice, side by side

Kind: **T** = technical or presentation (a "works better" remark is allowed); **G** = [GST RULE], described only, see 6.7. Evidence is READ unless marked.

#### 6.5.1 Start-up and the window frame

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D1 | Starting without pywin32 | Console report of each package. On Windows: says pywin32 is required, offers to install it, refuses to start without it. Off Windows: the install offer appears at every start and, after "No", the window opens without pywin32 (A audit §10.3.2). | No check. On any system the program stops before a window appears, with advice ("pip install PyQt5 openpyxl") that does not name pywin32. | `A:main.py:57-102`; `A:utils/requirements_check.py:93-133`; `B:main.py:49-54`; RAN (C6) for B | T: A tells the user what is missing and how to fix it; B's advice does not fix the cause |
| D2 | Launching only the extractor | `python launch_extractor.py` opens a separate window "Power Query Extractor" (1000×700) that scans the remembered target at once | The same script opens B's full four-tab window, switched to Step 4 | `A:launch_extractor.py:12-27`; `A:power_query_extractor/gui/extractor_window.py:1032-1049`; `B:launch_extractor.py:11-24` | T |
| D3 | Version shown | "GST File Organizer v3.0"; status bar "v3.0 \| Production Ready"; console "v3.0 - Production Ready" | "GST File Organizer v4.0" in the title only | RAN (C6) titles; `A:gui/utils/status_bar.py:29`; `B:gui/pyqt_main_window.py:71` | T: neither label says which version is newer (A has 34 commits of its own after the split, B 5) |
| D4 | Window size | 1200×800, minimum 1000×700, centred. At the minimum the Dark Mode tick and the status bar disappear (RAN, A audit §10.2.2) | 1400×900, minimum 1200×800 - taller than a 1366×768 laptop screen | RAN (C6) sizes; `A:gui/main_window.py:57-70`; `B:gui/pyqt_main_window.py:72-73` | T: on a 768-px-high laptop B cannot fit on screen (MEDIUM, not tried on such a screen); A can be shrunk but then loses two controls |
| D5 | Top of the window | Blue title strip with title, subtitle and Dark Mode tick | Menu bar File / View | RAN (C6) A01, B01; `A:gui/widgets/title_bar.py:17-51`; `B:gui/pyqt_main_window.py:95-118` | T |
| D6 | Tab titles | With emoji: "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing", "🔄 Step 4: Extract" | Plain: "Step 1: Setup" … "Step 4: Extract & Refresh" | RAN (C6) tab titles; `A:gui/tabs/setup_tab.py:161`, `A:gui/tabs/validation_tab.py:21`, `A:gui/tabs/processing_tab.py:20`, `A:gui/tabs/extract_tab.py:50`; `B:gui/pyqt_main_window.py:90-93` | T |
| D7 | Tab status dots | None | Green = a step has run, orange = the next step is ready. They are never reset, Step 2's dot never turns green, and green also follows an empty scan, a run in which every refresh failed, or a stopped run | RAN (C6): after the scan `#107c10, #ff8c00, -, -`; after processing `#107c10, #ff8c00, #107c10, #ff8c00`; after the failed extraction Step 4 green; `B:gui/pyqt_main_window.py:305-306, 391-392, 551, 567-590` | T: B's dots mean "ran", not "succeeded" |
| D8 | Status bar messages | "💡 Step N: <name>" on every tab change (this overwrites the scan result at once), "Source: …" after Browse, "Selected <n> clients"; nothing from Step 3 or Step 4 | "Ready", "Scanning... P% (i/N)", "Scan complete: 3 clients, 19 files, 102.0 KB", "Settings saved" (3 s), "Processing failed", "Extraction complete"; "Processing complete" is overwritten by "Settings saved" at once | RAN (C6): A after the scan "💡 Step 2: Validation", B "Scan complete: …"; B at the completion box "Settings saved"; `A:gui/main_window.py:331-342`; `B:gui/pyqt_main_window.py:214, 235, 242, 308-312, 375-381, 552` | T: B's bar carries results, A's mostly the tab name |
| D9 | Settings remembered, and when they are saved | File `~/.gst_organizer_cache.json`: four paths, workflow, mode, include-name, length, dark mode. Saved after each Browse, the workflow radio, the length arrows, the include-name tick and dark mode; typed paths and the mode are written only when one of these saves happens later; the Step 4 options are never saved; closing does not save | `~/.gst_organizer/settings.json`: four paths, mode, include-name, max length, wait time, suffix, skip refresh, dark mode. Saved at each scan, at the end of processing, by File > Save Settings, on the dark-mode switch and on closing | `A:gui/main_window.py:358-389`; A-R116; `B:gui/pyqt_main_window.py:194-214, 601-608`; RAN (C6) both settings files | T |
| D10 | Closing during a run | Asks "Processing is still running. Stop and exit?" during Step 3 only; asks nothing during Step 4 | Never asks; saves the settings and closes, cutting off any job | `A:gui/main_window.py:344-355`; `B:gui/pyqt_main_window.py:601-608`; RAN (A audit §10.4.1; B audit §10.8) | T: A protects Step 3; neither protects Step 4 |
| D11 | An unexpected error inside a button | Python's error listing goes to the console; the window stays open | The whole window vanishes (the console reports exit code 134); unsaved settings are lost | RAN (A audit §10.3.3; B audit §10.8, e7) | T: B loses the session; neither shows a message on screen |

#### 6.5.2 Step 1 - Setup

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D12 | Layout | One scrolling column of seven cards | Two columns with a draggable splitter: the input cards on the left; help cards and a SCAN LOG on the right | RAN (C6) A01, B01 | T |
| D13 | Live checklist | ✓ / ○ / ✕ per path, re-checked 350 ms after typing; pill "N of 4 items set"; readiness strip; Scan greyed until the source is ✓ | None; paths are checked only when SCAN or START is pressed | RAN (C6) A01, A05; `A:gui/tabs/setup_tab.py:282-314, 506-608`; `B:gui/pyqt_main_window.py:218-225, 352-361` | T: A shows problems before a click. Its check is existence only, so a file passes as a folder (A audit §10.5.3), and a target folder that does not exist yet shows ✕ although Step 3 creates it (RAN (C6), A05 then A10) |
| D14 | "What should this run do?" | Two choices: organise only / full pipeline (the second hands over to Step 4 and scans it, but does not refresh) | No choice; after every processing run B moves to Step 4 with the list filled - in effect A's "full", also without refreshing | RAN (C6) A10, probeA P12, B10; `A:gui/handlers/processing_handler.py:136-175`; `B:gui/pyqt_main_window.py:384, 401` | T |
| D15 | Browse buttons | Four visible grey "Browse…" buttons | Three invisible buttons (white on white; they work if clicked blindly) and one red "BROWSE" | RAN (C6) A01/A06b against B01/B06; cause `B:gui/widgets/colored_section.py:62` | T: B's three buttons cannot be found without being told |
| D16 | Template file types | Picker offers `*.xltx *.xlsx`; the checklist accepts `.xlsx .xltx .xltm .xlsm` | Picker offers `*.xlsx *.xlsm`; the card says "(.xlsx or .xlsm)"; START checks only that something exists at the path | `A:gui/handlers/file_handler.py:31, 42`; `A:gui/utils/setup_validation.py:21`; `B:gui/views/setup_view.py:221, 612, 619`; `B:gui/pyqt_main_window.py:353-358` | T |
| D17 | Where a wrong path is reported | At once: the ✕ row and a red line under the box ("Folder not found", "File not found", "Not an Excel template") | Only in a "Warning" box at SCAN ("Folder not found: <path>") or START ("Please select a valid ITC template.") | RAN (C6) A06b; probeB P2, P8 | T |
| D18 | Mode radio texts and help | "🆕 Fresh Run" / "🔄 Re-Run" / "▶️ Resume", each with a grey description line and a tooltip; the choice is written only at the next save of another kind | The same descriptions in one line ("Fresh - Create new folder structure (recommended for first time)" …); the help card next to them says "Re-Run: Overwrite existing", "Resume: Skip already processed"; saved at the next scan, end of processing, menu save or close | `A:utils/constants.py:174-190`; `A:gui/tabs/setup_tab.py:393-412`; `B:gui/views/setup_view.py:341-359, 510-515` | T: B's two texts contradict each other; in both the Resume wording does not match what the code does (A audit §10.12 T4; B audit §10.4) |
| D19 | Include-client-name tick and its note | "Include client name in Level 4 folder names"; note "Overrides the per-client setting on the next tab." (the per-client setting exists) | "Include client name in brackets for Level 4 folders"; note "Names >10 chars will show warning. This overrides individual client settings if checked." (B has no such warning; its organiser still reads a per-client setting, but no screen in B can set one - MAP §1 A-F10) | `A:gui/tabs/setup_tab.py:417-427`; `B:gui/views/setup_view.py:369, 385`; `B:core/file_organizer.py:117-126` | T: B's note describes features B's screens do not offer |
| D20 | Length box | "Max client folder name length:" always active, arrows 15-100, hint "(default 35 · names over 35 are shortened)"; typed text is accepted unchecked (empty or non-number text breaks saving and Step 3, RAN A audit) | "Max length:" greyed unless the tick is on, but its value is always applied; Qt accepts only numbers 15-100 | RAN (C6) A06b; `A:gui/tabs/setup_tab.py:429-451`; `B:gui/views/setup_view.py:373-382`; `B:gui/pyqt_main_window.py:203` | T: A's box can break a run; B's box looks off while it is on |
| D21 | Scan buttons | Grey "🔍  Scan Files" and "🔄  Re-scan", both greyed until the source is ✓; Re-scan first empties the old results | Green "SCAN FILES" and orange "RE-SCAN", the same action; RE-SCAN greyed until the first scan of the session, then never greyed again | RAN (C6): A01 both "disabled"; B01 SCAN on, RE-SCAN off, B07 RE-SCAN on; `B:gui/pyqt_main_window.py:124-125, 302` | T |
| D22 | Where the Scan button is | Below the visible part: page 1,521 px tall, 640 px visible | Below the visible part: column 1,070 px tall, 815 px visible, SCAN FILES at 989 px | RAN (C6) both | T: the same problem in both |
| D23 | During a scan | The window freezes (the scan runs on the window's own thread); no progress is shown | The window keeps working; the SCAN LOG lists every file ("[i/N] <name>") and the status bar shows the percentage; both help cards fold away | `A:gui/handlers/file_handler.py:62-72`; `B:gui/pyqt_main_window.py:227-242`; RAN (C6) B07b | T: B stays responsive. A scan of 20 local files took 0.007 s (A) and 0.045 s (B) here, RAN (C6); on a network share A freezes for the whole scan (A audit §10.5.8) |
| D24 | Expected-file-names card | Purple card at the bottom: type and pattern with `.xlsx` | Lilac card on the right, folding: six patterns without extension, then three mode lines | RAN (C6) A06, B01 | **G** - see 6.7 G1 ⚠ |
| D25 | Hover help | Tooltips on every Step 1 control | None anywhere | `A:gui/tabs/setup_tab.py` (9 `Tooltip(` calls); grep `setToolTip` in B: 0 | T |

#### 6.5.3 Step 2 - Validation, and the Dry Run

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D26 | Layout | Left column a fixed 450 px | Draggable splitter, larger share on the right | RAN (C6) A07, B07 | T |
| D27 | Instructions | Folding box (▼/▶); emoji legend "✅ = complete, ⚠️ = missing files"; "Select clients using click or keyboard (↑↓ + SPACE)" | Fixed box; colour legend "green = complete, orange = missing files"; "Select clients using checkboxes" | RAN (C6) A07, B07; `A:gui/tabs/validation_tab.py:44-66`; `B:gui/views/validation_view.py:72-115` | T: both end "Click 'Start Processing' when ready" while their button reads "START"; neither legend covers "Has duplicates" |
| D28 | Scan summary | Plain text: "Total Excel Files Found: 19", "Unparsed/Variations: 1", "🔍 FILE VARIATIONS: 1 files" | Formatted: "Total files: 19", "Unparsed: 1", "Variations: 2" | RAN (C6) A07, B07 | T: the same 20 files give "1" in A and "2" in B, because A drops the too-small file without a word while B counts it as a variation (MAP §1 A-F7 / B-F6) |
| D29 | Table columns | ✓ · 🏢 Client Name · 🌏 State · 📊 Status · 📁 Files · ❌ Missing (names of the missing types) · ➕ Extra · 📁 Name | Select · Client · State · Status · Files · Missing (a number) · Extra · Folder Name ("<client>-<state code>") · Last Run | RAN (C6) A07 (ClientB "IMS Reco, Sales Reco"), B07 (ClientB "2") | **G** for the Missing column and the state code - see 6.7 G3, G4 |
| D30 | Ticks after a scan | None ticked | All ticked, complete or not | RAN (C6): A rows "☐"; B check state 2 (ticked) on all three | T: in B every client is processed unless unticked; in A the user must tick on purpose |
| D31 | Row colours and "Complete Only" | Whole-row tint: pale green for "Complete", pale amber for anything else; "Complete Only" ticks rows whose status CONTAINS "Complete" | Text colour: green / orange (contains "Missing") / red-orange (anything else); "Complete Only" ticks rows whose status EQUALS "Complete" | RAN (C6): A tags `complete`, `incomplete`, `incomplete`; B colours `#107c10`, `#ff8c00`, `#d83b01`; `A:gui/handlers/client_handler.py:52-59, 147-159`; `B:gui/views/validation_view.py:387-393, 441-447` | **G** - see 6.7 G2 |
| D32 | Keyboard and sorting | ↑/↓ moves the highlight, Space ticks (the hint above the table says so); headings cannot be clicked | Qt's standard table keys only: ↑/↓ move the cursor and Space ticks or unticks a row while the cursor is in the "Select" column (Space in any other column does nothing; nothing on screen mentions the keyboard - RAN (VERIFY-C6), `keysB.py`); clicking a heading sorts the rows, and processing follows the sorted order (numbers sort as text) | `A:gui/tabs/validation_tab.py:162-169, 241-244`; `B:gui/views/validation_view.py:262, 321-324, 369`; B audit §10.5 | T |
| D33 | Per-client "📁 Name" toggle | Present (with "Long Client Name Warning" and "Global Setting Disabled" boxes); off-screen at 1200×800 and not explained on screen | Absent | RAN (C6) A07; `A:gui/handlers/client_handler.py:175-209` | T |
| D34 | Client count | Status bar "Selected <n> clients" after ticking | Label "3 clients (1 complete)" | RAN (C6) A08, B07 | T |
| D35 | "EXPORT LIST" | Saves the ticked clients to a `.txt` file and opens it | Does nothing (no action connected) | RAN (C6) probeB P14 (0 receivers, that is, no action connected to the button; RAN (VERIFY-C6) the same); `A:gui/handlers/client_handler.py:225-257`; `B:gui/views/validation_view.py:186-200` (no `clicked.connect`) | T: B's button is dead |
| D36 | How the Step 2/3 buttons react | To the mouse press only; the keyboard does nothing; a greyed STOP still reacts | Normal buttons: react on release, ignore clicks when disabled, can be pressed with Tab + Space | `A:gui/tabs/validation_tab.py:118, 130, 142, 190`; `A:gui/tabs/processing_tab.py:74, 87`; `B:gui/views/validation_view.py:184, 217`; Qt default behaviour (HIGH) | T: B follows normal button behaviour |
| D37 | What must be chosen before DRY RUN | Both templates must exist, a target must be typed, and at least one client must be ticked | Only a ticked client | RAN (C6) probeA P5 ("Please select both ITC and Sales template files"), probeB P5 (preview printed) | **G** - see 6.7 G5 |
| D38 | Dry-run text | Every line "[HH:MM:SS]"; "🎯 Mode: fresh", "👥 Selected: 3 clients"; indented; a time stamp alone on the line before each client | No time stamps, no mode, no count; indentation lost | RAN (C6) A09, B09 | T: A's preview carries more information |

#### 6.5.4 Step 3 - Processing

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D39 | Layout | Two boxed cards: progress (blue band) and log (green band) | Green band "FILE PROCESSING"; status line; bar; Start / Stop / Clear Log; log | RAN (C6) A10, B10b | T |
| D40 | Progress display | Bar pale grey on grey in light mode; label "📊 Progress: 100.0%"; operation line "Client i of n: <client> \| ETA: …", "Client 1/3: Copying <file> \| ETA: …", "Processing complete \| Total time: 0m 0s" | Green bar with "n%" inside; status line "Client i/N: <client> \| ETA: …", "Creating summary report...", "Complete \| 0m 0s" | RAN (C6) A10, B10b; `A:gui/handlers/processing_handler.py:236-289, 457-482`; `B:gui/controllers/processing_controller.py:86-90, 185-198` | T: B's bar is readable; A's is hard to read in light mode and shows more detail per file |
| D41 | Stop button | Orange-red with greyed text while inactive (but a click still reaches it); stops before the next client | Bright red whether active or not; stops before the next client; no "stopping…" feedback | RAN (C6) A09/A10, B10b; `B:gui/theme/theme_manager.py:104-111` | T: A shows the inactive state; B does not |
| D42 | Log content | "[HH:MM:SS]" on every line; the three folder levels; one line per copied file; template and output names; "[DEBUG]" lines in the Sales part (8 per client whose Sales report is made - 24 for the three placeholder clients, RAN (VERIFY-C6); the code holds 14 different ones, `A:gui/handlers/processing_handler.py:331-411`); summary-workbook lines; totals "Files organized: 18", "Reports created: 6" | No time stamps; no per-file lines; no DEBUG lines; leading spaces and any "<…>" text lost; "Clear Log" button; never cleared otherwise | RAN (C6) A10, B10b; `A:gui/utils/ui_helpers.py:16-33`; `B:gui/views/processing_view.py:142-157` | T: A's log is a fuller record; B's is easier to scan but drops text |
| D43 | Completion box | "Processing Complete" - "Successfully processed 3/3 clients\n\nFiles organized in:\n<target>" | The same text | RAN (C6) both | T: in both it says "Successfully" also after Stop and when reports failed (A audit §10.7.3; B audit §10.6) |
| D44 | After the box | Explorer opens the run folder (any system); in "organise only" the window stays on Step 3 with an info line; in "full" it moves to Step 4 and scans | The window moves to Step 4 (list already rebuilt); Explorer opens on Windows only | RAN (C6) A10, probeA P12, B10; `A:gui/handlers/processing_handler.py:498-519`; `B:gui/pyqt_main_window.py:394-410` | T |
| D45 | A run that fails as a whole | Log line "💥 Fatal error: <error>" and the box "Processing Error" | Only the box "Processing Error"; nothing in the log; status "Processing failed" | `A:gui/handlers/processing_handler.py:522-525`; `B:gui/pyqt_main_window.py:412-415`; RAN (B audit §10.6) | T: A leaves a trace in the log |
| D46 | A template changed on Step 1 during a run | Used for the next client (A re-reads the paths per client; the log's "Template:" line shows it) | Not used; the paths are fixed when Start is pressed | `A:gui/handlers/processing_handler.py:301, 310, 332`; `B:gui/pyqt_main_window.py:350-371` | **G** - see 6.7 G6 |
| D47 | Controls during a run | Only START is greyed; Scan, Re-scan, Browse, Dry Run, Export and all of Step 4 stay usable (Step 4's Start stays greyed until a Step 4 scan) | Only Start Processing is greyed; SCAN FILES, START, DRY RUN and Start Extraction stay active | RAN (C6): A "scan normal, panel_btn disabled"; B "v_start True, scan True, e_start True" | T: neither locks the other jobs (A-R135; B-G74) |

#### 6.5.5 Step 4 - Extract

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D48 | Layout | Explanation strip and two sub-tabs: the list on "Setup", the log on "Processing" | One page: options, clients, buttons, bar, status line and log all visible | RAN (C6) A11-A15, B11-B14 | T: in B the user sees the list and the log together |
| D49 | Which folder Step 4 looks in | A "Folder Path" box, pre-filled with Step 1's target and following it; "📁 Browse" and "🔍 Scan"; must press Scan (in "organise only") | No folder box; B uses the target last taken from the screen (at a settings save or an accepted START), shown as "Source: <run folder>"; "Refresh List" rebuilds | RAN (C6) A11-A12, B11; `A:gui/tabs/extract_tab.py:108-131`; `B:gui/pyqt_main_window.py:419-504` | T |
| D50 | Which batch and versions are listed | The newest `Annual Statement-*` by date modified; per client only the newest `Version-*` by date modified | The run folder whose name sorts last (names start with the day of the month); per client every Version folder that holds a report, by name Z-A | `A:power_query_extractor/gui/extractor_window.py:622-648`; `B:gui/pyqt_main_window.py:432-464` | T: B can show an older batch (e.g. "Annual Statement-311225 …", a run on 31-12-2025, sorts after "Annual Statement-211026 …", a run on 21-10-2026, B audit §10.7); A's code comment names this trap (`A:power_query_extractor/gui/extractor_window.py:643-647`) |
| D51 | The list | One row per client: ITC and Sales ticks pre-ticked where the report exists, greyed where not; "Never refreshed" or "ITC: <yyyy-mm-dd hh:mm> \| Sales: …" | A tree: client → versions; every square empty after each rebuild; a missing report's square is still drawn empty; "Pending" (orange) or "Refreshed (dd/mm/yy HH:MM)" (green) | RAN (C6) A12, B11-B12; `A:power_query_extractor/gui/extractor_window.py:670-808`; `B:gui/views/extraction_view.py:291-356` | T |
| D52 | Selection buttons | Teal "✓ Select All", grey "✗ Deselect All", blue "ITC Only", blue "Sales Only" | Blue "ITC Only", orange "Sales Only", green "Select All", plain "Deselect All" | RAN (C6) A12, B11 | T |
| D53 | Options | Wait time 5-60 by arrows (typed values not checked), suffix, Skip Refresh; nothing saved | Wait time 5-120, suffix, skip refresh; saved with the settings | `A:power_query_extractor/gui/extractor_window.py:259-325`; `B:gui/views/extraction_view.py:124-142`; `B:gui/pyqt_main_window.py:205-207` | T |
| D54 | Start button | Wide blue "🚀 Start Processing", greyed until a Step 4 scan finds clients | Teal "Start Extraction", always active (a "Warning" box if nothing is ticked) | RAN (C6) A04/A11 greyed, A12 active; B04 active, probeB P13 | T |
| D55 | Stop | None | "Stop" (acts before the next version; looks active when it is not) | `A:power_query_extractor/gui/extractor_window.py` (no "stop" anywhere); `B:gui/views/extraction_view.py:272-276, 392-395` | T: B can stop a long run, though only between versions |
| D56 | Log and failure reasons | On the "Processing" sub-tab; "[HH:MM:SS]"; per report "✅ ITC Report - Refreshed successfully" or "⚠️ ITC Report - <reason>" in amber; the refresher's own lines in black; the user can type into the log box | On the same page; no time stamps; the refresher's lines all blue; per version "Result: ITC: FAIL, Sales: FAIL" in red with no reason - for many failures the reason goes only to `gst_organizer.log` | RAN (C6): A14 ("⚠️ ITC Report - pywin32 is not installed …"), B14 + 12 ERROR lines in the log file; `A:power_query_extractor/gui/extractor_window.py:917-938`; `B:gui/controllers/extraction_controller.py:76-84` | T: A tells the user on screen why a report failed; B does not |
| D57 | End of the run | Box "Complete" - "Processing complete!\n\nOpen the consolidated report?" (Yes opens it in Excel); the bar empties and the label returns to "Ready"; the ticks are kept | No box; "Extraction complete" in the status bar and status line; bar 100%; green dot; list rebuilt with every tick cleared | RAN (C6) A14 + box, B14 | T: A offers the workbook; B leaves the user to find it in Explorer |
| D58 | Status after a refresh that failed | Shows the time of a refreshed copy if one exists (the copy is made before the refresh, so a failed refresh still shows a time - the original report's date, which the copy keeps unless Excel saved it); without pywin32 no copy is made and the row stays "Never refreshed" | "Refreshed (<time>)" in green whenever a "_Refreshed" copy exists - after a failed refresh too | RAN (C6): A15 "Never refreshed" (no copy made without pywin32); B14 "Refreshed (23/10/26 17:32)" after three failures; `A:power_query_extractor/core/report_processor.py:164`; `B:power_query_extractor/core/report_processor.py:120`; `B:gui/pyqt_main_window.py:466-479` | T: B's word "Refreshed" claims a success the program never checked; A's time is also no proof of success |
| D59 | Record of the run | No file in the run folder and no run history, only the on-screen log (A-F25); the refresher's own lines also go into the program log `gst_organizer.log`, as in B (`A:power_query_extractor/core/report_processor.py:53-60`) | `Extraction_Log_<YYYYMMDD_HHMM>.txt` in the run folder; run history | RAN (C6) B folder listing; MAP §1 A-F25 / B-F24, B-F25 | T |
| D60 | What happens on screen during a refresh (Windows) | Excel opens visible and maximised and Alt+A, R, A are typed into the window in front; error-looking pop-ups of any program are clicked away; nothing warns the user | Identical: the steps that drive Excel on screen are the same code (the refresh function differs only by an unreachable `return` and the kinds of error caught - MAP §1 A-F21) | READ `A:power_query_extractor/core/report_processor.py:218-227, 260-292`; `B:power_query_extractor/core/report_processor.py:174-183, 216-248`; MAP §1 A-F21 / B-F19, A-F26 / B-F26 | T (⚠ Excel's shortcut keys can change between Excel versions): the same risk in both: a click elsewhere during a refresh sends the keys to the wrong window (HIGH, A audit §10.8.5; B audit §10.7) |
| D61 | Extracted figures in the on-screen log | Every value read is printed ("✓ <field> (<cell>) = <value>", summary lines) | The same lines; also saved in the Extraction_Log file | READ, same `_extract_data` (MAP §1 A-F23): `A:power_query_extractor/core/report_processor.py:590, 609`, `B:power_query_extractor/core/report_processor.py:541, 560`; each line also goes to the program log through `logger.info` (`A:power_query_extractor/core/report_processor.py:53-60`, `B:power_query_extractor/core/report_processor.py:30-37`); A audit §10.8.4; B audit §10.7 | T (confidentiality: client figures appear on screen and in the program log file `gst_organizer.log` in both - in A's standalone extractor `pq_extractor.log` - and, in B, also in the Extraction_Log text file in the run folder) |

#### 6.5.6 What happens after each step

| After … | A | B | Evidence |
|---|---|---|---|
| A scan that finds clients | Jumps to Step 2; status "💡 Step 2: Validation"; two lines are added to the (unseen) Step 3 log ("🔍 Starting file scan...", "✅ Found 3 clients"); rows unticked | Jumps to Step 2; status "Scan complete: …"; dots green/orange; help cards folded; SCAN LOG filled; RE-SCAN enabled; rows ticked | RAN (C6) A07, A09 log; B07, B07b |
| A scan that finds no client | Stays on Step 1; status "💡 No valid GST files found"; two lines are added to the (unseen) Step 3 log ("🔍 Starting file scan...", "⚠️ No valid GST files found") | Jumps to Step 2 with "0 clients (0 complete)"; status "Scan complete: 0 clients, 0 files, 0.0 B" | RAN (C6) probeA P3, probeB P3 |
| Dry Run | Jumps to Step 3; preview (after the template and target check) | Jumps to Step 3; preview | RAN (C6) |
| Start | Jumps to Step 3; START greyed, STOP active | Jumps to Step 3; Start greyed, Stop active | RAN (C6) |
| End of processing | Box (from the worker thread); then the run folder opens; "organise only": stays on Step 3; "full": Step 4 filled and scanned | List rebuilt, dots set, box; then Step 4; Explorer on Windows only; status "Settings saved" | RAN (C6) A10, probeA P12, B10 |
| Start of the refresh | The Processing sub-tab opens; log cleared; Start greyed | Start greyed, Stop active; the page stays as it is | RAN (C6) A13-A14, B13 |
| End of the refresh | Box "Open the consolidated report?"; bar back to empty, "Ready"; status texts re-read; ticks kept | No box; "Extraction complete"; bar 100%; green dot; list rebuilt, ticks cleared | RAN (C6) A14-A15, B14 |
| Closing | Question only during Step 3 | No question | READ (D10) |

#### 6.5.7 Progress and logs at a glance

| Where | A | B |
|---|---|---|
| Step 1 | No log; status texts are set during a scan but not painted (frozen window) | SCAN LOG (dark), one line per file plus results; status percentage |
| Step 2 | Text summary | Formatted summary |
| Step 3 bar | Pale grey on grey (light); blue (dark) | Green with "n%" |
| Step 3 log | Dark, time-stamped, detailed, never cleared | Dark, no time stamps, shorter, "Clear Log" |
| Step 4 bar | On the Processing sub-tab; empties at the end | On the page; stays at 100% |
| Step 4 log | Light, time-stamped, coloured by level, cleared at each Start, editable | Dark, no time stamps, refresher lines all blue, "Clear Log" |
| Log file | `gst_organizer.log` in the start folder (main window); `pq_extractor.log` in the start folder (standalone extractor, D2) | `gst_organizer.log` in the start folder (only with `main.py`) and `Extraction_Log_…txt` in the run folder |

Evidence: RAN (C6) screenshots; `A:gui/utils/ui_helpers.py:16-40`; `A:main.py:20-27`; `A:power_query_extractor/extractor_main.py:12-19`; `A:power_query_extractor/gui/extractor_window.py:494-571, 1021-1029`; `B:gui/views/setup_view.py:549-579`; `B:gui/views/processing_view.py:114-157`; `B:gui/views/extraction_view.py:57-102, 402-416`.

#### 6.5.8 Colours and dark mode

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D62 | Dark-mode switch | A tick in the title strip (outside the window at the 1000-px minimum width); remembered | View > Dark Mode (a tick item); remembered | RAN (C6); `A:gui/widgets/title_bar.py:27-36`; `B:gui/pyqt_main_window.py:111-118` | T |
| D63 | What dark mode changes | Every screen turns charcoal; coloured buttons all become the same grey (their colour meaning is lost); the bar turns blue; known unreadable spots after scans and in the Step 4 log | Frame, menus, tabs and the Step 3/4 backgrounds turn dark; Steps 1-2 stay light; the Step 3/4 status lines become dark grey on black | RAN (C6) A16-A19, B15-B18 | T: neither is complete; A covers more screens, B leaves two status lines unreadable |

Colours in light mode come from one shared palette in both projects: blue `#0078D4`, green `#107C10`, orange `#FF8C00`, red-orange `#D83B01`, teal `#0099BC`. B's theme file says so itself: "Color palette matching the old Tkinter UI" (`B:gui/theme/theme_manager.py:9`; `A:utils/constants.py:145-156`).

#### 6.5.9 Messages, help and keyboard

| # | Difference | A | B | Evidence | Kind / technical remark |
|---|---|---|---|---|---|
| D64 | Message boxes | 25 calls. Titles: "Error" for eight boxes (input problems, a failed launch of the standalone extractor, a failed Step 4 run), "Warning" (3), "Scan Error", "No Selection", "Processing Complete", "Processing Error", "Complete", "Processing Active", and others. Includes Yes/No questions | 14 calls. Every input problem is titled "Warning"; plus "Scan Error", "Processing Complete", "Processing Error" and "Extraction Error". OK only; no questions | grep counts (READ); RAN (C6) probes; A audit §10.11; B audit §10.8 | T: wording differs for the same situation (6.6) |
| D65 | Keyboard | No shortcuts; Step 2 table ↑/↓ + Space; Step 1 and Step 4 buttons by Space; Step 2 and Step 3 buttons not by keyboard | Alt+F and Alt+V open the menus; buttons by Tab + Space (Qt default); in the Step 2 table ↑/↓ and Space tick rows only while the cursor is in the "Select" column (Qt default, not explained on screen; RAN (VERIFY-C6), D32) | `A:gui/tabs/validation_tab.py:241-244`; `B:gui/pyqt_main_window.py:99, 112` | T |
| D66 | Help texts that disagree with the code | Several: the Resume tooltip, the length tooltip, the target tooltip, the "organise only … no Excel automation" tooltip, the full-pipeline promise, "press Process to refresh", the Links-sheet warning | Several: the mode lines against the radios, the ">10 chars" note, "Click 'Start Processing'", "Start processing in Processing tab", the two-colour legend | A audit §10.12.1 (T1-T25); B audit §10.4, §10.5, §10.14 U7 | T: both need their help texts rewritten against the merged code |

### 6.6 The same mistakes, two sets of messages (RAN (C6), probeA / probeB)

Both projects were given the same placeholder folder and the same mistakes. Texts are exact; "\n" is a line break.

| Situation | A | B |
|---|---|---|
| Scan with an empty source box | The Scan button is greyed, so this cannot happen by a click. Called directly: "Error" - "Please select a source folder containing GST files" | "Warning" - "Please select a source folder first." |
| Scan with a folder that does not exist | Normally prevented by the greyed button. Called directly: "Error" - "Source folder does not exist" | "Warning" - "Folder not found: /nonexistent/folder" |
| Scan of an existing but empty folder | No box; stays on Step 1; status "💡 No valid GST files found" | No box; jumps to Step 2, "0 clients (0 complete)"; status "Scan complete: 0 clients, 0 files, 0.0 B" |
| Dry Run with no templates and no target **[GST RULE]** (6.7 G5) | "Error" - "Please select both ITC and Sales template files"; no preview | No box; the preview is printed |
| Start with no templates | "Error" - "Please select both ITC and Sales template files" | "Warning" - "Please select a valid ITC template." |
| Start with templates but no target | "Error" - "Please select a target folder for organized files" | "Warning" - "Please select a target folder." |
| Start when the ITC template path does not exist | "Error" - "ITC template file not found" | "Warning" - "Please select a valid ITC template." |
| Start with no client ticked | "Warning" - "Please select at least one client to process" | "Warning" - "No clients selected." |
| Step 4 Start with nothing ticked | Before any Step 4 scan the button is greyed; after a scan with every tick removed: "No Selection" - "Please select at least one report to process" | "Warning" - "No versions selected for extraction." |
| Step 4 Scan with an empty folder box | "No Folder" - "Please select a folder first" | (no folder box in B) |
| EXPORT LIST | Works (save dialog, "Export Complete") | Nothing at all |
| End of Step 3 | "Processing Complete" - "Successfully processed 3/3 clients\n\nFiles organized in:\n<target>" | Identical |
| End of Step 4 | "Complete" - "Processing complete!\n\nOpen the consolidated report?" (Yes / No) | No box |

The order of the checks also differs: A checks the templates and the target before the ticked clients (`A:gui/handlers/processing_handler.py:78-84`); B checks the ticked clients first (`B:gui/pyqt_main_window.py:346-361`). Technical: A's titles separate errors from warnings; B's single "Warning" title is used for every input problem.

### 6.7 Screen items that are GST rules [GST RULE] - described only

Nothing below says which version is right. Each item gives what each project does, one worked example through both, and when each version came in. The examples use the placeholder clients of 6.0.

**G1. The expected file names shown on Step 1 ⚠** (A-R16 and the card of A-F31 / B-R13, B-R141; MAP 2.1 item 3; D24)

- *A*: the purple card prints, for each of the six types, the type and the pattern's own example text kept next to the recognition pattern: "GSTR-2B Reco - GSTR-2B-Reco-ClientName-State-Period.xlsx", "IMS Reco - ImsReco-ClientName-State-DDMMYYYY.xlsx", "GSTR-3B Export - GSTR3B-ClientName-State-Month.xlsx", "Sales - Sales-ClientName-State-StartMonth-EndMonth.xlsx", "Sales Reco - SalesReco-ClientName-State-Period.xlsx", "Annual Report - AnnualReport-ClientName-State-Year.xlsx" (`A:gui/tabs/setup_tab.py:477-503`, `A:utils/constants.py:19-59`; RAN (C6) A06).
- *B*: the lilac card prints a hand-typed list: "GSTR-2B-Reco-Client-State-Period", "ImsReco-Client-State-DDMMYYYY", "GSTR3B-Client-State-Month", "Sales-Client-State-Start-End", "SalesReco-Client-State-Period", "AnnualReport-Client-State-Year" - with no extension and no type names (`B:gui/views/setup_view.py:492-503`; RAN (C6) B01). B's code holds the same longer texts as A (`B:utils/constants.py:19-59`) but does not show them.
- *Worked example*: the placeholder file `ImsReco-ClientA-Maharashtra-31032025.xlsx` is recognised as IMS Reco in both (RAN (C6), both scans). A's card shows it as "ImsReco-ClientName-State-DDMMYYYY.xlsx", B's as "ImsReco-Client-State-DDMMYYYY". Neither card mentions that `.xls` is also accepted, that a hyphen inside a client name splits the name, or the "(N)" duplicate forms (MAP 2.1 item 1 for the one place where the two recognition patterns differ).
- *History*:
  - The short list is the shared Tk Setup tab's (`47a1492:gui/tabs/setup_tab.py:362-373`, from `edac6a3`, 31-08-2025, "Initial commit").
  - B's first PyQt5 screen showed the longer texts with `.xlsx` in an HTML help panel (a panel of formatted text, as on a web page) (`c258047:gui/views/setup_view.py:168-173`, 07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"). `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections") went back to the short list.
  - A switched to the longer texts in `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"); its code comment says the card is "read from FILE_PATTERNS so this card can never drift from the regexes the parser actually uses" (regexes = the recognition patterns; what the card actually reads is each pattern's example text, 6.10 item 2).

**G2. How the client status is shown on Step 2, and "Complete Only"** (A-R29 / B-R27; MAP 2.1 item 5; D31)

- *A*: status text with an icon ("✅ Complete", "⚠️ Missing 2 files", "⚠️ Has duplicates"); the whole row is tinted pale green for exactly "Complete", pale amber for anything else. "✅ Complete Only" ticks rows whose status *contains* "Complete".
- *B*: the same status words without icons; text coloured green (exactly "Complete"), orange (contains "Missing") or red-orange (anything else). "Complete Only" ticks rows whose status *equals* "Complete".
- *Worked example* (RAN (C6), A07 and B07): ClientC has all six types and two Sales files. A shows "⚠️ Has duplicates" in the same pale amber as ClientB ("⚠️ Missing 2 files"); B shows ClientC in red-orange and ClientB in orange. "Complete Only" leaves ClientC unticked in both (READ: "⚠️ Has duplicates" does not contain "Complete").
- *History*: the status logic is shared (`edac6a3`; the function is identical in both, MAP 2.1 item 5). A's row tints are the shared Tk code; tinting by each row's own status came with `1c48722` (29-10-2025, shared, "Remove redundant text color handling in client tree display"; before it every row got the "complete" tint in light mode, `edac6a3:gui/handlers/client_handler.py:51-61`; A audit §10.15). B's three text colours and its "equals Complete" test came with `c258047` (07-02-2026; `c258047:gui/views/validation_view.py:141-146, 199`); `7e09e5c` (08-02-2026) rebuilt the screen around them without changing them (`B:gui/views/validation_view.py:387-393, 446`).

**G3. Which missing return files are shown on Step 2 ⚠** (expected types A-R26 / B-R24, the same in both; display A-R138 / B-R134; D29)

- *A*: the "❌ Missing" column names the missing types, joined by ", " ("None" if nothing is missing) (`A:gui/handlers/client_handler.py:31, 41`).
- *B*: the "Missing" column shows the number of missing types; the names are not shown anywhere on screen (`B:gui/views/validation_view.py:377`).
- *Worked example* (RAN (C6)): ClientB has GSTR-2B Reco, GSTR-3B Export, Sales and Annual Report files. Both give "Missing 2 files". A's Missing cell reads "IMS Reco, Sales Reco"; B's reads "2".
- *History*: A's names are the shared Tk code (`47a1492:gui/handlers/client_handler.py:31`, from `edac6a3`). B's count came with `c258047` (07-02-2026).

**G4. The client key with its state code on screen ⚠** (state table A-R21 / B-R21; client key A-R17 / B-R17; where the code is used A-R23 / B-R21, B-R32, B-R51; client-folder name A-R35 / B-R32; made-up code for an unknown state A-R22 / B-R22; round trip MAP 2.1 item 4; D29)

- *A* shows "<client>-<state code>" in the Step 3 log ("Level 2: ClientA-MH") and as the Step 4 row names. Step 2 shows the state *name* and rebuilds the key from the row's text when Dry Run, Start or Export runs (`A:gui/handlers/client_handler.py:161-173`).
- *B* shows the key in Step 2's "Folder Name" column and in the Step 4 tree, and keeps the key inside the row (`B:gui/views/validation_view.py:371, 379`).
- *Worked example* (RAN (C6)): ClientA / Maharashtra → "ClientA-MH", ClientB / Gujarat → "ClientB-GJ", ClientC / Kerala → "ClientC-KL" in both; the report names carry the same code ("ITC_Report_ClientA_MH_231026_1731.xlsx" in A, "…_1732.xlsx" in B). For a number-looking client name such as "0042", A's rebuilt key loses the zeros and Start stops with "Fatal error", while B keeps "0042-MH" (RAN in the A audit and MAP 2.1 item 4; not re-run here).
- *History*: the key is shared (`edac6a3`). B's "Folder Name" column and stored key came with `c258047` (07-02-2026).

**G5. What must be chosen before DRY RUN** (A-R51 / B-R48 / B-R133; MAP 2.1 item 7; D37)

- *A*: both template paths must exist and a target must be typed (`A:gui/handlers/processing_handler.py:40`; `A:gui/handlers/file_handler.py:132-151`).
- *B*: only a ticked client is needed (`B:gui/pyqt_main_window.py:326-343`).
- *Worked example* (RAN (C6), probeA P5 / probeB P5): with the template and target boxes emptied and three clients ticked, A shows "Error" - "Please select both ITC and Sales template files" and no preview; B prints the preview with "📊 Would create: ITC Report" and "💰 Would create: Sales Report" for every client.
- *History*: the gate is in the shared code (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`); B's PyQt5 Dry Run without it came with `c258047` (07-02-2026).

**G6. Template paths read once or per client** (A-R60 / B-R50, B-R60; MAP 2.1 item 11; D46)

- *A*: re-reads both template paths from Step 1 for every client; each client's log block shows "Template: <file name>" (`A:gui/handlers/processing_handler.py:301, 310, 332`; RAN (C6) A10 log).
- *B*: takes the paths once, when Start is pressed (`B:gui/pyqt_main_window.py:350-371`); its log shows only "Output: <report name>".
- *Worked example* (READ; MAP 2.1 item 11): three clients; after client 1 the user picks a different ITC template on Step 1. A builds clients 2-3 from the new template; B builds all three from the template chosen at Start.
- *History*: per-client reading is the shared Tk code (`edac6a3` … `47a1492`); B's snapshot came with `c258047` (07-02-2026).

**G7. "Would create: ITC Report / Sales Report" for every client** (A-R60 / B-R50; same in both)

- Both previews promise both reports for every ticked client, whatever files it has (RAN (C6) A09, B09: ClientB, which has no Sales Reco or IMS Reco file, still gets both lines). Shared since `edac6a3`.

**G8. The Step 4 success check shown in the log ⚠** (A-R73 / B-R70, B-R71; same code)

- Both log "Checking cell BB2 in Info sheet..." and then "✅ Cell BB2 populated successfully: <value>" or "❌ Cell BB2 is empty (None)" / "is blank" / "contains 0" / "contains error" (READ `A:power_query_extractor/core/report_processor.py:655-680`; `B:power_query_extractor/core/report_processor.py:606-631`). Neither could be RAN here (no Excel). The check function is identical in both (MAP §1 A-F21: `_validate_refresh` IDENTICAL). Shared since `edac6a3` (31-08-2025, "Initial commit"; RAN (VERIFY-C6): `git log -S "Checking cell BB2"`, a search of each project's history for the commits that added or removed that text, finds only that commit in both).

**G9. Report names with the state code in the Step 3 log ⚠** (A-R59 / B-R51; same in both)

- RAN (C6): "ITC_Report_ClientA_MH_231026_1731.xlsx" (A) and "ITC_Report_ClientA_MH_231026_1732.xlsx" (B): the same form, the minute differs only because the runs were a minute apart. Shared since `edac6a3`.

### 6.8 Every screen finding of the two audits, and where it stands in the other project

#### 6.8.1 Project A's findings (A audit §10.16, U1-U17)

| A finding | In B | 6.5 row |
|---|---|---|
| U1 "Full pipeline" does not refresh by itself; the texts promise it does | B has no workflow choice and no such promise; it moves to Step 4 after every run and does not refresh by itself either | D14 |
| U2 Step 2/3 buttons react to the mouse press, ignore the keyboard, and a greyed STOP still reacts | B's buttons are normal Qt buttons; its disabled Stop ignores clicks but looks active (B U6) | D36, D41 |
| U3 The window freezes during a scan (Step 1 and Step 4) | B scans in a worker thread; but its Step 4 list is built on the window's own thread at start-up, after processing, after each extraction that ends normally or is stopped (not after one that ends in an error) [Corrected during review - originally "after each extraction"; Section 4, 4.7 item 2] and on "Refresh List" (`B:gui/pyqt_main_window.py:190, 384, 561`, `B:gui/views/extraction_view.py:193`; B audit §10.7, B-G58) | D23 |
| U4 Step 2 "📁 Name" column off-screen and unexplained | B has no per-client name toggle; its "Folder Name" and "Last Run" columns need a 44-px horizontal scroll | D29, D33 |
| U5 Dark-mode defects | B's dark mode has other defects (B U10) | D63 |
| U6 At the minimum size the status bar and the Dark Mode tick disappear | B's minimum (1200×800) is taller than a 768-px laptop screen (B U24) | D4 |
| U7 "Processing Complete" always says "Successfully", also after Stop | Same in B (B U15) | D43 |
| U8 Step 4 gives no visible feedback on its "Setup" sub-tab when a scan finds nothing (when no `Annual Statement-…` folder is found, the message goes only to the "Processing" sub-tab's log, `A:power_query_extractor/gui/extractor_window.py:625-628`) | B shows the reason in its light-blue strip ("No 'Annual Statement-...' folders found in target.") | D49 |
| U9 No Stop in Step 4; closing during Step 4 gives no warning; Excel is driven by keystrokes | B has a Stop (acting between versions); closing gives no warning either; the same keystrokes (B U30) | D55, D10, D60 |
| U10 Technical error texts in boxes | Same in B ("Processing Error", "Extraction Error", "Scan Error" show raw error text) | D64 |
| U11 Help texts that disagree with the code | Same kind of problem in B (B U7) | D66 |
| U12 The Step 4 fallback card blames pywin32 for any failure | B has no fallback: no window opens without pywin32 (B U1) | D1 |
| U13 Standalone extractor: the selected sub-tab is drawn smaller | B has no standalone window | D2 |
| U14 Template picker `.xltx/.xlsx` against four endings in the checklist | B's picker offers `.xlsx/.xlsm`; no checklist | D16 |
| U15 Settings saved only on some actions | B saves at different moments, including on closing (B U20 for the parts never read back) | D9 |
| U16 Needs a console; start-up problems only as console text | B does not need a console: its first line checks for one (`B:main.py:12-13`; RAN (VERIFY-C6): a no-console import of each `main.py`, which runs the first lines but not the window, passes B's and stops A's with "'NoneType' object has no attribute 'buffer'"; the B audit's history section records the same fix in `c258047`). But B's start-up problems also appear only as console text and in `gst_organizer.log` (B U1) | D1 |
| U17 Lists not sorted by client name | B's Step 2 is sorted by client key and can be re-sorted by header; its Step 4 is sorted by folder name | D32, D50 |

#### 6.8.2 Project B's findings (B audit §10.14, U1-U30)

| B finding | In A | 6.5 row |
|---|---|---|
| U1 No window without pywin32; the advice does not name it | A starts without pywin32 off Windows and, on Windows, names pywin32 and offers to install it | D1 |
| U2 Three Browse buttons invisible | All four visible in A | D15 |
| U3 For many Step 4 failures no reason on screen | A writes "⚠️ ITC Report - <reason>" per failed report (RAN (C6)) | D56 |
| U4 "Refreshed (date)" in green and a green dot after every refresh failed | A has no dots; its status shows a time whenever a refreshed copy exists, which can also follow a failed refresh | D58, D7 |
| U5 Nothing locked during jobs; closing during a job is not questioned | A also locks nothing but Start; it asks on close during Step 3 only | D47, D10 |
| U6 Disabled Stop looks enabled; Stop acts only after the current item | A's STOP looks inactive (but still reacts to a click); it also acts only between clients; Step 4 has no Stop | D41, D55 |
| U7 Help texts contradict the code and each other | A has its own list (A U11) | D66 |
| U8 "EXPORT LIST" does nothing | Works in A | D35 |
| U9 "Max length" greyed but applied | Always active in A (typed values unchecked) | D20 |
| U10 Dark mode partial; Step 3/4 status lines unreadable | A darkens all screens, with other defects (A U5) | D63 |
| U11 Step 4 shows one run folder chosen by name; uses the saved target, not the box | A takes the newest by date modified from its own folder box | D49, D50 |
| U12 Untickable Step 4 squares still drawn; all ticks lost at every rebuild | A greys unavailable ticks and keeps the ticks after a run | D51, D57 |
| U13 SCAN FILES below the visible part; Step 2 overflows; the summary shows only half | A's Scan is also below the visible part (RAN (C6)); its fixed 200-px summary box also shows only part of the report (15 of 25 lines here, RAN (VERIFY-C6); 13 in the A audit; font-dependent, 6.4.4) | D22, D28 |
| U14 On-screen logs lose indentation and "<…>" text and slow down as they grow | A's Tk logs keep indentation and all text; their speed with very long logs was not measured | D42 |
| U15 "Processing Complete" also after Stop and failed reports; "Processing complete" status overwritten | The same box in A; A's status bar never shows processing results | D43, D8 |
| U16 No extraction completion box; "x/y clients processed" counts attempts | A asks "Open the consolidated report?"; it gives no failure count either ("✅ Processing complete!") | D57 |
| U17 A Processing Error leaves the Step 3 log empty | A writes "💥 Fatal error: <error>" in the log | D45 |
| U18 "No Level 1 folder available." cannot appear | No counterpart in A | - |
| U19 Missing / duplicate types and problem files never named on screen | A names the missing types; problem files are counts only in both | D29, D28 |
| U20 Opens on Step 1 at 1400×900 although the last tab and size are saved | A saves neither; opens on Step 1 at 1200×800 | D4 |
| U21 No tooltips, no shortcuts beyond menus | A has Step 1 tooltips and Step 2 table keys; no shortcuts | D25, D65 |
| U22 Every extracted figure printed in the Step 4 log | Same in A | D61 |
| U23 An uncaught error makes the window vanish | A's window stays open; the error goes to the console | D11 |
| U24 Minimum window taller than 768 px | A's minimum is 700 px | D4 |
| U25 Dots mean "ran"; Step 2 never green; never reset | A has no dots | D7 |
| U26 After an "Extraction Error" the status bar keeps its old text | A's Step 4 messages never reach the main status bar | D8 |
| U27 A damaged `session_state.json` aborts several actions | A has no session file. A damaged A settings file, RAN (VERIFY-C6, `damagedA.py`, made-up content): a file that is not valid JSON (the text format the settings are stored in) is skipped with "Error loading cache" in the log, A opens with empty boxes and overwrites the file at the next save; a file that holds a list instead of named settings (`[1, 2]`) stops A before its window opens (the window code fails with "'list' object has no attribute 'get'", `A:gui/main_window.py:375-376`; started through `main.py` this ends as "💥 Critical error", `A:main.py:130-134`, READ) | - |
| U28 `launch_extractor.py` writes no program log | A's standalone writes `pq_extractor.log` | D2 |
| U29 The error-pop-up closer may click other programs' buttons | Same code in A | D60 |
| U30 Keystrokes go to whatever window has focus | Same in A | D60 |

#### 6.8.3 Index: every row of the feature map (MAP §1) and where this section shows it on screen

Added during the check of this section, so that each feature pair of the map can be found in the click-by-click tables (6.2, 6.3) and the difference rows (6.5). The feature verdicts themselves are in MAP §1 and Section 3; this table only points.

| A feature | B feature | Where it shows on screen in this section |
|---|---|---|
| A-F1 start-up checks | B-F1 | 6.2 step 1, 6.3 step 1, 6.4.1, D1, D11 |
| A-F2 window frame, closing | B-F1, B-F29 | 6.4.2, D3-D8, D10 |
| A-F3 the four paths and pickers | B-F3 | 6.2 step 5, 6.3 steps 3-5, D15, D16, D9 |
| A-F4 live checklist | — (only A) | 6.2 step 3, D13, D17 |
| A-F5 workflow choice | — (only A) | 6.2 step 4 and variants, D14, D44 |
| A-F6 length box | B-F4 | 6.2 step 6, 6.3 step 6, D20 |
| A-F7 scanning | B-F6 | 6.2 step 7, 6.3 step 7, D23, D28 |
| A-F8 Step 2 table and selection | B-F8, B-F7 | 6.2 steps 8-9, 6.3 steps 8-9, D29-D32, D34 |
| A-F9 problem files ("variations") | B-F9 | D28 |
| A-F10 include-name tick and per-client toggle | B-F4, B-F5 | D19, D33 |
| A-F11 Dry Run | B-F11 | 6.2 step 10, 6.3 step 10, D37, D38, G5, G7 |
| A-F12 EXPORT LIST | B-F10 | 6.2 step 11, 6.3 step 11, D35 |
| A-F13 folder tree and copying | B-F12 | 6.2 step 13, 6.3 step 13 (what the log shows), D42 |
| A-F14 ITC report | B-F13 | D42, G9 |
| A-F15 Sales report | B-F14 | D42 ("[DEBUG]" lines), D46, G6 |
| A-F16 summary workbook | B-F15 | 6.2 step 13, 6.3 step 13, D42 |
| A-F17 Step 3 progress, Stop, completion box | B-F16 | 6.2 steps 12-14, 6.3 steps 12-14, D39-D45, D47 |
| A-F18 Step 4 page | B-F18, B-F24 | 6.2 step 15, 6.3 step 15, 6.4.6, D48, D49 |
| A-F19 which batch Step 4 shows | B-F18 | 6.2 step 16, D50 |
| A-F20 Step 4 list and refresh status | B-F18 | D51, D58 |
| A-F21 the refresh | B-F19, B-F20 | 6.2 step 18, 6.3 step 16, D53, D60, G8 |
| A-F22 refreshed-copy suffix | B-F21 | D53, D58 |
| A-F23 reading the cells | B-F22 | D61 |
| A-F24 consolidated workbook | B-F23 | 6.2 step 19, 6.3 step 17, D57 |
| A-F25 Step 4 log, Stop, record | B-F24, B-F25 | D55, D56, D59 |
| A-F26 closing error pop-ups | B-F26 | D60 |
| A-F27 settings file | B-F27, B-F2 | D9, 6.3 variant "First start after using A" |
| A-F28 dark mode | B-F28 | 6.4.8, D62, D63 |
| A-F29 standalone extractor | B-F30, B-F31 | 6.2 and 6.3 variants, D2 |
| A-F30 log files | B-F32 | 6.4.1, 6.5.7 "Log file" |
| A-F31 help texts, tooltips | B-F33 | D24, D25, D27, D66, G1 |
| — | B-F2 settings import (only B) | 6.3 variant "First start after using A" |
| — | B-F17 run history, "Last Run" column (only B) | D29, D59 |
| A-F7, A-F17, A-F21 threads | B-F34 | D23, D47, 6.8.1 U3 |

### 6.9 Compared with the author's Standards 12 and 13 (facts only)

| Standard says | A | B |
|---|---|---|
| Desktop windows with PySide6 (Qt 6) - the official Python edition of the newer Qt toolkit (`rebuild_pipeline_FINAL.md:833`) | Tkinter | PyQt5 (Qt 5) |
| Users start the app with `Run [App Name].bat` and `pythonw.exe` (Python without a console window), with no console window (`:951-969`) | `python main.py` from a console; the program needs one (RAN (VERIFY-C6): with no console A stops at its first line) | `python main.py` from a console; its first lines also pass with no console (RAN (VERIFY-C6), first lines only; the window itself was not started that way here), but no launcher exists |
| Nothing is installed on users' PCs (`:924`) | Offers to pip-install packages into the user's Python | Tells the user to pip-install packages |
| `sys.excepthook` / `threading.excepthook` (Python's last-resort handlers for errors nothing else caught) write the error to a file under `%LOCALAPPDATA%\[App Name]\` (the user's own local settings folder on Windows) and show a plain-English message box (`:997-1025`) | None; errors after start-up go to the console only | None; an uncaught error closes the window |
| `-B` so that users' PCs write no files into the shared app folder (`:969-970`) | Creates `logs` and `temp` folders in the program folder at every start and writes `gst_organizer.log` into the start folder | Writes `gst_organizer.log` into the start folder |
| Settings and logs under `%LOCALAPPDATA%\[App Name]\` (`:1035-1040`) | Settings `~/.gst_organizer_cache.json`; log in the start folder | Settings `~/.gst_organizer/`; log in the start folder |
| Excel through COM: a separate, hidden Excel; never take over the user's Excel; refresh with `RefreshAll()` (Excel's own "refresh every query" command, called directly instead of through key presses) and wait for the queries (`:844-878`) | Step 4 shows Excel visibly and types Alt+A, R, A (same code as B) | Same as A |

Which of these the merged program follows is the author's decision.

### 6.10 Corrections and clarifications found while writing this section

1. **MAP §1 row A-F2** says "Both switch tabs on their own after a scan and when processing starts." For a scan this is true in A only when at least one client is recognised; B switches to Step 2 after every completed scan, including an empty one. RAN (C6): probeA P3 (empty folder → stays on Step 1, "💡 No valid GST files found"), probeB P3 (→ Step 2, "0 clients (0 complete)"). Code: `A:gui/handlers/file_handler.py:83-91`; `B:gui/pyqt_main_window.py:304-307`.
2. **MAP §1 row A-F31** says A has "a file-name card built from the real patterns". The card prints the separate `description` text stored next to each recognition pattern (`A:gui/tabs/setup_tab.py:494-503`; `A:utils/constants.py:19-59`), not something generated from the patterns themselves. So it leaves out what the patterns also accept (`.xls`, any letter case, "(N)" forms), as the A audit §10.5.10 says. MAP 2.1 item 3 words it correctly ("the pattern's own example text"). The same `description` texts also exist in B's code (`B:utils/constants.py:19-59`), where nothing reads them; B showed the same texts (typed by hand into a help panel, not read from the patterns) in `c258047` and replaced them with the short list in `7e09e5c` (6.7 G1).
3. **Clarification, not an error (MAP §1 A-F20; A audit §10.8.3).** After a refresh that fails inside Excel, A's "Last Refresh Status" also shows a time, because A, like B, copies the report before refreshing it (`A:power_query_extractor/core/report_processor.py:164`; `B:power_query_extractor/core/report_processor.py:120`). The difference the user sees is the wording: A shows "ITC: <time>", B shows "Refreshed (<time>)" in green. Without pywin32, A fails before copying, so on this machine its rows stayed "Never refreshed" (RAN (C6), A15), while B's stand-in run left a "_Refreshed_" copy beside every report (RAN (C6)).
4. **Clarification of the B audit §10.0 Limits.** The B audit says emoji are drawn as small boxes in its Linux screenshots. In this section's B screenshots the log emoji appear as small pictures, while A's Tk screenshots show nothing for them. Either way the Windows look was not seen; it does not change any finding.

Added during the check of this section (RAN (VERIFY-C6) and READ in both projects' code):

5. **MAP §1 row A-F15** says "A shows 14 '[DEBUG]' lines in the Sales part of the Step 3 log". The code holds 14 different "[DEBUG]" lines (`A:gui/handlers/processing_handler.py:333, 342, 353, 362, 363, 367, 371, 374, 391, 393, 398, 403, 404, 411`), but they sit on different paths: a client whose Sales report is made gets 8 of them, and no client can get more than 9. RAN (VERIFY-C6): 24 "[DEBUG]" lines for the three placeholder clients, 8 each.
6. **A audit §10.6** says "Only the first 13 lines fit in the 200-px box" of the Step 2 summary. The box is a fixed 200 px high (`A:gui/tabs/validation_tab.py:83-85`), so the number of lines depends on the font: RAN (VERIFY-C6, `measureA.py`) counted 15 of 25 lines with this machine's look-alike font. Either way only part of the report shows; the Windows count is not sure (how to check: count the visible lines after a scan on an office PC).
7. **MAP 2.1 item 5** gives B's Step 2 colours as coming from "`c258047` / `7e09e5c`". The three text colours and the "equals Complete" test of "Complete Only" are already in `c258047` (07-02-2026; `c258047:gui/views/validation_view.py:141-146, 199`); `7e09e5c` (08-02-2026) rebuilt the screen and kept them unchanged (READ, `git show` of both commits).
8. **Clarification of the B audit §10.5** ("B has no keyboard-selection code"). True of B's own code, but Qt's standard table keys still work: with the cursor in the "Select" column, ↓ moves to the next row and Space ticks or unticks it; Space in any other column does nothing, and nothing on screen says so. RAN (VERIFY-C6), `keysB.py` (B's real Step 2 view, two placeholder rows, real key presses sent through Qt): both rows ticked → Space → first row unticked → ↓, Space → second row unticked → cursor moved to the "Client" column, Space → no change. D32 and D65 are written accordingly.

No other statement of the two audits' Section 10 or of the map was found wrong in what this section re-checked: titles, sizes, tab names, the Step 1 scroll facts of both, the status-bar texts, the dots, the invisible Browse buttons, the dead EXPORT LIST, the message texts in 6.6, the end-of-run boxes, and the Step 4 statuses and logs.

### 6.11 Questions for the author (screens)

1. How do you and your staff start the tool today: from a console with `python main.py`, through `launch_extractor.py`, or another way? (Standard 13 expects a no-console launcher.)
2. Do you want Step 1 to check each path as it is typed (A's checklist), or only when Scan is pressed (B)?
3. After Step 3, should the tool move to Step 4 by itself (B, or A's "full pipeline"), stay on Step 3 (A's "organise only"), or let you choose? Should the Power Query refresh ever start without a click?
4. After a scan, should the clients start ticked (B) or unticked (A)?
5. On Step 2, do you want the missing return files named (A: "IMS Reco, Sales Reco") or counted (B: "2")? **[GST RULE]** - your call.
6. Which wording of the expected file names should Step 1 show (A's "…-ClientName-State-Period.xlsx" with type names, or B's shorter list), and should it mention `.xls` and the "(N)" duplicate forms? **[GST RULE]** ⚠
7. Should Dry Run require the templates and target first (A) or not (B)? **[GST RULE]**
8. Do you use "EXPORT LIST", the per-client "📁 Name" toggle, or the Step 2 keyboard keys?
9. Step 4: one page with a client → version tree (B), or a folder box with a Scan button and two sub-tabs (A)? Should Step 4 let you pick which `Annual Statement-…` batch to refresh?
10. At the end of Step 4, do you want the question "Open the consolidated report?" (A), a summary of how many reports succeeded and failed, and a text log of the run in the batch folder (B)?
11. Should closing the window during any run ask first?
12. Do you use dark mode?
13. What screen sizes and Windows display scaling do the office PCs and laptops use? (B's minimum window is taller than a 1366×768 laptop screen; at A's minimum two controls disappear.)
14. During a Step 4 refresh, have you seen letters appear in the wrong window, or another program's pop-up closed by itself? (Same code in both.)

## 7. Tech stack against my standards

This section puts the building blocks of the two old versions (the language, the window toolkit, the Excel libraries, how Excel is driven, how Power Query is protected, how the app starts, where it keeps its files, its tests and its package list) next to your own standards. Those standards are Standard 12 ("Choosing the Language and Libraries") and Standard 13 ("How Users Start the App") in `/home/user/Annual_Statement_2.0/.build-docs/toolkit/rebuild_pipeline_FINAL.md`, plus the "Tech Stack" block of the CLAUDE.md copy in the same file. Section 8 then looks at code health: tests, error handling, dead code, silent bugs and file safety.

Nothing in Sections 7 and 8 decides a GST rule. Where a point touches one (for example a test that pins a file-name pattern), it is tagged **[GST RULE]** and only described. "Closer to the standard" and "healthier" are technical judgements only: crashes, the safety of Excel and of files, tests, speed and clear messages.

### How to read Sections 7 and 8

**The two projects.** A = `goyal-sidhant/AnnualStatement_oldv1`, branch `main`, snapshot `bf2f1dd` (27-07-2026). It is a Tkinter window calling itself "GST File Organizer v3.0", with 288 tests. B = `goyal-sidhant/AnnualStatement_oldv2`, branch `main`, snapshot `b0e7d0e` (09-02-2026). It is a PyQt5 window calling itself "v4.0", with no tests. The names mislead: A is not the older one, and B is not newer or better. Both share their first 12 commits (`edac6a3`, 31-08-2025 … `47a1492`, 30-10-2025) and split after `47a1492`. B then made 5 commits of its own (07-02-2026 to 09-02-2026); A made 34 (2 in March 2026, 32 in July 2026). GitHub's records say B's repository was created on 31-08-2025 and A's on 07-02-2026.

**Evidence marks.**

| Mark | Meaning |
|---|---|
| `A:path:line`, `B:path:line` | A place in A at `bf2f1dd` or in B at `b0e7d0e`; `path:line@commit` is older code. |
| `Std:NNN` | Line NNN of `/home/user/Annual_Statement_2.0/.build-docs/toolkit/rebuild_pipeline_FINAL.md` (Standard 12 = lines 817-898, its "Excel through COM" part = 844-878; Standard 13 = 900-1061; the CLAUDE.md copy's "Tech Stack" = 205-218). |
| READ | Established by reading the code of both projects. |
| RAN (C78) | Run for these two sections in `/root/merge-scratch/run/C78/`, on copies of the code, with `python -B` and made-up data only. The list is in 8.8. |
| RAN (A audit), RAN (B audit) | Run by that project's audit; the run was re-read here and the code checked again. |
| HIGH / MEDIUM / LOW | How sure an inference is that was not run. |
| ⚠ | A value that changes over time: Python, package and Qt versions, Excel's key tips, template cell positions. |

**COM cannot run here.** COM (Windows' way for one program to remote-control another; here Python drives Microsoft Excel) needs Windows and Excel. This computer is Linux without either. So everything Excel itself would do is READ only. The audits drove the Python code with fake COM objects (stand-ins that only record which calls are made). That proves the order of the program's own calls, never Excel's reaction.

**Words used in these two sections.**

| Word | Meaning |
|---|---|
| Python | The programming language both projects are written in. |
| Window toolkit | The library that draws windows, buttons and lists. *Tkinter* is Python's built-in one (A). *Qt* is a large separate one; *PyQt5* (B) and *PySide6* (your standard) are two Python wrappers for it, for Qt 5 and Qt 6 respectively. |
| openpyxl | A Python library that reads and writes `.xlsx` files without Excel. When it saves an existing workbook it drops Power Query. |
| pywin32 | The Python package that provides COM (`pythoncom`, `win32com`) and Windows window functions (`win32gui`, `win32api`). |
| Power Query, DataMashup | Excel's data-import tool. Its queries live inside the workbook in a hidden part (`customXml/item*.xml`) that contains an element called `DataMashup`. "Refresh" re-runs the queries. |
| Console | The black text window. `python.exe` opens one; `pythonw.exe` (the Python your launcher uses) does not, so `print()` shows nothing. |
| Launcher | Standard 13's `Run [App Name].bat`, which starts `pythonw.exe -B -s main.py` from the shared runtime. |
| Error hook (`sys.excepthook`, `threading.excepthook`) | Python's last-chance handler for an error nothing else caught. Standard 13 uses it to write an error file and show a message box. |
| `requirements.txt`, pinned | The list of packages `pip` (Python's package installer) installs. *Pinned* = an exact version (`==3.1.5`); a *minimum* (`>=3.1.2`) takes whatever is newest that day. |
| Test, pytest | A small program that runs part of the real code on known inputs and checks the result. *pytest* is the tool that finds and runs them. |
| `%LOCALAPPDATA%`, `%USERPROFILE%` | Windows' per-user folders: `C:\Users\<name>\AppData\Local` and `C:\Users\<name>`. |
| Staging file | A temporary local copy that is worked on and then copied to its final place. |
| `GetActiveObject`, `Dispatch`, `DispatchEx` | Three ways for Python to get an Excel through COM. `GetActiveObject` takes over an Excel that is already running; `Dispatch` asks Windows for "an Excel", which may also be one already running; `DispatchEx` always starts a separate, new Excel (what Standard 12 asks for). |
| Worker thread | A background job that runs beside the window, so the window stays responsive while files are processed. |
| Free-threaded Python (`python3.14t`) | A special build of Python 3.14 that can run threads truly in parallel. Packages must be installed for it separately, so it can lack packages the normal `python.exe` has. |
| Wheel | A ready-built package file that `pip` downloads; its "metadata" states, among other things, which Python versions it needs and its licence. |
| Tcl/Tk | The drawing engine underneath Tkinter. |
| Exit code | The number a program hands back to Windows when it ends; 0 means "ended normally", 134 means "aborted". |
| cp1252 | The older Western-European character set that many Windows PCs still use by default; it has no ✓ or ❌. |
| SHA-256 fingerprint | A short code computed from a file's bytes; it changes if even one byte changes, so equal fingerprints mean an unchanged file. |
| Regression | Something that used to work and breaks after a later change. |
| Re-raise | An error handler that records an error and then passes it on to the code that called it. |
| Docstring | The description written at the top of a Python file or function. |
| Re-export line | A line in a package's `__init__.py` that makes a function available under the package's name; it is not a real use of the function. |

### 7.1 The table: one row per topic

| Topic | A | B | What my standard says | Where each project falls short |
|---|---|---|---|---|
| **Python version: stated and actually needed** ⚠ | States "Python 3.7 or higher" (`A:requirements.txt:2`; `A:README.md:51` adds "3.8+ recommended"). The code's grammar fits 3.7 (RAN (A audit) 9.4). The packages that ran here need more: openpyxl 3.1.5 needs 3.8 or newer, pytest 9.1.1 needs 3.10 or newer (RAN (A audit), package metadata). On your PC, `py` starts the free-threaded `python3.14t` without pywin32 and `python` starts `python.exe` with pywin32; the commit messages (`c463256`, `f363946`, `be6418c`, 26-07-2026) do not give `python.exe`'s version (MEDIUM that it is 3.14). A needs Tkinter. | States the same two lines (`B:requirements.txt:2`; `B:README.md:51`). The grammar fits 3.7, and `dataclasses` (`B:models/app_state.py:6`) also needs 3.7 (RAN (B audit) 9.4). All 38 files compile on 3.11, 3.12, 3.13 and 3.14 (RAN (B audit)). Windows builds of every package B uses exist for Python 3.14; the newest `PyQt5-sip` build needs 3.10 or newer (RAN (B audit), wheel metadata, 23-10-2026). The 2025 developer used 3.12 and 3.13 (compiled-file headers). B does not need Tkinter. | Language: "Python 3" (`Std:206`). "All testing used Python 3.14. Move `_runtime` to a newer Python only after checking that every package listed has a Windows wheel for it - ask me first" (`Std:931-932`). | **Both** promise a 3.7 floor that nobody tested, and neither names 3.14. **Both** clean client names differently on 3.14 than on 3.11-3.13: a name ending in a dot keeps the dot on 3.14 (A-G40, B-G80; RAN in both audits). One shared runtime on one Python version removes that difference. **A only:** it needs Tkinter, and python.org's "embeddable package", which `_runtime` is built from (`Std:913-914`), does not include Tkinter (MEDIUM, general knowledge, not checked here; how to check: look for `_tkinter.pyd` and a `tcl` folder in the embeddable zip). **Closer:** B (MEDIUM; only because of Tkinter). |
| **Window toolkit** ⚠ | Tkinter with ttk (Tk's themed widgets). 17 of A's program files import it (RAN (C78) grep; the A audit counts 25 import lines in those 17 files). Tcl/Tk 9.0 here; the office version is NOT SURE (check with `python -c "import tkinter; print(tkinter.TkVersion)"`). Present since the first commit `edac6a3`. | PyQt5, `PyQt5>=5.15.0` (`B:requirements.txt:5`). 14 program files import it (RAN (C78) grep). This is the Qt 5 family: PyQt5 5.15.11 here; the newest Windows build of Qt itself (`PyQt5-Qt5`) is 5.15.2, a 50 MB download. The wheel metadata names the licences: PyQt5 (the Python wrapper) "GPL v3", PyQt5-Qt5 (the Qt library) "LGPL v3", PyQt5-sip "BSD-2-Clause" (RAN (B audit) 9.1 and 8.2.2; re-read here in the `License:` fields of the installed packages, RAN (VERIFY-C78)). Came in with `c258047` (07-02-2026, "PyQt5 migration"); `a62d3cc` the same day deleted the Tkinter files. | "PySide6 (the `PySide6-Essentials` package, Qt 6)"; "Pin it to 6.12.x while any PC runs Windows 10 (Qt 6.12 is the last Qt for Windows 10)" (`Std:833`). The CLAUDE.md Tech Stack: "PySide6 (the `PySide6-Essentials` package, Qt 6) [— or PyQt5, only for an existing PyQt5 app that is being kept as it is]" (`Std:208-209`). Charts need PySide6-Addons, so ask first or prefer an Excel report (`Std:840`). | **Neither** uses PySide6. **A:** every window is Tk and would have to be redrawn in Qt from scratch. **B:** its windows are already Qt widgets, with Qt signals (messages between parts of the program) and Qt worker threads. PySide6 offers the same building blocks, because both wrap Qt. But B is Qt 5, not Qt 6, so moving it is still a port with some changed calls (MEDIUM, general knowledge; the size of that work was not measured). Whether the PyQt5 exception in `Std:209` covers "Annual Statement 2.0" is your decision. B's licence line is stated as a fact only. **Closer:** B, in structure; neither is compliant. |
| **Excel libraries** ⚠ (package versions) | openpyxl `>=3.1.2` (`A:requirements.txt:5`) writes the two workbooks A creates itself: the summary (`A:core/excel_handler.py:376-405`) and the consolidated Step 4 workbook (`A:power_query_extractor/core/data_consolidator.py:20-186`). It also writes the fallback report (`A:core/excel_handler.py:262-322`). pywin32 `>=305; sys_platform == 'win32'` (`A:requirements.txt:9`, added by `064c634`, 23-07-2026) drives Excel for Steps 3 and 4. lxml is listed but never imported; openpyxl uses it by itself. No Excel library of the program (openpyxl or pywin32) ever opens a client's input file: Python only reads its first 8 bytes and copies it (`A:utils/helpers.py:275-283`, run for every listed file from `:354-357`; `B:core/file_parser.py:196`). The input files are read only by the templates' own Power Query queries when Excel refreshes a report (READ). | The same split. openpyxl `>=3.1.2` (`B:requirements.txt:8`) writes the summary (`B:core/excel_handler.py:318-340`), the consolidated workbook (`B:power_query_extractor/core/data_consolidator.py:167`) and the fallback report (`B:core/excel_handler.py:219-262`), which the window can never reach (row 5). pywin32 is needed, but its line is commented out: `# pywin32>=306` (`B:requirements.txt:13-14`). lxml is the same as in A. | openpyxl for workbooks that have no Power Query (`Std:827`). "Writing into, or refreshing, a workbook that has Power Query queries": "Excel itself, driven through COM with `pywin32`" (`Std:828`). openpyxl "silently drops Power Query queries (tested) and shapes when it saves". "Reading a Power Query workbook with openpyxl is fine: only saving loses the queries" (`Std:827`). | **Both** keep an openpyxl path that writes into a copy of a Power Query template (against `Std:828` and `Std:846-850`; row 5 and 7.2). **Both** read the Step 4 figures through Excel; `Std:827` would also allow openpyxl for that. This is a fact, not a recommendation. **B only:** `requirements.txt` leaves out a package without which the window cannot start. **Closer:** equal on the libraries. A's package list is correct; B's is not. |
| **Excel through COM** (each rule of `Std:844-877`; details in 7.2) | Three Excel sessions: Step 3 making reports, Step 4 refresh, Step 4 reading figures. All three were written in `edac6a3` (shared). A changed none of the Excel settings after the split; it added local staging, a shared "already open?" helper and guarded imports, plus two changes that are not Excel calls: Step 4 refuses each client when pywin32 is missing, and Skip Refresh looks for the configured suffix (`2aaa2b1`) (READ: `git diff A/main B/main`). Verdicts per session (7.2): Step 3: 4 met, 8 not met, 8 do not apply. Refresh: 6 met, 11 not met, 3 do not apply. Reading: 6 met, 6 not met, 8 do not apply. | The same three sessions with the same Excel calls (READ: the differences in these two files are A's staging, its helper, the types of error caught and one unreachable `return` that B removed, plus three that are not Excel calls: A's guarded pywin32 import, A's "pywin32 missing" refusal per client, and A's configured-suffix Skip Refresh lookup). The verdicts are identical to A's. | `DispatchEx`, hidden, prompts off, never the user's own Excel (`Std:851-855`). `Quit()` in `finally` (`Std:856`). `IgnoreReadOnlyRecommended=True` and the `ReadOnly` check (`Std:857-859`). `BackgroundQuery = False`, `RefreshAll()`, `CalculateUntilAsyncQueriesDone()`, no `Refreshing` before saving, never `RefreshDate` as proof (`Std:860-867`). `CoInitialize` first, the Excel work in a separate function, `CoUninitialize` after it returns (`Std:872-875`). Excel installed, a logged-in user, never a scheduled job (`Std:876-877`). | **Both the same:** no `DispatchEx` anywhere (RAN (C78) grep: 0 in each). Step 3 takes over the user's open Excel, hides it, switches its prompts back on and quits it (`A:core/excel_handler.py:124, 131, 240, 244`; `B:core/excel_handler.py:111, 118, 205, 209`). The refresh is three key presses (Alt+A, R, A ⚠) plus fixed waits, and the workbook is saved whatever happened. `RefreshAll`, `BackgroundQuery`, `CalculateUntilAsyncQueriesDone`, `Refreshing`, `RefreshDate`: 0 hits in each (RAN (C78) grep). `IgnoreReadOnlyRecommended` appears only in the last fallback open (`A:power_query_extractor/core/report_processor.py:781`; `B:power_query_extractor/core/report_processor.py:729`). `ReadOnly` is never read. `CoUninitialize` runs inside the same function while Excel objects are still held. **Closer:** neither (identical code). |
| **How Power Query is protected** | **On purpose, with a gap.** A start-up check (`f363946`, made blocking by `be6418c`, both 26-07-2026) treats pywin32 as essential on Windows. If it is missing, A offers to install it and otherwise refuses to start (`A:utils/requirements_check.py:44-60`; `A:main.py:75-102`). The code comment says why: "Producing damaged reports is worse than refusing to start" (`A:utils/requirements_check.py:52-53`). Gap: the check only looks for a `pythoncom` file (`find_spec`, `A:utils/requirements_check.py:64-70`). A pywin32 that is present but cannot load passes the check, and Step 3 then quietly falls back to openpyxl (`A:core/excel_handler.py:19-28, 90-102`). The reports lose their Power Query and the screen says "Success" (A-G7; RAN (A audit) X1 R1 and R4, simulated). Off Windows the fallback runs with two log-file warnings only, and Step 4 refuses each client (`A:power_query_extractor/core/report_processor.py:66-77`). The templates themselves are only ever copied. There is no `DataMashup` check (RAN (C78) grep: 0). | **By accident, and complete through the window.** `B:power_query_extractor/core/report_processor.py:11-15` imports pywin32 without any guard, and the window loads that module at start-up. So without pywin32 the window cannot open at all, and the openpyxl fallback is never reached (B-G32, B-R97; RAN (B audit) M1). This began with `c258047` (07-02-2026), as a side effect of the PyQt5 rewrite. At the split point the Tkinter organiser did not load the extractor, so a Windows PC without pywin32 took the fallback and lost the queries (READ (B audit) X1.1). The fallback code is still there (`B:core/excel_handler.py:219-262`) and goes live the moment that import is made optional (HIGH, B audit). There is no `DataMashup` check (RAN (C78) grep: 0). | "Before saving with openpyxl into a workbook the app did not create, check it for Power Query (a `customXml/item*.xml` part containing `DataMashup`). If one is found, use COM, or stop and tell the user. Never let openpyxl save over such a workbook" (`Std:846-850`). | **Both** lack the `DataMashup` check that `Std:846-850` asks for. **A:** the guard is deliberate and explained, but leaky: a broken pywin32 gets through, and the result is silent. **B:** the protection is a side effect. `requirements.txt` leaves pywin32 out, and the error message advises `pip install PyQt5 openpyxl`, which does not fix it (`B:main.py:49-54`). **Closer:** neither meets the rule. Through the window, B never strips Power Query, but Steps 1-2 cannot run without pywin32 either. A tells the user what is missing and why. |
| **How the app starts: entry point, console, launcher** | `python main.py` in a console (`A:main.py:136-137`; `A:README.md:112`). `python launch_extractor.py` opens the Tkinter extractor as its own window and skips the start-up check (`A:launch_extractor.py:12-30`). A third way, `python -m power_query_extractor.extractor_main`, also works (A audit 2.7.3). There is no `.bat`, no shortcut and no `pythonw`. **A needs a console:** `A:main.py:17` wraps `sys.stdout.buffer` without a check. RAN (C78): loaded with no console streams (what `pythonw.exe` gives), it stops with `AttributeError: 'NoneType' object has no attribute 'buffer'` before any log exists. Start-up problems are 14 `print()` calls and a `_pause()` that waits for Enter (`A:main.py:45-54`). There is no error hook (RAN (C78) grep: 0). It puts its own folder first on the search path (`A:main.py:31-32`). | `python main.py` (`B:main.py:62-63`). `launch_extractor.py` opens the same window on Tab 4 and sets up no logging (`B:launch_extractor.py:11-24`). The old `power_query_extractor/extractor_main.py` cannot even be imported, because its line 9 imports a file that `a62d3cc` deleted (B-G45). There is no `.bat`, no shortcut and no `pythonw`. **B survives without a console:** `B:main.py:12-13` checks for one first (since `c258047`). RAN (C78), re-run (VERIFY-C78): with no console streams it loads and creates `gst_organizer.log` (empty at that point) in the start folder. But start-up errors are only printed (`B:main.py:49-59`), so under `pythonw` they are invisible (B-G34). There is no error hook (RAN (C78) grep: 0), and an error inside any window handler ends the whole program with exit code 134 (B-G35; RAN (B audit)). `B:main.py:26-29` puts its folder first, without `.resolve()`. | Users double-click a desktop shortcut and no console appears (`Std:902-906`). Launcher `Run [App Name].bat`: `start "" /D "%~dp0" "%PYW%" -B -s "%~dp0main.py"`, with a message box if the shared Python is missing (`Std:951-975`); shortcut set to Run: Minimized (`Std:977-980`). `main.py` starts with `APP_DIR`, `LOG_DIR`, `report_fatal_error`, `sys.excepthook`, `threading.excepthook` and `sys.path.insert(0, str(APP_DIR))` (`Std:984-1029`). "There is no console: `print()` shows nothing. Report problems in the window" (`Std:1035-1036`). Never a `.vbs` or `.ps1` launcher (`Std:974-975`). | **Neither** has the launcher, the shortcut or the error hooks. **A** dies silently under `pythonw` (A-G22). **B** starts, but every start-up failure is invisible, and a window error kills it (B-G34, B-G35). Both `launch_extractor.py` files *append* their folder to the search path instead of putting it first (`A:launch_extractor.py:10`; `B:launch_extractor.py:8`). **Closer:** B (it survives no console). Neither is compliant. |
| **Where settings, cache and logs live** | **Settings:** `%USERPROFILE%\.gst_organizer_cache.json`, one file (`A:gui/utils/cache_manager.py:15`), rewritten whole in place (`:45-46`). Step 4's wait time, suffix and Skip Refresh are not saved. **Logs:** `gst_organizer.log` in the *current folder* (`A:main.py:24`), and `pq_extractor.log` there too for the standalone extractor (`A:power_query_extractor/extractor_main.py:16`). Both hold client names, Links values and extracted figures (A-G56). **Program folder:** empty `logs\` and `temp\` folders (plus any missing `__init__.py`) are created there at every start and never used (`A:main.py:35-43`). **Temp:** `%TEMP%\gst_excel_temp` for report staging (`A:core/excel_handler.py:53-54`). **Cache:** nothing on disk besides the settings file, which A calls its "cache". The speed-ups for the network share (one folder listing per scan, remembered path checks) are in memory only (A audit 8.4.4). No run history. | **Settings:** `%USERPROFILE%\.gst_organizer\` holds `settings.json` (11 keys, including Step 4's options), `session_state.json` (written; read back only in order to rewrite it, never to restore anything, because `load_session_state` and `get_recent_source_folders` have no caller), `run_history.json` (every client's runs, including extracted figures) and the empty marker file `.migrated` (`B:models/cache_manager.py:16-37, 99`). The three JSON files are rewritten whole in place (`:135-142`; the run history by `B:models/run_history.py:128-137`). A's old `.gst_organizer_cache.json` is read once (`:93-131`), but its recent-folder list is lost because the key names differ (B-G24). **Logs:** `gst_organizer.log` in the *current folder*, only when started through `main.py` (`B:main.py:20`). Also `Extraction_Log_<YYYYMMDD_HHMM>.txt` in the run's top folder (`Annual Statement-<DDMMYY HHMM>`, which holds that run's client folders) (`B:gui/controllers/extraction_controller.py:109-121`). **Temp:** `%TEMP%\gst_excel_temp` is created at each run and never used. **Cache:** no folder-listing cache (B audit 8.4.4). | "Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`), never in the shared app folder" (`Std:1039-1040`); `LOG_DIR = %LOCALAPPDATA%\<App>\logs` (`Std:997`). "Find the app's own files from `APP_DIR`, never from the current folder" (`Std:1034`). `-B` "stops users' PCs writing files into the shared folder" (`Std:969-970`). "Never keep … any file several people write at the same time, in the shared folder" (`Std:1046-1047`). | **Both** keep settings per user, as intended, but in the profile root instead of `%LOCALAPPDATA%\<App>\`. **Both** put the log in the current folder. With the launcher, that is the shared app folder (`Std:962`), so every user would append to one log file there, which holds client names and figures (A-G56, B-G64). **A only:** it creates `logs\` and `temp\` in the program folder at every start; on a read-only share where they do not exist yet, the start stops (A-G23). **B only:** `run_history.json` keeps every extracted figure outside the client folders and is never trimmed (B-G59, B-G64), and the extraction log sits inside client data. **Closer:** B, narrowly (it never writes into the program folder). Neither is compliant. |
| **Tests** | 18 test files, 183 test functions, 288 tests (RAN (C78) `--collect-only`: 288), all added after the split (first `6c1e1c6`, 23-07-2026). `pytest.ini` has `[pytest]` and `testpaths = tests`, but no `pythonpath = .`; `tests/conftest.py:9` adds the folder instead (`A:pytest.ini:1-6`). `requirements.txt` lists `pytest>=7.0.0` (`A:requirements.txt:15`). Details in 8.1. | None: no `tests` folder, no test file, no `pytest.ini`, now or in any of B's 17 commits (RAN (C78) `git ls-tree` and `git log` grep: 0). pytest appeared only as a commented line (`requirements.txt:11@edac6a3`), which `c258047` removed. | Every test file lives in `tests/` (`Std:293-294`). Tests are written first, by a second model (`Std:299`, `Std:344`). Tests run with `..\_runtime\python\python.exe -s -m pytest`; `pytest.ini` holds `[pytest]`, `pythonpath = .` and `testpaths = tests` (`Std:944-947`). pytest is not in the app's `requirements.txt` (`Std:937-938`). | **A:** no `pythonpath = .`; pytest is in the users' package list; three tests start a second Python without `-B` and write compiled files into the project folders (8.1). **B:** no tests at all. **Closer:** A, clearly. |
| **Dependency handling** ⚠ | `requirements.txt` gives minimum versions only (`A:requirements.txt:5-15`): openpyxl, pywin32 (Windows only), lxml (never imported) and pytest. A second list lives in the code and drives the start-up check (`A:utils/requirements_check.py:37-61`). If something is missing, A offers to run `<this Python> -m pip install -r requirements.txt` (600-second limit), then restarts itself (`A:utils/dependency_installer.py:23-74, 88-158`; `a5ad487`, 26-07-2026). This installs the whole file, pytest included (A-G53, A-G59). Off Windows the offer appears at every start and can never succeed. On a Python without pip (normally true of the embeddable package; MEDIUM) it fails with "No module named pip" (RAN (A audit) 9.11). | `requirements.txt` gives minimum versions only: PyQt5, openpyxl and lxml (never imported); pywin32 is commented out (`B:requirements.txt:5-14`). The Qt runtime and its glue package (`PyQt5-Qt5`, `PyQt5-sip`) come in automatically, unpinned. There is no start-up check and no installer. A missing module gives a console message recommending `pip install PyQt5 openpyxl` (`B:main.py:49-54`). | "Nothing is installed on users' PCs" (`Std:924`). Packages go into `_runtime` with `pip install --target … <package>==<version>`, then each is imported with `python.exe -s` before it counts as ready (`Std:933-936`). "Each app keeps its own `requirements.txt` with the exact versions of the packages it imports (not pytest)… Never upgrade or remove a package already there without asking me" (`Std:937-940`; also `Std:919`). | **Both** use minimums, not exact versions, so two PCs set up months apart get different versions. **A:** lists everything it needs, but adds pytest and lxml, and its install offer changes whatever Python runs it, against `Std:924` and `Std:937-940` when used with a shared runtime. **B:** never runs pip (in line with `Std:924`), but leaves out a package it cannot start without and gives the wrong advice. **Closer:** mixed. A on completeness and clear messages; B on "nothing is installed on users' PCs". |

### 7.2 Excel through COM, rule by rule (`Std:844-877`)

**The sessions.** Each project has the same four places where a workbook is made or driven:

- **S1**: Step 3 makes each ITC and Sales report through Excel (`A:core/excel_handler.py:104-260`; `B:core/excel_handler.py:95-217`).
- **S2**: Step 4 refreshes a copy of each report (`A:power_query_extractor/core/report_processor.py:209-415`; `B:power_query_extractor/core/report_processor.py:165-366`).
- **S3**: Step 4 reads the figures from that copy (`A:…/report_processor.py:527-636`; `B:…/report_processor.py:478-587`).
- **F**: the openpyxl fallback for Step 3, which is not COM but is what the first three rules are about (`A:core/excel_handler.py:262-322`; `B:core/excel_handler.py:219-262`).

**Why one verdict column per session is enough.** The Excel calls are the same in A and B (READ: `git -C /root/merge-scratch/combined diff A/main B/main` for both files). The only differences in the Excel calls are A's local staging in S1 and F, A's shared "already open?" helper (`A:utils/excel_com.py`), the error types each `except` catches, and one unreachable `return` that B removed. The same two files also differ in three places that are not Excel calls: A imports pywin32 inside a guard (`A:power_query_extractor/core/report_processor.py:20-36`; B imports it without one, `B:power_query_extractor/core/report_processor.py:11-15`), A's Step 4 refuses each client when pywin32 is missing (`A:…/report_processor.py:66-77`), and A's Skip Refresh looks for the configured suffix (`A:…/report_processor.py:137-146`, `2aaa2b1`, 25-07-2026) where B's looks only for `_Refreshed_` (`B:…/report_processor.py:100-102`). None of these changes an Excel call, so each verdict below holds for **both** projects. Rule numbers follow the B audit's split into 20 single rules (X1.3.1); the A audit's numbers (C1-C18, X1.3.1) are given alongside.

| # | What Standard 12 asks (`Std` line) | A audit / B audit rule | S1 Step 3 making reports | S2 Step 4 refresh | S3 Step 4 reading | F openpyxl fallback | Evidence (A / B) |
|---|---|---|---|---|---|---|---|
| 1 | COM only for a workbook with Power Query that the app writes into or refreshes (`846-847`; table row `828`) | C1 / R1a | MET | MET | MET in letter: COM reads a copy the app has just refreshed. The A audit scores this NOT MET, because S3 only reads, and openpyxl could read (`Std:827`). Same code, different reading of the rule. | NOT MET: it writes into a template copy with openpyxl. The A audit files this under C2. | `A:…/report_processor.py:584, 622`; `B:…/report_processor.py:535, 573` (the cell read and the close without saving); F `A:core/excel_handler.py:282-307` / `B:core/excel_handler.py:231-255` |
| 2 | Check for a `customXml/item*.xml` part with `DataMashup` before an openpyxl save (`847-849`) | C2 / R1b | n.a. | n.a. | n.a. | NOT MET: no check anywhere. RAN (both audits): Power Query parts lost, function returned `True`. | RAN (C78) grep `DataMashup`, `customXml`: 0 / 0 |
| 3 | If found, use COM or stop; never let openpyxl save over it (`849-850`) | C2 / R1c | n.a. | n.a. | n.a. | NOT MET: saves the template copy (A: the staging copy, then copied to the report path; B: in place at the report path). The template file itself is never saved over. | `A:core/excel_handler.py:306, 308`; `B:core/excel_handler.py:255` |
| 4 | `DispatchEx("Excel.Application")`: a separate Excel (`851-852`) | C3 / R2a | NOT MET: `GetActiveObject` first, then `Dispatch` | NOT MET (`Dispatch`) | NOT MET (`Dispatch`) | n.a. | `A:core/excel_handler.py:124, 127`; `A:…/report_processor.py:218, 536` / `B:core/excel_handler.py:111, 114`; `B:…/report_processor.py:174, 487` |
| 5 | `Visible = False` (`853`) | C4 / R2b | MET in form, but when S1 attached to your Excel, your Excel is hidden and never shown again | NOT MET: visible and maximised, on purpose, because the key presses need a window | MET | n.a. | `A:core/excel_handler.py:131`; `A:…/report_processor.py:225-227, 537` / `B:core/excel_handler.py:118`; `B:…/report_processor.py:181-183, 488` |
| 6 | `DisplayAlerts = False`, and still False when `Quit()` runs (`853-855`) | C5 / R2c | NOT MET: set back to True just before `Quit()` | MET, with a gap: after five failed tries it "continues anyway" | MET | n.a. | `A:core/excel_handler.py:132, 240, 244`; `A:…/report_processor.py:226, 538` / `B:core/excel_handler.py:119, 205, 209`; `B:…/report_processor.py:182, 489` |
| 7 | Never take over the user's own open Excel (`853-854`) | C6 / R2d | NOT MET (deliberate: the log says "Using existing Excel instance") | NOT MET (`Dispatch` may attach to it; MEDIUM-HIGH, cannot be tested here) | NOT MET (same) | n.a. | RAN (both audits, fake COM): `Quit()` called on "your" Excel while its unsaved workbook was open |
| 8 | `Quit()` in a `finally` block, even after an error (`856`) | C7 / R3 | NOT MET: `ScreenUpdating = True` and `DisplayAlerts = True` sit unprotected before it; if Excel has died, they raise and `Quit()` and `CoUninitialize()` are skipped | MET in structure | MET in structure | n.a. | `A:core/excel_handler.py:238-246` / `B:core/excel_handler.py:203-211`; RAN (both audits, fake COM) |
| 9 | Open with `IgnoreReadOnlyRecommended=True` (`857`) | C8 / R4a | NOT MET | NOT MET (last fallback only) | NOT MET (same helper) | n.a. | `A:…/report_processor.py:781`; `B:…/report_processor.py:729`: the only uses |
| 10 | If `workbook.ReadOnly`, close it and ask the user to close it first (`857-859`) | C9 / R4b | NOT MET | NOT MET | NOT MET (and the screen says "read-only mode" for a read-write open) | n.a. | RAN (C78) grep `.ReadOnly`: 0 / 0 |
| 11 | `BackgroundQuery = False` on each connection (`860-861`) | C10 / R5a | n.a. | NOT MET | n.a. | n.a. | grep: 0 / 0 |
| 12 | `workbook.RefreshAll()` (`861`) | C11 / R5b | n.a. | NOT MET: `SendKeys` Alt+A, R, A ⚠ into whichever window has focus | n.a. | n.a. | `A:…/report_processor.py:278-293` / `B:…/report_processor.py:238-248` |
| 13 | `CalculateUntilAsyncQueriesDone()`, and only then save (`861-865`) | C12 / R5c | n.a. | NOT MET: fixed waits, then `wb.Save()` whatever happened | n.a. | n.a. | `A:…/report_processor.py:297-349` / `B:…/report_processor.py:253-305`; about 39 s per report at the default wait (RAN (both audits)) |
| 14 | No connection still `Refreshing` before saving (`865-866`) | C13 / R5d | n.a. | NOT MET | n.a. | n.a. | grep: 0 / 0 |
| 15 | Never `RefreshDate` as proof of a refresh (`866-867`) | C14 / R5e | n.a. (the A audit scores MET: not used) | MET in letter. The substitute proof, "`Info!BB2` not empty, not 0, not an error" [GST RULE] ⚠, cannot tell a new value from one already in the file. | n.a. (A audit: MET) | n.a. | `A:…/report_processor.py:638-685` / `B:…/report_processor.py:589-636` |
| 16 | `Queries.Add` works; loading into a new table failed; ask first (`868-871`) | C15 / R6 | n.a. | n.a. | n.a. | n.a. | never used |
| 17 | `pythoncom.CoInitialize()` first in a worker thread (`872-873`) | C16 / R7a | MET | MET | MET | n.a. | `A:core/excel_handler.py:120`; `A:…/report_processor.py:215, 533` / `B:core/excel_handler.py:107`; `B:…/report_processor.py:171, 484` |
| 18 | The Excel work in a separate function (`873-874`) | C17 / R7b | NOT MET | NOT MET | NOT MET | n.a. | each session is one function |
| 19 | `CoUninitialize()` after that function returns, so no Excel object is still held (`874-875`) | C17 / R7c | NOT MET | NOT MET | NOT MET | n.a. | `A:core/excel_handler.py:250`; `A:…/report_processor.py:412, 633` / `B:core/excel_handler.py:215`; `B:…/report_processor.py:363, 584` |
| 20 | Excel on every PC, a logged-in user, never a scheduled job (`876-877`) | C18 / R8 | MET (interactive only; nothing checks that Excel is installed) | MET, but it also needs an unlocked desktop that nobody touches for about 33-39 s per report | MET | n.a. | no scheduler code in either |

**Counts, the same for A and B.** S1: 4 met, 8 not met, 8 do not apply. S2: 6 met, 11 not met, 3 do not apply. S3: 6 met, 6 not met, 8 do not apply. F: all 3 rules that apply are not met.

**Where the two audits score the same code differently** (this table uses one scoring for both):
1. Rule 1 for S3: the A audit says NOT MET, the B audit MET.
2. Rule 1 for F: the A audit says "not applicable" and puts the failure under C2; the B audit says NOT MET.
3. Rule 15 for S1 and S3: the A audit says MET, the B audit "not applicable".

None of these is a difference in the code.

**Things beyond Standard 12, in both** (the same shared code):
- The error-dialog closer looks at every window on the desktop and clicks "OK", "Close", "Yes" or "Continue" in other programs' dialogs too (A-G55, B-G61; RAN (both audits), fake desktop).
- The key presses are sent even when no workbook could be opened (A-G73, B-G62).
- Nothing times out, so a hung Excel blocks the job for ever. A's Step 4 cannot be stopped at all; B's stops only between items (A-G50, B-G39).
- Steps 3 and 4 can drive Excel at the same time (A-G71, B-G74).
- No code kills an Excel process (no `taskkill`, `psutil` or `TerminateProcess`; READ in both audits).

**History.** Every COM pattern above is already in `edac6a3` (31-08-2025, shared): `GetActiveObject` then `Dispatch`, the always-true quit test, `Visible = True` with `SendKeys`, the `BB2` check, and `IgnoreReadOnlyRecommended` only in the last fallback (READ (both audits) `core/excel_handler.py:122-125, 218@edac6a3`; `power_query_extractor/core/report_processor.py:154-228@edac6a3`). `bd176bf` (30-10-2025, shared) made the wait configurable. Before it, a comment read "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop, which suggests the wait was tuned by trial (MEDIUM, B audit X1.4.5). Neither project changed how Excel is driven after the split.

### 7.3 Notes on the rows

**Python version.** "3.7 or higher" has not changed since `edac6a3`, and no 3.7 interpreter was ever shown to run either project. The two audits show what the code itself needs (grammar 3.7), what today's packages need (3.8 to 3.10 ⚠) and what was actually used (3.12/3.13 in 2025; in July 2026 A's commit messages name the free-threaded `python3.14t` on your PC, but not the version of the `python.exe` that ran with pywin32 and the tests, MEDIUM that it is also 3.14 ⚠). Your Standard 13 settles the version for the new tool (`Std:931-932`). The one technical trap that exists in both old versions is that a client name ending in a dot is cleaned differently on 3.14. With a single shared runtime every PC gives the same answer.

**Window toolkit.** Standard 12 names PySide6-Essentials pinned to 6.12.x while any PC runs Windows 10 (`Std:833`) ⚠. The CLAUDE.md Tech Stack allows PyQt5 only "for an existing PyQt5 app that is being kept as it is" (`Std:209`). A uses Tkinter, which Standard 12 does not list for windows. B uses PyQt5 on Qt 5. Nothing in either project uses Qt's chart module, so `Std:840` (PySide6-Addons) does not arise. B's Qt 5 runtime is its own 50 MB copy (`PyQt5-Qt5`). If PySide6 also goes into the shared runtime, `_runtime` would hold two Qt copies, and both are locked while any app runs (`Std:1051-1054`; MEDIUM, B audit 9.9).

**Power Query protection.** A's July 2026 work shows the intent clearly: `c463256` guarded the imports so that the app would start without pywin32. Ten minutes later `be6418c` reversed the view and made pywin32 essential, "so reports cannot lose Power Query". Two texts in A still call pywin32 optional (`A:power_query_extractor/core/report_processor.py:11-13`; `A:gui/tabs/extract_tab.py:21-23, 101-105`). B reached the same "cannot start without pywin32" state by accident in `c258047`, with no message that explains it. In both, the openpyxl fallback is still present and still reports "Success" after stripping the queries (RAN in both audits on made-up workbooks with an injected `DataMashup` part). The same save also drops shapes drawn on a template sheet, such as text boxes and buttons, and still reports success: both audits' completeness checks RAN this on a made-up template with one text box (A audit Appendix X1.4.1 and gap A-G7; B audit Appendix X1.4.3 F5 and gap B-G32). Standard 12's table names shapes as well (`Std:827`), and a `DataMashup` check would not detect this loss. Whether your real templates hold drawn shapes is NOT SURE (how to check: open each template and look for buttons or text boxes on its sheets). [Corrected during review - the shapes finding was added from the two final audits' completeness checks, which came after this section was written.] Standard 12's check (`Std:846-850`) exists in neither.

**How the app starts.** Under Standard 13's launcher (`pythonw.exe`):
- A stops at `A:main.py:17` before any window, log or message (RAN (C78)).
- B gets further. A missing pywin32 or a damaged settings file then ends it with nothing on screen, because its errors go only to `print()` (B-G34; RAN (B audit) M1, C2, C3).

Neither has the `report_fatal_error` hook of `Std:984-1029`. Both projects' workers already catch their own errors and show them in the window, which is close to what `Std:1036-1037` asks ("Worker threads send their errors back to the window"):
- A: `A:gui/handlers/processing_handler.py:522-525`; `A:power_query_extractor/gui/extractor_window.py:957-960`. A's Step 4 sends the message box to the window's thread through its `_ui` helper (`A:power_query_extractor/gui/extractor_window.py:989-1002`). A's Step 3 worker calls `messagebox.showerror` itself, from the worker thread (`A:gui/handlers/processing_handler.py:525`; A-G70), so it does not "send the error back" in the standard's sense (READ).
- B: the `run()` methods in `B:gui/controllers/*_controller.py`, which send an `error` signal that `B:gui/pyqt_main_window.py:318-322, 412-415, 563-565` turns into a message box.

**Where files live.** Both put the log in the start folder (`logging.FileHandler('gst_organizer.log')`, `A:main.py:24`, `B:main.py:20`). This line is unchanged since the shared code, so under the launcher both would write one shared log into `\\server\share\Apps\<App>\` (HIGH; `Std:962` makes the app folder the start folder). The log is written in the PC's default encoding, so on a cp1252 PC lines containing ✓ or ❌ are dropped (A-G56, B-G44; RAN (both audits), simulated).

**Dependencies.** A's start-up check and installer were built after a real failure: the app died with "No module named 'pythoncom'" and a message that blamed missing source files (`A:utils/requirements_check.py:3-8`). B's start-up message today recommends exactly the command that does not fix that failure (`B:main.py:51-53`).

### 7.4 What the table means for "Annual Statement 2.0" (technical only)

1. **Excel layer.** Every Standard 12 failure sits in code the two projects share from the first commit, and neither changed it. So neither version offers a compliant Excel layer to keep. Writing it new to `Std:844-877` costs the same whichever version is kept (HIGH). The key-press refresh, and what counts as proof that a refresh worked (`Info!BB2` [GST RULE] ⚠), are open questions in both audits (A audit 9.10 Q1; B audit X1.9).
2. **Power Query protection.** Neither old version has the `DataMashup` check. The rule in `Std:846-850` would replace both A's leaky start-up gate and B's accidental one (HIGH).
3. **Window toolkit.** If you choose PySide6 (`Std:833`), B's Qt structure (views, controllers, signals, worker threads) is the nearer starting point. A's Tk windows would need a full redraw (MEDIUM; the porting effort was not measured).
4. **Start-up.** Neither has the launcher or the error hooks. B's guarded console line and its signal-based worker errors are nearer to `Std:984-1038`; A's start-up check and install offer conflict with `Std:924` and `Std:937-940` if a shared runtime is used (HIGH).
5. **Settings and logs.** Both need moving to `%LOCALAPPDATA%\<App>\` (`Std:1039-1040`). Whether B's run history, which keeps extracted figures, is wanted at all is your decision (B-F17, B-F25).
6. **Tests.** Only A has tests. 8.1 shows which parts of A's suite also run against B's shared logic.

## 8. Code health

This section compares how sound the two code bases are as programs: how well they are tested, how they handle errors, how much code is dead, how many bugs give wrong results without any warning, and how safely they treat files. The last column of each table says which project is healthier on that point and why, on technical grounds only.

### 8.1 Tests

| Point | A | B | Healthier, and why |
|---|---|---|---|
| How many | 18 test files, 183 test functions, 288 tests (one function can run several times with different made-up inputs). RAN (C78) `--collect-only`: 288. | None, now or in any of its 17 commits (RAN (C78): `git ls-tree` and `git log --name-only` show no test file; 0 `def test_` or `import pytest` in the code). | **A**: B has nothing that notices a regression. |
| When added | All after the split: from `6c1e1c6` (23-07-2026, 81 tests) to `bf2f1dd` (27-07-2026, 288). | — | — |
| Result on this computer | RAN (C78), virtual display: **1 failed, 287 passed in 3.29 s**. The failure expects Windows path splitting (`A:tests/test_setup_validation.py:65-66`) and would pass on Windows (HIGH, A audit X3.2.1). RAN (C78), no display: **14 failed, 241 passed, 33 skipped** (the window tests need a screen). Both match the A audit exactly. | Nothing to run. | **A**. |
| What the tests cover | File-name recognition [GST RULE] ⚠ (21 tests); name cleaning, state codes and client keys (47); which files the scan accepts (22); the whole scan with grouping and completeness [GST RULE] ⚠ (5); "is this workbook already open?" with fake COM (3); Step 3 cannot start twice (2); the Setup checklist (19); the Setup tab built in a real window (13); dark mode (11); the Step 4 panel (10); the standalone extractor (6); the main window's tabs (7); starting without pywin32 (4); the start-up check (15); the install offer (13); refreshed-file names (19); organise-only versus full pipeline (7); scroll bindings (64). (A audit X3.1, X3.3.) | — | — |
| What they do not cover (A audit X3.5) | Organising files (`core/file_organizer.py`: 0 of 12 functions entered). Report creation (`core/excel_handler.py`: 1 of 15). The consolidated workbook (`data_consolidator.py`: 0 of 6). The refresh and reading (`report_processor.py`: 3 of 15). `main.py` is never loaded. No Standard 12 rule is tested. The 9-column consolidated mismatch [GST RULE] ⚠ (A-G1) is caught by no test. Features with no test at all: A-F8, A-F10 to A-F16, A-F21, A-F23, A-F24, A-F26, A-F27, A-F30. | — | — |
| Do the tests catch old bugs? | The A audit re-introduced 30 single bugs, one at a time ("mutation checks"): **14 caught, 16 survived** (A audit X3.4; four repeated and confirmed in verification). Survivors include the Skip Refresh lookup (M21), the per-client folder key (M16), one state-table entry [GST RULE] ⚠ (M23) and GSTR-3B duplicate flagging [GST RULE] ⚠ (M24). | — | — |
| Side effects and hygiene | One test rewrites your real `~/.gst_organizer_cache.json` when it holds `"dark_mode": true` (A-G60; RAN (A audit)). Three tests start a second Python without `-B` and write 38 compiled files into the project folders (RAN (A audit); with `PYTHONDONTWRITEBYTECODE=1` none were written in the C78 runs). `tests/README.md` is stale ("No GUI … <1s"). There are weak or vacuous tests (`A:tests/test_requirements_check.py:27-32` can never fail). `pytest.ini` lacks `pythonpath = .` (`Std:945`). | — | — |
| Do A's tests work on B's code? | RAN (C78): A's three pure-logic files (`test_file_patterns.py`, `test_helpers.py`, `test_scan.py`) run against B: **7 failed, 66 passed**. Four failures are IMS Reco names with a browser "(1)" after the date. A accepts them since `5c8ce95` (23-07-2026, "Detect IMS Reco files with a trailing (N) duplicate suffix"); B's pattern has no "(N)" part; it is unchanged since `edac6a3` (31-08-2025, shared; `B:utils/constants.py:23`, A's is `A:utils/constants.py:23`). Worked example (the made-up name A's own tests use): `ImsReco-ABC-Delhi-01012024 (1).xlsx` is recognised as an IMS Reco file by A (client `ABC`, state `Delhi`, date `01012024`) and not recognised by B (RAN (VERIFY-C78) `ims_example.py`). That pattern is a [GST RULE] ⚠ classification; which is wanted is your decision. The other three failures are name-cleaning tests: B still replaces "Private" and "Limited" as plain, case-sensitive text (`B:utils/helpers.py:78-79`), which A changed to whole-word, any-case matching in `33d0baf` (23-07-2026; `A:utils/helpers.py:92-93`). `test_file_ops.py` cannot load on B, because B has no `scan_excel_files`. So A's tests mark exactly the places where the shared logic now differs. | — | — |
| What a test would have caught in B | — | The B audit (X3.5) lists 8 defects that a simple test would have caught on the day they came in. Examples: the window cannot start without pywin32 (`c258047`, B-G32); the summary's "Variations" sheet is always empty (`c258047`, B-G20); the settings import reads the wrong key (`c258047`); EXPORT LIST is never connected (`c258047`); a 27-minute window in which the program could not start (`a62d3cc` → `c21f788`, 11:29:56 to 11:57:02 on 07-02-2026; RAN `git log --date=iso`) [Corrected during review - originally "28-minute", as in audit B's Appendix X3.5; audit B's Sections 1, 2 and 12A give 27 minutes, which the commit times confirm]; white-on-white Browse buttons (`7e09e5c`, B-G67); Step 4 choosing the run folder by name (`b0e7d0e`, B-G12). The 9-column mismatch [GST RULE] ⚠ dates from the shared `47a1492`. | **A**. |

**Verdict: A is clearly healthier on tests.** A's suite is narrow: it does not cover the organiser, the report builders, the consolidated workbook or any Excel session. But it exists, runs in about 3 seconds, and its pure-logic part already runs against B's shared code. B has never had a test.

### 8.2 Error handling

An *exception handler* is the `except` block that runs when an error happens. A *bare* `except:` names no error type, so it catches every error, even Ctrl+C.

| Point | A | B | Healthier, and why |
|---|---|---|---|
| Handler counts (RAN (C78) `scripts/excepts.py`, program files only) | 143 `try` blocks, **147 handlers**, **23 bare `except:`**, 102 `except Exception`, 41 whose body is only `pass`, 7 with `finally`. | 90 `try` blocks, **91 handlers**, **0 bare**, 79 `except Exception`, 15 whose body is only `pass`, 3 with `finally`. `c258047` (07-02-2026) replaced 26 bare handlers ("Fixed all 26 bare except: clauses"). | **B**: no catch-everything handlers. |
| Where the error ends up (each audit read every handler) | 68 of 147 (46%) leave no trace at all: 42 swallow the error and 26 quietly use a substitute value. 124 of 147 (84%) never show anything to the user; 23 do (A audit 6.8.1). | 45 of 91 never reach the screen although the job carries on; 3 print to the console only; 15 sit in dead code; 27 reach the screen directly (22) or through a re-raise (5); 1 would show a red line but cannot be reached in practice (B audit 6.3; 45 + 3 + 15 + 27 + 1 = 91). | **B**, MEDIUM: the two audits classified with different schemes, so the shares are only roughly comparable. |
| The style | One established idea in both, from the shared core: catch per file or client and hand back a "failure value", so one bad item never stops the batch. Five failure styles coexist: a status dictionary, True/False, None, raising, and swallowing (A audit 8.7 P1). | The same idea and the same mix: True/False, a status dictionary, "Success/Failed/Skipped", "OK/FAIL", "✓/✗" (B audit 8.7.1). | Equal. |
| Errors nothing caught | No error hook (RAN (C78) grep `excepthook`: 0). An error inside a button callback is printed by Tk to the console (invisible without one); the click does nothing and the program carries on (A-G30). | No error hook (RAN (C78) grep: 0). An error inside any window handler ends the whole program with exit code 134, without saving (B-G35; RAN (B audit) MW e10). Only `_on_scan_completed` wraps its whole body in a `try`. | **A** on this point: one failed click does not kill the window. Both fall short of `Std:984-1029`. |
| Errors in worker threads | Both workers catch their own errors and show them (`A:gui/handlers/processing_handler.py:522-525`; `A:power_query_extractor/gui/extractor_window.py:957-960`). But the Step 3 worker also touches the window directly: 45 `log_message` calls, message boxes and button changes from the worker thread (A-G70). | Every worker catches everything in `run()` and sends an `error` signal that the window turns into a message box. Workers never touch the screen (RAN (B audit): all scan-log and status-bar writes on the main thread). | **B**: this matches `Std:1036-1037`, and Qt signals are the safe way to talk to the window. |
| Start-up errors | Console `print()` plus `_pause()`. Without a console it dies before the log exists (A-G22; RAN (C78)). | `print()` only, so they are invisible without a console (B-G34). | **B**, narrowly (it at least starts). |
| Excel clean-up | Two unprotected lines in Step 3's `finally` (`A:core/excel_handler.py:239-240`): if Excel died, `Quit`, `CoUninitialize` and the staging-file delete are skipped (A-G28; RAN (A audit), fake COM). | The same two lines (`B:core/excel_handler.py:204-205`; B-G38; RAN (B audit), fake COM). | Equal. |
| Messages to the log file versus the screen | Step 4's `_log` writes every line at INFO and never passes a level, so even "❌" lines appear in the plain colour (`A:power_query_extractor/core/report_processor.py:53-60`). The Step 3 worker sends most problems to the screen only and only exceptions to the log file. | The same `_log` (`B:power_query_extractor/core/report_processor.py:30-37`). Many `logger.warning` / `logger.error` lines, including the text of a detected Excel error dialog, reach only the log file (B audit X1.3.7). | Equal. |
| All Python warnings switched off | `warnings.filterwarnings('ignore')` at import (`A:core/excel_handler.py:17`; A-G61). This also hides openpyxl's warnings about features it is about to drop. | The same line (`B:core/excel_handler.py:17`; B-G31). | Equal. |

**Verdict: B is healthier on error-handling style, narrowly.** It has no bare `except:`, a larger share of its handlers reach the screen, and its workers report through Qt signals. A is more forgiving of an unexpected error in a button. Both lack the error hooks and message box of Standard 13, and both stop in the same Excel clean-up.

### 8.3 Dead code

*Dead code* is code that exists but is never used by the running program. Counts below are from each audit's Section 2.8, where every item was first found by a script and then read by hand. The audits used different scripts, and B counts some levels as table rows, so small differences are not meaningful (MEDIUM).

| Level | A (A audit 2.8) | B (B audit S2 2.8) | Notes |
|---|---|---|---|
| Dead program files (never imported) | **0** | **3**: `models/signals.py`, a "signal hub" that nothing imports and whose docstring describes an architecture B does not use; `power_query_extractor/extractor_main.py`, which cannot even be imported; `power_query_extractor/gui/__init__.py`. Plus **1** file loaded but never used (`gui/widgets/collapsible_group.py`). | B only: `c258047`, `a62d3cc`. |
| Dead classes | **0** (`UIHelpers` is never created on purpose: a box of static helpers) | **3**: `AppSignals` (with 18 dead signals), `CollapsibleGroup`, `PowerQueryExtractor` | — |
| Functions and methods never called | **16** | **34** (25 in live code, 9 inside the dead class and file) | 9 shared names checked in both (RAN (C78) grep: only the definition, or the definition plus a re-export line): `cleanup_on_error`, `get_summary`, `validate_results`, `get_date_only`, `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`, `_write_section`, `_create_details_sheet`. All are dead since `edac6a3`. |
| Code that runs but can never act, cannot be reached, or never runs with the shipped settings | **17** | **22** | Shared: the `ExcelHandler.__del__` quit; the always-true "Don't quit if Excel was already open" test (`A:core/excel_handler.py:241-246`; `B:core/excel_handler.py:206-211`); the calculated-column branch [GST RULE]; the "description" branch; the progress callback; the unreachable "Open method failed" branch. B adds the openpyxl fallback (unreachable through the window) and the per-client folder-name option that nothing can switch on (B-F5). |
| Dead constants and module-level names | **22** | **13** rows | B's `GUI_CONFIG` and `MESSAGES` died with the Tkinter files (`a62d3cc`); A still uses parts of them. |
| Dictionary entries never read | **13** | **17** rows | — |
| Stored values never read | **20** kinds | **19** kinds | — |
| Parameter findings | **21** rows (5 never used, 11 rows never supplied, 5 always supplied) | **28** rows | — |
| Unused imports | **26** import lines | **21** rows | — |
| Unused local variables | **6** | **5** | — |
| Used only by tests | **8** | **0** (no tests) | — |
| Dead screen elements, style rules, leftover texts | not a separate level in A (for example the unused card and border colours are counted under dictionary entries) | **7** | — |
| Dead code inside the test folder | **5** | — | — |

**Shared versus own.** Most of A's dead code has been dead since `edac6a3` and was inherited by B (A audit 2.8.13). B still carries most of it (its `c258047` removed a few shared items, among them the unreachable `return` in the refresh session and the empty `logs\` and `temp\` folders; READ `git diff A/main B/main`), and adds its own:
- the never-imported `signals.py` and the unused `CollapsibleGroup`;
- `load_session_state` and `get_recent_source_folders`, so the session file is written but never read;
- `ProcessingView.reset` and `ThemeManager.toggle`;
- helpers left behind by its own commits: `get_file_info` and `validate_excel_file` (dead since `7e09e5c`), `truncate_path` and `ReportProcessor.cleanup` (dead since `a62d3cc`);
- the `%TEMP%\gst_excel_temp` folder it creates and never uses.

A's own dead items (A audit 2.8.13, "Became dead or test-only in A's own commits") are fewer:
- the "needs pywin32" card that can no longer appear (`c463256`);
- an unused import, `DEFAULT_SUFFIX_PATTERN` (`2aaa2b1`);
- one stored colour, `_original_root_bg`, and the never-darkened `THEMED_OPTIONS` entries (`aeb83cf`);
- the `card` and `border` colour keys (`169c7b5`);
- the Tk delegator method `PowerQueryExtractorApp.__getattr__` (`a677670`);
- `_default_exists` (`322717f`);
- helpers that only the tests still reach: `find_excel_files` and the "no stat given" branches (`56aaad9`), `validate_excel_file` (`4f84ba9`) (RAN (VERIFY-C78) grep: no program call of `find_excel_files(` or `validate_excel_file(` in A).

The empty `logs\` and `temp\` folders are not A's own: `main.py` has created them since `edac6a3` (shared; `main.py:36@edac6a3`). A kept that code (`A:main.py:35-43`); B removed it in `c258047` (READ).

**Verdict: A is healthier on dead code.** It has no dead files or classes, about half as many never-called functions, and no broken entry point. B's dead "signal hub" and its unimportable extractor file mislead anyone reading the code.

### 8.4 Silent bugs

Both audits' Section 7 lists every gap (a defect or an unhandled situation). A *silent* bug runs without any error but gives a wrong result or hides a problem.

**Counts by severity** (each audit's 7.10):

| Severity | A (79 gaps) | B (82 gaps) |
|---|---|---|
| HIGH | 13 | 20 |
| MEDIUM-HIGH | 1 | 3 |
| MEDIUM | 35 | 39 |
| LOW-MEDIUM | 6 | 5 |
| LOW | 24 | 15 |

The two audits used the same words but rated independently. Some identical defects carry different ratings: one client split by spelling or state text is MEDIUM in A (A-G35, A-G36) but HIGH in B (B-G16). "Success" when both reports failed is LOW-MEDIUM in A (A-G19) but HIGH in B (B-G19). So the raw counts overstate the difference. The cross-project map (`/root/merge-scratch/compare/MAP.md` §3) pairs every gap: **98 rows, 73 the same defect in both (some only in part), 16 only in A, 9 only in B**.

**The silent-logic scan side by side** (each audit's 7.7):

| Kind of silent bug | A | B |
|---|---|---|
| A function made in a loop that sees only the last value | none live (fixed in `0890022`, A only) | none |
| A shared default value that is changed between calls | none | none |
| Off-by-one (one too many or one too few) | cosmetic: the final ETA is one client behind (A-G72); display helpers miscount (A-G75). Also listed there: a typed folder-name length of 8 or less drops the state code from the client folder (A-G21) [GST RULE]; a typed wait of 0 waits nothing (A-G26). A's code has the same cosmetic extended-wait counter as B (`A:power_query_extractor/core/report_processor.py:327-329`, identical to `B:…:283-285`), which A's audit does not list (READ) | cosmetic: the extended-wait counter (B-G79). B's spin box cannot go below 15, so the state-code cut is unreachable from the screen (B audit 7.7) |
| Text compared where a number or date was meant | client or state text that looks like a number breaks the Tk table's keys (A-G25) | the "latest" run folder is chosen by its day-first **name**, so Step 4 can show and write into an older batch (B-G12, HIGH); numbers sort as text (B-G73) |
| A value that never takes the value the code relies on | `self.excel_app` is never set, so the user's Excel is always quit (A-G54) | the same (B-G63) |
| State shared between threads | a re-scan empties the data the Step 3 worker uses (A-G13); inputs are re-read live during a run (A-G14) | the worker rewrites the window's scan data (B-G75); "Skip refresh" travels through a module-wide setting (B-G76). Both are also in A's code, although A's audit does not list them under this heading (MAP §3: B-G75 pairs with A-G19 in part; B-G76 "same defect in both"; READ: A writes `EXTRACTION_OPTIONS['skip_refresh']` when the tick changes, `A:power_query_extractor/gui/extractor_window.py:486-487`, B at every Start, `B:gui/controllers/extraction_controller.py:54-55`) |
| Window touched from a worker thread | yes (A-G70) | no |
| Writer and reader disagree | the Annual Report folder is `'root'` in one table and the Version folder in the organiser (A-R16 [GST RULE]; the table's value is never read, so the organiser's value is what happens) | the same Annual Report `'root'` / Version-folder mismatch (B-R13 [GST RULE]; `B:utils/constants.py:58`, `B:core/file_organizer.py:294`); the settings import reads the wrong key (B-G24); extraction history is filed under a key no screen reads (B-G25); Skip refresh ignores a changed suffix (B-G9). A does not share B-G9 as such: since `2aaa2b1` (25-07-2026) A's Skip Refresh and status column look for the configured suffix (`A:power_query_extractor/core/refresh_naming.py:23-66`). A has a related weakness of its own: some suffix shapes make originals look refreshed, and an empty or `{timestamp}`-only suffix is never recognised (A-G16; MAP §3 pairs A-G16 with B-G9 "in part") |
| Cell or column wiring that differs from the program's own configuration, or that a template may not expect [GST RULE] ⚠ (technical description only; which cell or file is right is the author's call) [Corrected during review - originally "Wrong cell or column wiring"; reworded so that no GST rule is judged, as both audits' completeness checks did for the same items] | 9 consolidated columns show the `Diff GSTR-3B` J15/K15/L15 values (A-G1); only the first file of a type is linked (A-G6); Links cells for missing types keep the template's text (A-G5) | the same three (B-G1, B-G3, B-G2) |

**HIGH silent defects present in both (MAP pairs):**
- A-G1 / B-G1: 9 columns from another cell [GST RULE] ⚠.
- A-G2 / B-G7: a failed refresh is later trusted as refreshed.
- A-G3 / B-G6: the refresh cannot be confirmed [GST RULE] ⚠ (`Info!BB2`).
- A-G4 / B-G4: a template without a `Links` sheet "succeeds".
- A-G8 / B-G15 and B-G14: two clients, or two files, end up under one name [GST RULE].
- A-G10 / B-G13: "Resume" does not resume.
- A-G33 / B-G17: a hyphen in a client name mis-splits the file name [GST RULE] ⚠.
- A-G54 / B-G63: the user's Excel is taken over.
- A-G55 / B-G61 and B-G62: keys and clicks reach other programs.

Further HIGH gaps that are in both projects but are not silent wrong results (MAP §3 pairs; ratings from each audit's 7.10): A-G22 / B-G34, start-up without a console (HIGH in both; 7.1); A-G27 / B-G40, closing the window during a run kills the work in progress (HIGH in A, MEDIUM in B); A-G44 / B-G51, Excel set-ups the code does not expect, such as another Excel language or a locked desktop ⚠ (HIGH in A, MEDIUM in B); A-G50 / B-G39 (with B-G57), nothing times out and Step 4 cannot be stopped or stops only between items (HIGH in B, MEDIUM in A). B also rates three more shared defects HIGH that A rates lower: B-G16 (one client split by spelling or state text), B-G19 ("Success" when both reports failed) and B-G20 (the problem-file list never reaches the user; in part only B, see below).

**Only in A (HIGH):** A-G7, reports lose Power Query when pywin32 is present but broken (likelihood LOW-MEDIUM). A's other only-A gaps are MEDIUM or lower: A-G13, A-G14, A-G25, A-G26, A-G31, A-G49, A-G53, A-G57, A-G59, A-G60, A-G62, A-G68, A-G70, A-G72, A-G75.

**Only in B (HIGH):**
- B-G5: a failed report leaves the unfilled template copy under the report's name, and Step 4 later refreshes it.
- B-G12: Step 4 picks an older batch.
- B-G32: B cannot start without pywin32 (not a wrong result: the start fails with a console message; under `pythonw` that message is invisible, B-G34).
- B-G67: three invisible Browse buttons (visible, not silent).

B's other only-B gaps are B-G25, B-G37, B-G41, B-G60 and B-G73. In part only B: the problem-file list never reaches the user at all, because B's summary sheet is always empty (B-G20, a regression in `c258047`; A shows the list in its summary workbook, A-G67).

**Verdict: about even, with A healthier by a narrow margin.** Three quarters of the defects are the same shared code. Of the silent ones that differ, B adds two that give wrong output without warning: the unfilled report (B-G5) and the older batch (B-G12). It also has the always-empty problem list (B-G20). A's one HIGH silent extra (A-G7) needs a damaged pywin32 install to happen. On the other side, B's Qt structure removes A's thread-safety and Tk-table defects (A-G13, A-G25, A-G26, A-G70).

### 8.5 File safety

| Point | A | B | Healthier, and why |
|---|---|---|---|
| The originals (client input files and templates) | Never changed: inputs are only read (8 bytes) and copied; templates only copied (A audit X1.1 #9, READ and RAN). | Never changed (B audit X1.1 #13: SHA-256 fingerprints of all 17 made-up source and template files identical after a full Step 1-4 run, RAN). | **Equal.** RAN (C78) grep of every delete or rename call: in A, only staging files (`A:core/excel_handler.py:150, 258, 276, 320`), a copy whose size does not match (`A:utils/helpers.py:207`) and the never-called `cleanup_on_error` (`A:core/file_organizer.py:392`); in B, only the last two (`B:utils/helpers.py:177`; `B:core/file_organizer.py:375`). No `os.replace` or `shutil.move` in either. |
| Where Step 3 builds each report | On local disk: the template is copied to `%TEMP%\gst_excel_temp\<report name>`, filled and saved there, then copied to the target in one write (`A:core/excel_handler.py:147-155, 212-215`). Since `12d505f` (23-07-2026, "Build Excel reports on local disk, copy to network once"). | Straight at the target: the template is copied to the final report name first and Excel opens and saves it there (`B:core/excel_handler.py:122-129`). If filling fails, the unfilled copy stays under the report's normal name, and Step 4 lists and refreshes it like any report (B-G5; RAN (B audit), fake COM). | **A**: a failed fill never leaves a report-looking file behind, and the network sees one write instead of many. An interrupted final copy can still leave a partial file under the report's name (A-G12, which MAP §3 pairs "in part" with B-G26 and B-G43). |
| Temporary files | Staging files are deleted after each report, best effort (`A:core/excel_handler.py:255-260, 316-322`). One is left behind when the clean-up itself raises (A-G28; RAN (A audit) R2 g). The folder is never removed. | `%TEMP%\gst_excel_temp` is created at each processing run and never used (`B:core/excel_handler.py:52-53`; B audit X2.5 S9). | Equal: neither leaves temporary files inside client folders. |
| Copying client files | `safe_copy_file` compares sizes and deletes a copy whose size differs. An error in the middle of a copy leaves the partial file under the final name (`A:utils/helpers.py:172-215`; A-G11). | Identical function (`B:utils/helpers.py:142-185`; B-G26; RAN (B audit): 500 of 3,604 bytes left under the final name). | Equal. |
| Backups | In Fresh mode an existing destination file is first copied to `<stem>_backup_<DDMMYY_HHMM><ext>` in the same folder (`A:utils/helpers.py:217-234`, called at `A:core/file_organizer.py:280`). Because each run makes a new `Version-<DDMMYY HHMM>` folder, this happens only when a run repeats within one clock minute (A-G9), or when two source files of one client clean to the same name within one run (then Fresh keeps the second file plus a backup of the first; A-G8 in part, B-G14; MAP §3, RAN in the audits). The backups then sit inside folders the template's queries read. If the backup fails, the overwrite goes ahead anyway. | Identical (`B:utils/helpers.py:187-204`, called at `B:core/file_organizer.py:263`; B-G27). | Equal. |
| Workbooks the program creates itself | The summary and the consolidated workbook are saved by openpyxl straight onto their final names (`A:core/excel_handler.py:402`; `A:power_query_extractor/core/data_consolidator.py:179`). An interrupted save leaves a broken file under that name (A-G12; RAN (A audit): a 3,010-byte file openpyxl cannot open). | Same (`B:core/excel_handler.py:339`; `B:power_query_extractor/core/data_consolidator.py:167`; B-G43). | Equal. |
| Step 4 refresh | Works on a time-stamped copy, never on the Step 3 report (`A:…/report_processor.py:164`). It saves whatever happened (`:349`), so a failed refresh leaves a `_Refreshed_` copy that is later trusted (A-G2). An empty suffix makes the copy name equal the original; Python refuses ("same file"), and the original is untouched (A-G16, RAN). | Same (`B:…/report_processor.py:120, 305`; B-G7; empty suffix B-G9, RAN). | Equal. |
| Runs in the same minute | Overwrite earlier output silently (A-G9). | Same, and the `Extraction_Log` too (B-G27). | Equal; B slightly worse. |
| Settings and history files | One settings file, rewritten whole in place, with no temporary file and rename (`A:gui/utils/cache_manager.py:45-46`). A broken file silently loads as defaults. | Three files rewritten whole in place (`B:models/cache_manager.py:135-142`; the run history by `B:models/run_history.py:128-137`). A broken file loads as defaults and is then overwritten with defaults at the next save (RAN (B audit) C4). With two copies of the program open, the last save wins, run history included (B-G82). | **A**, narrowly: there is less at stake. Neither writes safely. |
| Writing into the program folder | Creates `logs\`, `temp\` and any missing `__init__.py` at every start (`A:main.py:35-43`). Writes the log there when started from that folder. | Only the log, and only when started from that folder (`B:main.py:20`). | **B**. |
| Your own open Excel | Step 3 takes it over, hides it and quits it with prompts on. Unsaved work in it is at risk (A-G54). | Same (B-G63). | Equal (both unsafe). |

**Verdict: A is healthier on file safety, narrowly.** Almost everything is the same shared code. A's local staging (`12d505f`) means a report whose filling failed never sits under a report's name (only an interrupted final copy can leave a partial file there, A-G12). B is better only in leaving the program folder alone.

### 8.6 Healthier on each point: summary

| Point | Healthier | Why (technical only) | Confidence |
|---|---|---|---|
| Tests | **A** | 288 tests that run in about 3 s, against none. Its pure-logic tests already run on B's shared code (7 of 73 differ, and each marks a real difference). | HIGH |
| Error handling | **B**, narrowly | No bare `except:`, more errors reach the screen, workers report through Qt signals. A survives an unexpected error in a button; B aborts. | MEDIUM |
| Dead code | **A** | No dead files or classes, about half as many never-called functions, no broken entry point. | HIGH |
| Silent bugs | **About even; A** by a narrow margin | Three quarters are shared. B adds the unfilled report (B-G5), the older batch (B-G12) and the always-empty problem list (B-G20). A's extra silent HIGH (A-G7) needs a broken pywin32. B's Qt structure removes A's thread and Tk-table defects. | MEDIUM |
| File safety | **A**, narrowly | Local staging of Step 3 reports. B does not write into the program folder. | HIGH |
| Closeness to Standards 12 and 13 (Section 7) | **Neither**; B nearer on the toolkit, start-up and logs; A nearer on tests and on stating its dependencies | The Excel layer is identical and misses the same rules in both. | HIGH |

### 8.7 Corrections and cross-checks found while writing Sections 7 and 8

The code was taken as the truth. Each point below was checked in both projects.

1. **The Tech Stack wording in the task brief.** The brief summarises the CLAUDE.md copy's Tech Stack as "PySide6 with the PySide6-Essentials package, pinned to 6.12.x while any PC runs Windows 10". The Tech Stack (`Std:208-209`) names PySide6-Essentials and adds "[— or PyQt5, only for an existing PyQt5 app that is being kept as it is]". The 6.12.x pin is in Standard 12's table (`Std:833`), not in the Tech Stack. Both lines are quoted in 7.1.
2. **Standard 12 scoring.** The two audits score the identical COM code differently in three places: rule 1 for Step 4 reading, rule 1 for the fallback, and the `RefreshDate` rule for Steps 3 and 4 reading (7.2). This comparison uses one scoring for both.
3. **B's line count.** The B audit's Section 9.9 says every start "compiles B's 7,805 lines". That is a newline count (`wc -l`). Counting last lines that have no line end gives 7,821, as B's Section 2.1 says (RAN (C78)). [Corrected during review - audit B's final self-review (its corrections 37-38) changed Sections 8.2.1 and 9.9 to 7,821 as well, so the final audit and this section now agree; the 7,805 stands only in the section file `/root/merge-scratch/audit/B/S89.md` that was read here.] A: 8,930 program lines and 2,163 test lines (RAN (C78); matches the A audit's 2.1).
4. **Confirmed unchanged from the audits** (RAN (C78)):
   - exception-handler counts in both (147 / 23 bare; 91 / 0 bare / 15 `pass`-only);
   - A's test results (288 collected; 1 failed / 287 passed with a display; 14 / 241 / 33 without);
   - A's pure-logic tests on B (7 failed / 66 passed; `test_file_ops.py` cannot load);
   - no `DispatchEx`, `RefreshAll`, `BackgroundQuery`, `CalculateUntilAsyncQueriesDone`, `.Refreshing`, `RefreshDate`, `.ReadOnly`, `DataMashup`, `customXml`, `excepthook`, `pythonw` or `LOCALAPPDATA` anywhere in either program;
   - A dies under no-console streams at `A:main.py:17`, and B loads and writes its log in the start folder.
5. **Corrected during verification (VERIFY-C78), each checked in the code of both projects:** the licence line (PyQt5 is GPL v3, its Qt library PyQt5-Qt5 is LGPL v3); the list of differences between A's and B's Excel files (it now also names A's guarded import, its "pywin32 missing" refusal and its configured-suffix Skip Refresh, none of which is an Excel call); "A has the same weakness" for B-G9 (A has looked for the configured suffix since `2aaa2b1`; its own suffix weakness is A-G16); A's own dead items (the empty `logs\`/`temp\` folders are shared code from `edac6a3`, and six A-only items were missing); B's handler classification (one handler was not counted); where backups arise (also when two source files clean to the same name); that A's staging still leaves a partial file if the final copy is interrupted (A-G12); A's Step 3 worker showing its error box from the worker thread (A-G70); the run-history file's writer (`B:models/run_history.py:128-137`); defects present in both but listed under one project only (B-G75, B-G76, the Annual Report `'root'` label B-R13, the extended-wait counter); and HIGH shared gaps that were missing from the list in 8.4.

### 8.8 Runs made for these sections

All runs were in `/root/merge-scratch/run/C78/`, with `python -B` and `PYTHONDONTWRITEBYTECODE=1`, and `HOME` and `TMPDIR` pointed into that folder. Afterwards `/root/merge-scratch/A-copy` and `/root/merge-scratch/B-copy` held no compiled files and no new files (checked with `find`). No client data was used.

| Run | What it checked | Output |
|---|---|---|
| `scripts/excepts.py` on both copies (venv-A Python, `-I`) | `try`, handler, bare, `Exception`, `pass`-only and `finally` counts from the syntax tree | `out/excepts.txt` |
| `xvfb-run -a … pytest -p no:cacheprovider -q -o addopts=""` on `A/` (a fresh copy of A-copy) | A's suite with a virtual display; `--collect-only` count | `out/pytest_xvfb.txt` (1 failed, 287 passed in 3.29 s; 288 collected) |
| the same without `DISPLAY` | A's suite with no display | `out/pytest_nodisplay.txt` (14 failed, 241 passed, 33 skipped) |
| `BwithAtests/`: a copy of B-copy plus A's `conftest.py`, `pytest.ini` and four pure-logic test files, venv-B Python | Do A's tests hold on B's shared logic? | `out/A_tests_on_B.txt` (7 failed, 66 passed; `test_file_ops.py`: ImportError `scan_excel_files`) |
| `scripts/noconsole.py` with `A/main.py` (venv-A) and `B-copy/main.py` (venv-B), each in its own empty start folder `cwdA/`, `cwdB/` | Loading `main.py` with no console streams, as `pythonw.exe` would, without calling `main()` | A: `AttributeError: 'NoneType' object has no attribute 'buffer'`, no log; B: loads, `gst_organizer.log` created in `cwdB/` |
| Grep over both copies (program files) | The COM and standard keywords listed in 8.7 item 4; delete and rename calls; Tkinter and PyQt5 importers (17 and 14 files); dead-function names | quoted in 7.1, 7.2, 8.3, 8.5 |
| `git -C /root/merge-scratch/combined` (read only) | `ls-tree` and `log` for test files in B; `diff A/main B/main` of `core/excel_handler.py` and `report_processor.py`; dates and subjects of `12d505f`, `064c634`, `ba2a753`, `c463256`, `f363946`, `be6418c`, `a5ad487`, `6c1e1c6`, `5c8ce95`, `33d0baf`, `c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e` | quoted in place |

**Verification re-runs (task VERIFY-C78).** Every RAN (C78) result above was run again on fresh copies in `/root/merge-scratch/run/VERIFY-C78/`, with `python -B`, `PYTHONDONTWRITEBYTECODE=1` and `HOME`/`TMPDIR` pointed into that folder: the handler counts (A: 143 `try`, 147 handlers, 23 bare, 102 `Exception`, 41 `pass`-only, 7 `finally`; B: 90, 91, 0, 79, 15, 3; `out/excepts.txt`); A's suite (288 collected; 1 failed / 287 passed in 3.03 s with a virtual display, the same single Windows-path failure; 14 failed / 241 passed / 33 skipped without a display; `out/pytest_xvfb.txt`, `out/pytest_nodisplay.txt`); A's three pure-logic test files on B (7 failed / 66 passed, the same seven tests; `test_file_ops.py` stops at the import of `scan_excel_files`; `out/A_tests_on_B.txt`); and the no-console load (A: `AttributeError: 'NoneType' object has no attribute 'buffer'`, no log; B: loads, empty `gst_organizer.log` in its start folder). All match. The run time differs slightly (3.03 s against 3.29 s), as run times do. The keyword greps of 8.7 item 4, the delete/rename grep of 8.5, the Tkinter/PyQt5 importer counts (17 and 14 files, 25 Tkinter import lines) and the dead-name greps of 8.3 were repeated with the same results; the line counts of 8.7 item 3 (A 8,930 program and 2,163 test lines; B 7,821 lines in 38 files) were recounted. New in the verification: `ims_example.py` (the IMS Reco worked example of 8.1, run against each project's own `utils/constants.py` on the copies; A: match, B: no match). A-copy and B-copy held no compiled or new files afterwards (checked with `find`).

## 9. Gaps both share

This section lists every problem that the two audits found (their Section 7, "gaps") and that exists in project A **and** in project B. A *gap* is a defect or an unhandled situation: something the program does wrong, or a real-world case it does not cope with. The list starts from the gap map in `/root/merge-scratch/compare/MAP.md` §3, which pairs 73 gaps as "same defect in both". Each pair was checked again here in the code of both projects. Most of them came in through the 12 commits the two projects share, so the new "Annual Statement 2.0" inherits them from either parent unless they are fixed.

### How to read this section

**The two projects.** A = `goyal-sidhant/AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026), the Tkinter window ("GST File Organizer v3.0", 288 tests). B = `goyal-sidhant/AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026), the PyQt5 window ("v4.0", no tests). The names mislead: A is not the older one, and B is not the newer or better one. Both share their first 12 commits (`edac6a3`, 31-08-2025 … `47a1492`, 30-10-2025) and split after `47a1492`. A *commit* is one saved change in a project's history, named by a short code such as `edac6a3`.

**Evidence marks.**

| Mark | Meaning |
|---|---|
| `A:path:line`, `B:path:line` | A place in A at `bf2f1dd` or in B at `b0e7d0e`. |
| **RAN (C910)** | Run for this section on made-up inputs, in `/root/merge-scratch/run/C910/` with `python -B`, once with A's code and once with B's code. Invented names only (ClientA, Alpha Test Traders, Beta Demo Pvt Ltd and similar). The scripts and their outputs are listed in 9.5. |
| **RAN (VERIFY-C910)** | Re-run or new run made while checking this section (23-10-2026), in `/root/merge-scratch/run/VERIFY-C910/` with `python -B`, on fresh copies of both snapshots. All RAN (C910) results in 9.1-9.3 were reproduced there; the scripts are listed in 9.5. |
| **RAN (A audit)**, **RAN (B audit)** | Run by that audit. The trace was re-read here and the code lines were checked again. |
| **READ** | Established by reading the code of both projects. |
| HIGH / MEDIUM / LOW after an inference | How sure an inference is when it could not be run. |

**COM** is Windows' way for one program to drive another; here Python drives Microsoft Excel. It cannot run on this Linux machine. So for Excel itself, everything is READ only. Where a row says "fake COM", the Python code was driven with stand-in objects that only record which calls are made. That proves the order of the calls, never how Excel would react.

**Severity groups.** Each audit rated its gaps HIGH, MEDIUM or LOW, and sometimes in between (MEDIUM-HIGH, LOW-MEDIUM). A pair is placed in the group of the **higher** of the two ratings. An in-between rating is rounded up (MEDIUM-HIGH counts as HIGH, LOW-MEDIUM as MEDIUM), so that nothing is under-rated. One exception applies to gaps that are shared only "in part": when the higher rating clearly rests on a part that only one project has, or on a part that has its own, higher-grouped row in this section (for example A-G19, whose single LOW-MEDIUM rating covers the parts listed as H10, M8, M9 and L1), the pair is grouped by the rating of the part in that row. Each such row then says so ("Grouped …"). The original ratings of both audits are always shown, so the author can regroup.

**"Came in".** This gives the commit that wrote the lines responsible, found with `git blame` (a git command that shows, for every line of a file, the commit that last wrote it) on both histories (RAN (C910), `scripts/blame.sh`). "Shared" means one of the 12 common commits, so both projects carry the same lines. "Separately" means the two projects reached the same defect through different code.

**GST rules.** A *GST rule* is anything that decides a tax figure; which return, table, sheet or cell a figure belongs to (for example GSTR-9 Table 6J or 8D); a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; or a due date. Such rows are tagged **[GST RULE]**. Because every row here is the same in both projects, there is no choice between two versions of a rule. The section only describes the mechanism and what the code does against its own configuration. Whether the rule itself is right is the author's decision. **⚠** marks values that change over time: return and table names, template cell positions, file-name conventions, the state list, path limits, Excel versions and Python versions.

**Words used below (first use).**

| Word | Meaning |
|---|---|
| Links sheet | The sheet in each ITC and Sales template into which the program writes folder paths and file names, for the template's queries to read. |
| Power Query, refresh | Excel's data-import tool. Its queries live inside the workbook, and a *refresh* re-runs them so that the report reads the newly copied client files. |
| Consolidated workbook | `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`, the Step 4 workbook that collects the figures of every client into one sheet per topic. |
| Refreshed copy | Before refreshing, Step 4 copies each report to `<report>_Refreshed_<DDMMYY_HHMM>.xlsx` and refreshes the copy. |
| Skip refresh | A Step 4 tick box: instead of refreshing, re-use an existing refreshed copy. |
| Run folder, Version folder, category folders | `Annual Statement-<DDMMYY HHMM>` (one per run); inside it `<client>-<state code>`; inside that `Version-<DDMMYY HHMM>`; inside that the three category folders (A: `GSTR3B`, `ITC`, `Sales`; B: `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`). |
| Fresh, Re-run, Resume | The three processing modes on Step 1. |
| Date modified | When the computer last recorded a change to a file or folder. |
| openpyxl | A Python library that writes `.xlsx` files without Excel. |
| Console | The black text window a program can print to. *pythonw* starts Python without one. |
| Error hook | A program-wide catcher that shows or records any error nobody else handled. |
| Thread | A second line of work running beside the window. |
| Code page (cp1252) | The older Windows way of storing text in a file; it has no room for symbols such as ✓ or ❌. |
| Worker | The part of the program that does a long run (Step 3 or Step 4) in the background, on its own thread. |
| Dry Run | The Step 3 preview that lists what would be done without copying anything. |

### 9.0 Overview

| Group (by the rule above) | Pairs | Came in through the shared commits | Came in separately in A and B |
|---|---|---|---|
| HIGH | 20 (H1-H20) | 18 | 2 (H11 in part, H12) |
| MEDIUM | 40 (M1-M40) | 31 | 9 (M10 in part, M12, M15 in part, M25, M29, M31 in part, M32 in part, M33 in part, M35 in part) |
| LOW | 15 (L1-L15) | 13 | 2 (L3, L7) |
| **Total** | **75** | | |

The 75 rows cover all 73 "same defect in both" rows of MAP §3. The MAP row A-G19 / B-G19, B-G21, B-G22 combines three parts that the audits rated differently, so it is split here into H10, M8 and M9. The fourth part of that row, cleaned names after a run, is MAP's own row A-G19 / B-G75, which appears here as L1. All 75 rows were confirmed in the code of both projects; none had to be dropped.

**What stands out.** The shared business core is where the biggest shared risks sit:

- the consolidated workbook's column look-up (H1);
- the refresh, which cannot be confirmed (H2, H3);
- the Links-sheet handling (H4, H5, M1, M36);
- client identity and file names (H6, H7, H13-H16);
- the take-over of the user's own Excel and keyboard (H19, H20).

All of these were written in the first commit `edac6a3` (31-08-2025). The exception is H1, which became live with the last shared commit `47a1492` (30-10-2025). Neither project fixed any of them after the split.

### 9.1 HIGH

#### H1 · A-G1 / B-G1 · Nine columns of the consolidated workbook show another cell's figure · [GST RULE] ⚠

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Look-up | `A:power_query_extractor/core/data_consolidator.py:141-150` (match on "contains" `:142-145`, exact name only as a fallback `:147-150`) | `B:power_query_extractor/core/data_consolidator.py:129-138` (`:130-133`, `:135-138`) |
| Output titles | `A:power_query_extractor/config/cell_mappings.py:155-228` | `B:power_query_extractor/config/cell_mappings.py:155-228` (the whole file is byte-identical) |
| Values stored by field name and bare cell | `A:power_query_extractor/core/report_processor.py:584-588` | `B:power_query_extractor/core/report_processor.py:535-539` |

- **What happens in both.** For every output column, the consolidator takes the first extracted value whose field name merely *contains* the column title. The output sheets "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" have one-word titles ("IGST", "CGST", "SGST"). These are contained in `GSTR-3B_Difference_Diff_IGST/CGST/SGST`, the first group in the configuration. So those 9 columns show the figures read from `Diff GSTR-3B!J15/K15/L15` ⚠, not from their own configured cells `IMS_Reco!AA1/AD1/AG1`, `ExtraInPurchase!O1/P1/Q1` and `ExtraInGSTR2B!S1/T1/U1` ⚠.
- **Worked example, RAN (C910) in both** (`probe_step4.py`, scenario s1). The real extraction and consolidation code was run with fake COM, after putting a different made-up number into every mapped cell (`J15` = 1001, `K15` = 1002, `L15` = 1003, … `AA1` = 1009 … `U1` = 1017). In both projects the consolidated workbook showed IMS Reco = 1001 / 1002 / 1003, Extra In Purchase = 1001 / 1002 / 1003 and Extra In GSTR2B = 1001 / 1002 / 1003. The other 13 columns showed their own cells' numbers. A and B gave identical workbooks.
- **Came in.** The look-up dates from `edac6a3` (31-08-2025), shared. It did no harm while every title contained a space ("Diff IGST", "Taxable Value"). At `bd176bf` (the commit before `47a1492`) the configuration had only such titles (RAN `git show bd176bf:…cell_mappings.py`). The one-word titles came with `47a1492` (30-10-2025, "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets"), also shared. So the defect has been live in both projects since the last shared commit. Nothing in the message or the code suggests the repetition was intended (MEDIUM: an inference from the message, technical only).
- **What it does to the author's work.** In the annual-return season, the IMS Reco and Extra-in sheets of every client's consolidated row repeat that client's GSTR-3B difference figures. Nothing on screen or in the workbook shows this. Which cell should feed each column is the author's call; the code simply does not do what its own configuration says.

#### H2 · A-G2 / B-G7 · A refresh that failed is later re-used by Skip refresh as a success

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Copy made before Excel opens it | `A:power_query_extractor/core/report_processor.py:164` | `B:power_query_extractor/core/report_processor.py:120` |
| Saved whatever the result | `:349` | `:305` |
| Skip-refresh branch (status "success") | `:133-155` (search through `refresh_naming.py` since `2aaa2b1`, A only) | `:96-111` |

- **What happens in both.** Every refresh first copies the report to its refreshed name. If the refresh fails (Excel cannot open it, an error dialog appears, the `Info!BB2` ⚠ check fails), the copy stays under the refreshed name, already saved. The next run with Skip refresh ticked re-uses it with status success and "Skipped - Using existing refreshed file". Its figures then appear with ✓ in the consolidated workbook. A also shows the copy's date as "last refresh" (A-only screen part). B's Step 4 tree shows "Refreshed" for any such file (B only, `b0e7d0e`).
- **RAN (C910) in both** (scenario s6b, with the files listed in name order, as on a Windows disk). Run 1, with `Info!BB2` empty, gave the status "Refresh validation failed: Cell BB2 is empty", yet `wb.Save` had been called on `ITC_Report_…_Refreshed_….xlsx`. Run 2, with Skip refresh, gave `{'success': True, 'validation_message': 'Skipped - Using existing refreshed file'}`. No key was sent, and the consolidated row read "✓ 2501 2502 2503". A and B were identical.
- **Came in.** `edac6a3`, shared. The Skip option and its re-use branch were already in the first commit (B audit, corrected during its verification). A rewrote the search in `2aaa2b1` (25-07-2026, A only), and the defect stayed.
- **What it does to the author's work.** A report shown as failed in one run becomes "✓" in the next run, with stale or half-refreshed figures, if Skip refresh is ticked.

#### H3 · A-G3 / B-G6 · A refresh cannot be confirmed: key presses, fixed waits, one test cell, saved regardless · [GST RULE] ⚠ (the `Info!BB2` test)

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Keys Alt+A, R, A ⚠ | `A:power_query_extractor/core/report_processor.py:262-292` | `B:power_query_extractor/core/report_processor.py:218-248` |
| Fixed waits | `:296-330` | `:252-286` |
| Save `:349` before the check is used `:354` | as cited | save `:305`, use `:310` |
| The check | `:638-685` | `:589-636` |

- **What happens in both.** Nothing asks Excel whether the refresh has finished. The program presses keys, waits a fixed time, and then treats `Info!BB2` ⚠ as proof if it is not empty, not "0" and not an error text. A value already in the file passes.
- **RAN (C910) in both** (s1, s2; RAN (VERIFY-C910) again, same result). The recorded call order was the same in both: open, `AppActivate` (bring the Excel window to the front), `SendKeys('%a')`, `SendKeys('r')`, `SendKeys('a')` (type Alt+A, R, A into whatever window is in front), `wb.Save`, then the status. With BB2 filled the status was "Validated - Client Name: ClientA"; with BB2 empty it was "failed", but the copy had been saved first.
- **Came in.** `edac6a3`, shared. The wait became a setting in `bd176bf` (30-10-2025), also shared.
- **What it does to the author's work.** A slow query, or keys that landed in another window, is reported as a good refresh whenever BB2 already holds a value. Which test of BB2 proves a refresh is the author's rule; it is only described here.

#### H4 · A-G4 / B-G4 · A template without a `Links` sheet produces no error

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Excel path: another sheet with "link" in its name, else a new empty `Links` sheet is added | `A:core/excel_handler.py:166-183` (added at `:182`) | `B:core/excel_handler.py:135-153` (`:152`) |
| openpyxl path: nothing written, returns True | `:285-306` | `:236-258` |

- **RAN (C910) in both** (`probe_steps123.py`, section 7). On the openpyxl path, a template whose only sheet is "Data" gave `returned: True` and no cell written. The Excel path is READ.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** Every report of a run can be useless (its queries read nothing new) while the log says "✅ Success".

#### H5 · A-G5 / B-G2 · Links cells for file types a client does not have keep whatever the template holds · [GST RULE] ⚠

| | A | B |
|---|---|---|
| Audit rating | MEDIUM-HIGH | MEDIUM-HIGH |
| `if value:` skips empty values | `A:core/excel_handler.py:192` (Excel), `:300` (openpyxl) | `B:core/excel_handler.py:162`, `:249` |
| `''` set for a missing type | `A:core/excel_handler.py:603-705` | `B:core/excel_handler.py:540-622` |

- **RAN (C910) in both** (section 7). The test client had GSTR-3B and GSTR-2B files but no Annual Report and no IMS file. Its ITC report kept the template's own text in `Links!B4`, `B5`, `B10` and `B11` ⚠ (the cells for Annual Report and IMS). `B3`, `B6` and `B9` (ITC) and `B4` and `B7` (Sales) are never written at all (the Links-cell map, `A:utils/constants.py:102-126` = `B:utils/constants.py:96-120`, identical content; B's copy sits six lines higher because B's folder-name block is shorter).
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** If the master template still holds the paths of the client it was built with, a client without an IMS file can get that other client's IMS block in its report. Whether the real templates hold anything in those cells is not known (the templates are not in either repository). Which Links cell each file belongs in is the author's rule.

#### H6 · A-G8 / B-G15 · Two different clients can end up in one client folder · [GST RULE] ⚠ (the folder name carries the state code from the state list)

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Folder name: "Private" → "Pvt", "Limited" → "Ltd", cut to "Max length" | `A:utils/helpers.py:620-644` | `B:utils/helpers.py:513-537` |
| Client key on Step 2: name as written + state code | `A:core/file_parser.py:229-230` | `B:core/file_parser.py:261-262` |
| Folder created from it | `A:core/file_organizer.py:82-99` | `B:core/file_organizer.py:78-95` |

- **RAN (C910) in both** (sections 3 and 6). "Beta Demo Private Limited" and "Beta Demo Pvt Ltd" (Gujarat) became two clients on Step 2 but one folder, `Beta Demo Pvt Ltd-GJ`. Both Sales files were reported "Success". On disk, the copy under the real name held the second file, and the first survived only as `…_backup_211026_1005.xlsx` (Fresh mode). Two invented 41-character names that agree in their first 32 characters became the same folder, `Alpha Test Traders and Distribut-DL`. A and B were identical.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** One client's files and reports can be replaced by another's without any message. A template that reads a whole folder mixes two clients' data.

#### H7 · A-G8 (part) / B-G14 · Two source files of one client whose cleaned names coincide: one survives, which one depends on the mode

| | A | B |
|---|---|---|
| Audit rating | covered inside A-G8 (HIGH) | HIGH |
| Name cleaning (runs of spaces or underscores become one "_") | `A:utils/helpers.py:68-126` | `B:utils/helpers.py:66-112` |
| Existing-file handling | `A:core/file_organizer.py:276-294` | `B:core/file_organizer.py:259-277` |

- **RAN (C910) in both** (`probe_misc.py`, section 3). The source files were `GSTR3B-ClientA-Delhi-Apr 24.xlsx` and `GSTR3B-ClientA-Delhi-Apr_24.xlsx`.
  - Fresh: the copy holds the second file, plus a backup of the first; both are reported "Success".
  - Re-run: only the second file's content is left; both are reported "Success".
  - Resume: the first file is kept and the second is "Skipped".
  - A and B were identical.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** One month's GSTR-3B export can silently go missing from the client folder while the counts say everything was organised.

#### H8 · A-G10 / B-G13 · "Resume" does not resume; a same-minute skip leaves Links cells naming files that are not there

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| New Version folder named after the current minute | `A:core/file_organizer.py:101-107` | `B:core/file_organizer.py:97-103` |
| Skip only if the copy already exists in that folder | `:276-287` | `:259-270` |
| Cleaned name written back only after "Success" | `:233-246` | `:216-229` |
| Mode text "Continue previously interrupted processing" | `A:utils/constants.py:187` (shown under the Resume button via `A:gui/tabs/setup_tab.py:401-411`); A's hover help adds "Continue a run that was interrupted, skipping files that are already in place" (`A:gui/tabs/setup_tab.py:108-109`, `bf2f1dd`, A only) | `B:utils/constants.py:170`, `B:gui/views/setup_view.py:346` |

- **RAN (C910) in both** (section 5). A Fresh run at a fixed 10:00 was followed by a Resume at 10:05. The Resume created `Version-211026 1005` and copied every file again ("Success"). A and B were identical. The wrong-name Links cells after a same-minute Resume are RAN in both audits (A: EXT t12; B: me r2).
- **Came in.** `edac6a3`, shared. B's second, contradictory set of mode texts came with `c258047` / `7e09e5c` (B only).
- **What it does to the author's work.** After an interruption late in the season, Resume repeats the whole batch into a new Version folder instead of continuing.

#### H9 · A-G17 / B-G10, B-G52 · Figures that were read are shown as "Missing", or a ✓ row holds only "Missing" · [GST RULE]

| | A | B |
|---|---|---|
| Audit rating | MEDIUM | MEDIUM-HIGH (B-G10), MEDIUM (B-G52) |
| Status not success → every column "Missing" | `A:power_query_extractor/core/data_consolidator.py:126-129` | `B:power_query_extractor/core/data_consolidator.py:114-117` |
| Extraction sheet names compared with exact capitals | `A:power_query_extractor/core/report_processor.py:564-566` | `B:power_query_extractor/core/report_processor.py:515-517` |

- **RAN (C910) in both; RAN (VERIFY-C910) again, same result.** In s2, `BB2` ⚠ was empty: the row showed "✗" and "Missing" in every ITC column, although all 17 ITC figures had been read (34 entries in the extracted data, because each figure is stored twice, under its field name and under its bare cell, `A:…/report_processor.py:585-588` = `B:…:536-539`). In s3, the template's sheet names were in lower case: the row showed "✓" and "Missing" in every column. A and B were identical.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** A "Missing" in the consolidated workbook does not tell whether the figure is absent, the refresh failed, or a sheet was renamed. The reason is in no sheet. Which figure is shown or replaced is the author's rule.

#### H10 · A-G19 (part) / B-G19 · A client counts as "Success" although both its reports failed

| | A | B |
|---|---|---|
| Audit rating | LOW-MEDIUM (A-G19 as a whole) | HIGH |
| `successful_clients += 1` and status "Success" whatever the reports did | `A:gui/handlers/processing_handler.py:416-429` | `B:gui/controllers/processing_controller.py:135-147` |
| Completion box "Successfully processed p/t clients" | `A:gui/handlers/processing_handler.py:498-502` | `B:gui/pyqt_main_window.py:394-399` |

- **What happens in both.** Only a client whose processing raises an error goes to the "Errors" sheet. All other clients are listed "Success" on the "Client Status" sheet, even when both reports failed. A failed report only adds to `report_errors`.
- **Evidence.** READ in both (the same lines, carried into B's PyQt worker). RAN (B audit, me r8: "Successfully processed 1/1" after both reports failed); RAN (A audit, Section 3 t16 R2: "Success" means "no exception", not "both reports made"; A's Section 7 marks the both-reports-failed case itself READ).
- **Came in.** `edac6a3`, shared; carried into B's new worker by `c258047`.
- **Rating note.** The audits disagree (A: LOW-MEDIUM, B: HIGH) about the same code. It is grouped HIGH by the rule above.
- **What it does to the author's work.** A batch whose reports all failed looks like a full success in the box and in the summary workbook.

#### H11 · A-G22 / B-G34 (in part) · Started without a console, start-up failures are invisible; there is no error hook in either

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Start-up errors only printed to the console, plus one line in `gst_organizer.log` in the start folder (`logging.error`); no window or message box | `A:main.py:114-134` | `B:main.py:49-59` |
| Error hook | none (grep: no `sys.excepthook` outside the tests) | none (grep) |
| A only | dies at once at `A:main.py:17` (`sys.stdout.buffer` when there is no console) | starts, because `B:main.py:11-13` checks first |

- **Evidence.** READ; RAN (A audit, me: `sys.stdout = None` → that error; B audit, me r6a).
- **Came in.** The print-only errors are from `edac6a3` in both. B rewrote `main.py` in `c258047` and kept them, so this came in partly separately. A reworded its start-up messages in `f363946` and `be6418c` (26-07-2026, A only; RAN `git blame` of `A:main.py:114-134`) and also kept them print-only. A's line 17 is from `edac6a3`.
- **What it does to the author's work.** Under the author's Standard 13 launch style, which uses no console, a start-up problem shows nothing at all.

#### H12 · A-G27 / B-G40 · Closing the window during a run kills the work in progress

| | A | B |
|---|---|---|
| Audit rating | HIGH | MEDIUM |
| Code | `A:gui/main_window.py:344-355` (asks once, for Step 3 only, then destroys the window after 1 s); daemon workers `A:gui/handlers/processing_handler.py:129-133`, `A:power_query_extractor/gui/extractor_window.py:878-883` | `B:gui/pyqt_main_window.py:601-608` (`closeEvent` accepts at once, never asks) |

- **Evidence.** READ; RAN (A audit, GUI exp4; B audit, MW e11). A *daemon* thread is killed the moment the program ends.
- **Came in.** Separately: A in `edac6a3`, B in `c258047`.
- **What it does to the author's work.** Half-made output (folders without reports, no summary), and possibly an Excel left running.

#### H13 · A-G33 / B-G17 · A hyphen inside a client or state name shifts the parts of the file name · [GST RULE] ⚠ (file-name patterns)

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Patterns use "anything but a hyphen" for client and state | `A:utils/constants.py:15-55` | `B:utils/constants.py:15-55` (identical except the IMS line, see M38) |

- **RAN (C910) in both** (section 1).
  - `GSTR-2B-Reco-Alpha-Test-Traders-Maharashtra-Q1.xlsx` was read as client "Alpha", state "Test", period "Traders-Maharashtra-Q1".
  - `GSTR3B-Alpha-Test-Delhi-Apr.xlsx` was not recognised.
  - A and B were identical.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** A client whose name contains a hyphen gets a stray client row ("Alpha" in state "Test"), and the real client looks incomplete. Nothing warns. The naming convention is the author's rule.

#### H14 · A-G34 / B-G17 · `Sales-Reco-…` is read as a Sales file of a client called "Reco" · [GST RULE] ⚠

- **Ratings.** A: MEDIUM. B: HIGH (B-G17 as a whole).
- **Code.** `A:utils/constants.py:39` = `B:utils/constants.py:39` (the Sales pattern has four free parts).
- **RAN (C910) in both.** `Sales-Reco-Alpha Test Traders-Maharashtra-Q1.xlsx` was read as type Sales, client "Reco", state "Alpha Test Traders", start month "Maharashtra", end month "Q1".
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** A Sales-reconciliation file typed with one extra hyphen becomes a Sales file of an invented client.

#### H15 · A-G35 / B-G16 · One client spelled in several ways becomes several incomplete clients · [GST RULE] ⚠ (state list)

- **Ratings.** A: MEDIUM. B: HIGH (B-G16 as a whole).
- **Code.** Client key = name as written + state code: `A:core/file_parser.py:229-230`, `B:core/file_parser.py:261-262`.
- **RAN (C910) in both.** "Beta Demo Private Limited" and "Beta Demo Pvt Ltd" became the two keys `…Private Limited-GJ` and `…Pvt Ltd-GJ`. "Tamil Nadu" gives TN but "Tamilnadu" gives TAM, so one client becomes two.
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** Several rows for one company on Step 2, each "Missing n files", although all its returns were supplied. Combined with H6, the Pvt and Private rows then overwrite each other.

#### H16 · A-G36 / B-G16 · State text written in other ways gets invented codes, which can collide · [GST RULE] ⚠ (state list)

- **Ratings.** A: MEDIUM. B: HIGH.
- **Code.** `A:utils/helpers.py:592-618` = `B:utils/helpers.py:485-511`. The table is at `A:utils/constants.py:251-292` = `B:utils/constants.py:234-275` (identical).
- **RAN (C910) in both.** The same results in both:

| State as written | Code | State as written | Code |
|---|---|---|---|
| Tamil Nadu | TN | Tamilnadu / TamilNadu | TAM |
| Delhi / DL | DL | New Delhi | ND |
| NCT of Delhi | NO | Unknown Place | UP (the same as Uttar Pradesh) |
| Dadra and Nagar Haveli | DA | Daman and Diu | DA (the table only has the merged name → DD ⚠) |
| Jammu & Kashmir | J& | 07 | 07 |

- **Came in.** `edac6a3`, shared. The import line was moved in `6237597` (29-10-2025), also shared.
- **What it does to the author's work.** Two spellings of one state make two clients. Two different unknown places with the same letters merge into one. These are letter codes in folder and file names, not the numeric codes of a GSTIN; neither project has GSTIN or PAN logic. Which spellings are acceptable is the author's rule.

#### H17 · A-G44 / B-G51 · Excel set-ups the code does not expect ⚠

- **Ratings.** A: HIGH for (a), MEDIUM for (b). B: MEDIUM.
- **Code.** The keys are `A:power_query_extractor/core/report_processor.py:282-292` = `B:…:238-248`. The English dialog words are `A:…:425-436` = `B:…:376-387`. The button texts are `A:…:506` = `B:…:457`.
- **What happens in both.** Alt, A, R, A are the English ribbon key tips ⚠. Excel in another display language or version may map them differently, and the English dialog words may not match. A locked or busy desktop also stops the keys reaching Excel.
- **Evidence.** READ (cannot run Excel here).
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** On such a PC the refresh may simply not happen, and H3 then reports it as good whenever BB2 already holds a value.

#### H18 · A-G50 / B-G57, B-G39 · Step 4 waits a fixed time per report and nothing times out

- **Ratings.** A: MEDIUM. B: MEDIUM (B-G57), HIGH (B-G39).
- **Code.** The sleeps are at `A:power_query_extractor/core/report_processor.py:185, 253-255, 264-276, 282-292, 297-302, 310-330, 351, 363` and the same lines lower in B (`B:…:141, 207-211, 220-248, 253-286, 307, 319`). Neither has a time limit on any Excel call (grep). Stop: A's Step 4 has no Stop button (`A:power_query_extractor/gui/extractor_window.py:328-340`); B checks Stop only between versions (`B:gui/controllers/extraction_controller.py:60-63`).
- **Evidence.** READ.
- **Came in.** `edac6a3`, shared; the configurable wait came in `bd176bf`, shared.
- **What it does to the author's work.** About 36-39 s of fixed waiting per report at the default 10 s, about 2 hours for 100 clients. A hung Excel or an unrecognised dialog blocks the run for ever.

#### H19 · A-G54 / B-G63 · The program takes over the user's own Excel, hides it and quits it

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH |
| Step 3 attaches to a running Excel first | `A:core/excel_handler.py:123-128` | `B:core/excel_handler.py:110-115` |
| Hides it | `:131-133` | `:118-120` |
| Alerts back on, then `Quit()` always (`self.excel_app` is only ever None, `:56`) | `:238-246` | `:203-211` (`:55`) |
| Step 4 `Dispatch` (the call that starts Excel or connects to one already running) | `A:power_query_extractor/core/report_processor.py:218, 536` | `B:…:174, 487` |
| Same-named workbooks closed unsaved on the second open attempt | `:743-753` | `:691-699` |

- **Evidence.** READ; RAN (B audit, XL E3, fake COM: the user's instance hidden, then `Quit()`).
- **Came in.** `edac6a3`, shared. The comment "Don't quit if Excel was already open" (`A:…:241` / `B:…:206`) says the opposite of what the code does.
- **What it does to the author's work.** Unsaved work in the user's own Excel can be closed or hidden during a run.

#### H20 · A-G55 / B-G61, B-G62 · Key presses and dialog clicks reach whatever program is in front

| | A | B |
|---|---|---|
| Audit rating | HIGH | HIGH, HIGH |
| Keys sent even when Excel could not be brought forward ("continuing anyway") | `A:power_query_extractor/core/report_processor.py:262-292` (`:274`) | `B:…:218-248` (`:230`) |
| Every window on the desktop examined | `:417-453` | `:368-404` |
| First button containing "ok", "close", "yes" or "continue" clicked | `:492-525` | `:443-476` |

- **RAN (C910) in both** (s10). A made-up dialog of another program, titled "OtherApp - unable to sync", with the text "Delete local copy?" and the buttons Yes and No, was found as an "error window". Its "Yes" was clicked (`BM_CLICK`, a click sent by the program instead of the mouse) in both (RAN (VERIFY-C910) again, same result). In s9 the keys were sent although no workbook was open (see L8).
- **Came in.** `edac6a3`, shared.
- **What it does to the author's work.** During a Step 4 run, the PC cannot be used for anything else. Another program's question can be answered "Yes" without the user.

### 9.2 MEDIUM

| # | Gaps (audit ratings A / B) | What goes wrong in both | A (code) | B (code) | Evidence | Came in | What it does to the author's work |
|---|---|---|---|---|---|---|---|
| M1 | A-G6 / B-G3 (MEDIUM / MEDIUM) [GST RULE] | Only the first file of a type (in name order) is written into the Links sheet; the others are copied but not linked. | `core/excel_handler.py:628-630, 639-641, 650-652, 666-668, 682-684, 698-700`; order `utils/helpers.py:361` | `core/excel_handler.py:561-563, 572-574, 583-585, 595-597, 606-608, 617-619`; `utils/helpers.py:262` | RAN (C910) both: two GSTR-2B files `…-H1`, `…-H2` → `Links!B8` ⚠ = `…-H1` only | `edac6a3`, shared | Two half-year 2B reconciliations: the report reads one; nothing says which. Which file should feed the report is the author's rule. |
| M2 | A-G9 / B-G27 (MEDIUM / MEDIUM) | A repeat run in the same clock minute re-uses the same folders and names; backups land inside the linked category folders. | `utils/helpers.py:25-35, 217-234`; `core/file_organizer.py:276-287` | `utils/helpers.py:23-33, 187-204`; `core/file_organizer.py:259-270` | RAN (C910) both: two Fresh runs at 10:00 → one `Annual Statement-211026 1000` / `Version-211026 1000`; `…_backup_211026_1000.xlsx` next to each copy | `edac6a3`, shared | Earlier output silently replaced; a template reading a whole folder may read the backups. B also overwrites its Extraction_Log (B only). |
| M3 | A-G11 / B-G26 (MEDIUM / MEDIUM) | An interrupted copy leaves a partial file under the cleaned name; the report still links that type. | `utils/helpers.py:172-215` (copy `:199`, `except` `:213-215`) | `utils/helpers.py:142-185` (`:169`, `:183-185`) | RAN (C910) both: copy made to fail after 500 bytes → `False`, a 500-byte file left (source 3,604) | `edac6a3`, shared | A dropped network share mid-copy leaves a truncated client file that Power Query may read later. |
| M4 | A-G12 / B-G26, B-G43 (MEDIUM / MEDIUM), in part | Saves go straight to the final name, so a failure half-way leaves a broken file there. Step 3 reports: A builds locally and copies once (`12d505f`, A only; the copy itself can still be cut short); B saves in place. | consolidated `power_query_extractor/core/data_consolidator.py:179`; summary `core/excel_handler.py:402`; report copy `:215` | `data_consolidator.py:167`; `core/excel_handler.py:339`; report `:124` | RAN (C910) both: a time-zone-aware date → `TypeError`, a 2,729-byte `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` left that openpyxl cannot open ("File is not a zip file") | `edac6a3`, shared | A file that looks normal on disk but will not open, possibly replacing a good one from the same minute (M2). |
| M5 | A-G15 / B-G8 (MEDIUM / MEDIUM) | The search `ITC_Report_*.xlsx` also matches refreshed copies; the first file listed is treated as the original. | `power_query_extractor/core/report_processor.py:112-125` | `…:76-89` | RAN (C910) both on this Linux disk: the copy was listed first. **With Skip refresh**, A took that copy as its own refreshed copy (status success, nothing refreshed), while B made and refreshed a copy of the copy (`…_Refreshed_…_Refreshed_….xlsx`). On a Windows disk names usually list in name order, original first (MEDIUM). | `edac6a3`, shared | Growing double-suffix files, or figures from an older copy accepted. |
| M6 | A-G16 / B-G9 (MEDIUM / MEDIUM), in part | Shared: an empty suffix makes the copy target the original; a suffix without `{timestamp}` overwrites the previous copy. (A over-matches custom suffixes, B never recognises them: different.) | `report_processor.py:128-130` | `…:92-94` | RAN (C910) both: empty suffix → "… are the same file"; suffix `_R` (files in name order) → the second run overwrote `…_R.xlsx` | `bd176bf` (30-10-2025), shared | A typo in the "Refreshed File Suffix" box fails every report or overwrites last run's refreshed copies. |
| M7 | A-G18 / B-G11 (MEDIUM for (b) / MEDIUM) [GST RULE] | Values change type on the way into the consolidated workbook: text figures stay text; an Excel error cell probably arrives as a large negative number and is written as a figure. | `report_processor.py:584`; `data_consolidator.py:153-162` | `report_processor.py:535`; `data_consolidator.py:141-150` | RAN (C910) both: "1,234.50" stayed text; -2146826281 written as a number under ✓. Whether real COM returns that number is NOT SURE (check: map a cell holding `=NA()` on Windows). | `edac6a3`, shared | Totals in Excel skip text figures; a "#N/A" can appear as a huge negative amount. |
| M8 | A-G19 (part) / B-G21 (LOW-MEDIUM / MEDIUM) | Each client's `_Organization_Report_…txt` shows running totals of the whole batch. | `core/file_organizer.py:363-368` | `core/file_organizer.py:345-351` | RAN (C910) both: three clients of one file each → "Successfully Copied: 1, 2, 3"; "Folders Created" A 6, 11, 16 and B 6, 12, 18 (A counts the run folder once because it caches it, `275d9a4`, A only) | `edac6a3`, shared | Per-client reports whose numbers do not match the client. |
| M9 | A-G19 (part) / not registered in B's audit (nearest: B-G22) (LOW-MEDIUM / —) | A file whose client part is blank is counted "parsed" and also listed as a problem file. | `core/file_parser.py:210`, `:311` | `core/file_parser.py:243`, `:339` | RAN (C910) both: three files including `AnnualReport- -Delhi-2024.xlsx` → parsed 2, unparsed 1, variations 2 | `edac6a3`, shared | Scan numbers that do not add up. |
| M10 | A-G20 / B-G30, B-G50 (LOW-MEDIUM / MEDIUM, LOW-MEDIUM), in part | Shared: templates are checked for existence only (a folder passes); a target that does not exist is created; file pickers do not open at the chosen folder. (A's Start check, `gui/handlers/file_handler.py:132-151`, tests existence only; A's Setup checklist also tests the extension, `gui/utils/setup_validation.py:56-64`, so there a folder passes only if its name ends like `x.xlsx`.) | `gui/handlers/file_handler.py:132-151`; `gui/utils/setup_validation.py:47-64`; `core/file_organizer.py:65-69`; pickers `gui/handlers/file_handler.py:28-45` | `gui/pyqt_main_window.py:353-361`; `core/file_organizer.py:61-65`; `gui/views/setup_view.py:605-627` | READ both; RAN (A audit, TRACE T7; B audit, me r8) | A `edac6a3` (A's Setup checklist: `322717f`, 25-07-2026, A only); B `c258047` (its own checks): partly separately | A typo in the target creates a new folder tree somewhere else; a wrong template path fails every report later in the run. |
| M11 | A-G23 / B-G33 (MEDIUM / MEDIUM) | The log file is created in the start folder when the program loads, before any message can be shown. | `main.py:20-27` (`:24`) | `main.py:16-23` (`:20`) | RAN (C910) both: a folder named `gst_organizer.log` in the start folder → `IsADirectoryError` before `main()` | `edac6a3`, shared | Started from a protected folder, the program stops with a raw error, or with nothing visible under pythonw. |
| M12 | A-G24 / B-G36 (LOW / MEDIUM) | A settings file with valid JSON of an unexpected shape stops the start (different files, different code). JSON is the text format the settings are stored in. | `gui/utils/cache_manager.py:18-27`; `gui/main_window.py:373-389` | `models/cache_manager.py:57-62`; `models/app_state.py:41-46` | RAN (C910) both: `[1, 2]` → A "'list' object has no attribute 'get'", B "'list' object has no attribute 'items'"; a length "35" written as text → A `TypeError` at once, B accepted it (fails later at the screen, RAN B audit) | Separately: A `edac6a3`, B `c258047` | After a damaged or hand-edited settings file, the program will not open until the file is deleted, and nothing says so. |
| M13 | A-G28 / B-G38 (MEDIUM / MEDIUM) | Step 3's Excel clean-up can itself fail and skip `Quit` and `CoUninitialize` (the call that ends Python's link to COM). | `core/excel_handler.py:236-252` (`:239-240` unprotected) | `core/excel_handler.py:201-217` (`:204-205`) | READ; RAN (A audit CORE exp6, B audit XL E3, fake COM) | `edac6a3`, shared | A stray hidden Excel after a crash; a misleading error line. |
| M14 | A-G29 / B-G43 (MEDIUM / MEDIUM) | One problem while building the consolidated workbook loses the whole Step 4 run's figures. | `data_consolidator.py:20-186` (re-raise `:184-186`) | `data_consolidator.py:20-174` (`:172-174`) | RAN (C910) both: one made-up value with a control character (an invisible, non-printing character) → `IllegalCharacterError`; no workbook at all for a two-client run | `edac6a3`, shared | An hour of refreshing can end with "Processing failed" and the figures only in the logs. |
| M15 | A-G30 / B-G35 (MEDIUM / MEDIUM), in part | Shared: an error inside a button handler is invisible without a console (no error hook). A's click then does nothing; B's whole program aborts (B only). | no `sys.excepthook` (grep) | none (grep) | READ; RAN (A audit me r7; B audit MW e10) | A `edac6a3`; B `c258047`: partly separately | Buttons that "do nothing", with no clue why. |
| M16 | A-G37 / B-G46, B-G22 (MEDIUM / MEDIUM), in part | Shared: files in sub-folders and other formats are ignored without a count; `.xlsm` files are listed but never recognised. | `utils/helpers.py:309-365` (one level, extensions `:327`); patterns `\.xlsx?$` | `utils/helpers.py:249-266`; same patterns | RAN (C910) both: a file in a sub-folder and a `.csv` were not found; the `.xlsm` became "No matching pattern found" | `edac6a3`, shared (A rewrote the scan in `a603c02`, `56aaad9` and `4f84ba9`, 23-25 July 2026, A only; same outcome) | A client folder organised in sub-folders looks incomplete. |
| M17 | A-G38 / B-G47 (LOW-MEDIUM / MEDIUM), in part [GST RULE] ⚠ | Near-miss file names are refused, and the hint offers one guess only. | `core/file_parser.py:288-306` | `core/file_parser.py:316-334` | RAN (C910) both: `GSTR-3B-…` refused, hint "GSTR3B-ClientName-State-Month.xlsx"; "Claims-…" got the IMS hint | `edac6a3`, shared | Files typed slightly differently drop out of the batch; the naming convention is the author's rule. |
| M18 | A-G39 / B-G49, B-G23 (MEDIUM / MEDIUM) [GST RULE] ⚠ | No period or financial-year check; impossible dates accepted; name warnings stored but never shown. | `core/file_parser.py:97-106` (period text), `:110-119` (warnings); nothing reads either (grep) | `core/file_parser.py:93-102`, `:106-115`; same | RAN (C910) both: IMS date "99999999" and GSTR-3B month "13" accepted | `edac6a3`, shared | Last year's and this year's files of one client are grouped without remark (then M1 decides which one is linked). Whether such checks belong in the tool is the author's rule. |
| M19 | A-G40 / B-G18, B-G80 (LOW / MEDIUM, LOW), in part | Shared: text after a dot in a client name is left uncleaned; text before a "/" is dropped; a trailing dot depends on the Python version ⚠. | `utils/helpers.py:76-77, 121-126` | `utils/helpers.py:74-75, 107-112` | RAN (C910) both: "Omega Co. Private Limited" → `Omega_Co. Private Limited`; "Alpha/Beta Traders" → `Beta_Traders`; "Mu Test Traders." → `Mu_Test_Traders.` on Python 3.14 but `Mu_Test_Traders` on 3.11, 3.12 and 3.13 | `edac6a3`, shared | One client gets different report names on PCs with different Pythons. |
| M20 | A-G41 / B-G54 (MEDIUM / MEDIUM) ⚠ | Long paths are not measured before folders and files are made; the ~218-character figure appears only in fallbacks. | `utils/constants.py:81-96`; `core/excel_handler.py:356` | `utils/constants.py:81-89`; `core/excel_handler.py:293` | READ (A's shorter category names, `7347383`, A only, save 9-20 characters per path) | `edac6a3`, shared | Deep client names on the office share fail file by file deep inside a run. |
| M21 | A-G42 / B-G30, B-G66 (MEDIUM / MEDIUM, LOW) | A template or macro workbook is copied byte for byte under an `.xlsx` name. | output names `gui/handlers/processing_handler.py:298, 356`; picker `gui/handlers/file_handler.py:31, 42` (`.xltx`, `.xlsx`) | `gui/controllers/processing_controller.py:220`; picker `gui/views/setup_view.py:612, 619` (`.xlsx`, `.xlsm`) | READ; RAN (A audit XL E1; B audit DO probe_xltx) | The `.xlsx` name: `edac6a3`, shared; B's `.xlsm` offer: `c258047` | Reports that Excel may refuse to open. |
| M22 | A-G43 / B-G53 (MEDIUM / MEDIUM) | Read-only or locked files are not detected; normal opens never pass `IgnoreReadOnlyRecommended` (the Excel option that skips the "open as read-only?" question; only the last attempt does) and never read `ReadOnly` (Excel's flag saying the file was opened read-only). | `core/excel_handler.py:349, 364`; `report_processor.py:781` | `core/excel_handler.py:286, 301`; `report_processor.py:729` | READ (grep) | `edac6a3`, shared | A report a colleague has open fails to save, with the reason only in the log file. |
| M23 | A-G45 / B-G13 (part) (MEDIUM / HIGH for B-G13 as a whole) | Re-run and Resume pick the run folder most recently **modified**, so writing anything into an older batch makes it "the newest". Grouped MEDIUM: B's HIGH rests on the Resume part (H8). | `core/file_organizer.py:174-185` | `core/file_organizer.py:157-168` | RAN (C910) both: a note saved into `Annual Statement-010126 0900` → Re-run chose that older folder | `edac6a3`, shared | Re-run output goes into last month's batch. (Step 4: A uses the same rule; B uses names, B-G12, B only.) |
| M24 | A-G46 / no gap in B's audit (MEDIUM / —) | The "already open" check compares full paths, so a `Z:` drive and the `\\server\share` form of the same file differ; Links paths are absolute. | `utils/excel_com.py:19-29`; `utils/helpers.py:418-422` | `core/excel_handler.py:273-277`; `report_processor.py:656-660`; `utils/helpers.py:311-315` | READ | `edac6a3`, shared (A moved the same comparison into `utils/excel_com.py` in `ba2a753`, 23-07-2026, A only) | Reports built on one PC may refresh on another only if both map the share alike (MEDIUM). |
| M25 | A-G48 / B-G58 (MEDIUM / MEDIUM), in part | Shared: Step 4's walk of the target folder runs on the window's own thread (the window freezes). Only A: the source scan does too. | `power_query_extractor/gui/extractor_window.py:611-668` (button `:240`) | `gui/pyqt_main_window.py:419-504` | READ | Separately: A `edac6a3` (scan code later rewritten in `275d9a4` / `c1583d6`); B `b0e7d0e` | A frozen window on a slow share. |
| M26 | A-G51 / B-G57 (LOW-MEDIUM / MEDIUM) | Step 3 starts or attaches to Excel twice per client and pauses 0.5 s twice per report. | `core/excel_handler.py:104-260` (`:155`, `:210`) | `core/excel_handler.py:95-217` (`:125`, `:180`) | READ | `edac6a3`, shared | Tens of minutes on a few hundred clients. |
| M27 | A-G56 / B-G64 (MEDIUM / MEDIUM) | Client names, Links values and every extracted figure go to a log file in the start folder; on a narrow code page the ✓/❌ lines are lost. | `main.py:24`; `core/excel_handler.py:195`; `report_processor.py:56, 590` | `main.py:20`; `core/excel_handler.py:165`; `report_processor.py:33, 541` | RAN (C910) both, RAN (VERIFY-C910) again: the log appeared in the start folder; with an ASCII locale (a computer setting that allows only plain English letters and digits; a stand-in for cp1252) "✅ Success line" was missing from the file while the plain lines were kept | `edac6a3`, shared (B's run history and Extraction_Log copies: B only) | Under a shared launcher, one shared log collects every user's client figures; the lines that matter for diagnosis vanish. |
| M28 | A-G61 / B-G31 (LOW / LOW-MEDIUM) | All Python warnings are switched off for the whole program. | `core/excel_handler.py:17` | `core/excel_handler.py:17` | READ | `edac6a3`, shared | openpyxl's warnings that it is dropping workbook parts are never seen. |
| M29 | A-G63 / B-G68 (LOW / MEDIUM), in part | Small screens: the window does not fit or loses controls (different code). | minimum 1000×700, `utils/constants.py:136`, `gui/main_window.py:60-61` | minimum 1200×800, `gui/pyqt_main_window.py:72-73` | READ; RAN (A audit START; B audit me r7) | Separately: A `edac6a3`, B `c258047` | On a 1366×768 office laptop parts of B are off-screen; at A's minimum the dark-mode tick and status bar vanish. |
| M30 | A-G65 / B-G71, B-G74 (MEDIUM / MEDIUM), in part | Shared: every Step 4 processor line is logged as "info", errors included. | `report_processor.py:53-60`; no `_log` call passes a level (grep) | `report_processor.py:30-37`; same (grep) | READ | `edac6a3`, shared | Refresh errors are not coloured as errors in the Step 4 log. |
| M31 | A-G66 / B-G72 (MEDIUM / LOW-MEDIUM) | Help texts promise behaviour the code does not have. Shared text: Resume = "Continue previously interrupted processing" (H8). The other texts differ. | `utils/constants.py:187`; `gui/tabs/setup_tab.py:108-109` | `utils/constants.py:170`; `gui/views/setup_view.py:346` | READ | The constant: `edac6a3`, shared; the screens separately | Users choose Resume after an interruption and get a full re-copy. |
| M32 | A-G67 / B-G20, B-G71 (MEDIUM / HIGH, MEDIUM), in part | Shared: unrecognised files reach the screen only as a count. Grouped MEDIUM: B's HIGH rests on its empty Variations sheet (`c258047`, B only). | count `gui/main_window.py:320-321`; list only in the summary workbook | count `gui/pyqt_main_window.py:261-262, 294` | READ | A `edac6a3`; B `c258047`/`b0e7d0e`: partly separately | To learn which client files were left out, the user must open the summary workbook (A) or cannot learn it at all (B). |
| M33 | A-G69 / B-G70, B-G71 (LOW / MEDIUM, MEDIUM), in part | Shared: after Stop, the completion box still says "Successfully processed p/t clients"; neither the box nor the summary workbook says "stopped". Only the on-screen log has one line: A "🛑 Stop requested by user" when the button is pressed (`A:gui/handlers/processing_handler.py:180`), B "Processing stopped by user" when the loop breaks (`B:gui/controllers/processing_controller.py:82`). | `gui/handlers/processing_handler.py:229-230, 498-502` | `gui/controllers/processing_controller.py:80-83`; `gui/pyqt_main_window.py:394-399` | READ; RAN (A audit GUI exp1; B audit MW e5) | A `edac6a3`; B carried into `c258047` (RAN `git blame`) | A stopped batch reads like a finished one. |
| M34 | A-G71 / B-G74, B-G82 (MEDIUM / MEDIUM) | Nothing stops two runs, or two copies of the program, working on the same Excel, folders or files at once; there is no lock or single-copy check anywhere. | `is_processing` per step only, `gui/handlers/processing_handler.py:75`, `power_query_extractor/gui/extractor_window.py:872` | one job per kind only, `gui/controllers/*_controller.py` (`scan:59`, `processing:259`, `extraction:142`) | READ (grep both) | `edac6a3`, shared; B's controllers `c258047`; A's Step 3 re-entrancy guard at `:75` came with `33d0baf` (23-07-2026, A only) | Two colleagues on one target in the same minute overwrite each other; Step 3 can grab Step 4's Excel. |
| M35 | A-G14 (related) / B-G28 (LOW-MEDIUM / MEDIUM), in part | Shared: processing uses the last **scanned** files (a changed source box is not noticed); the completion box names the target as the box reads at the end. | `gui/handlers/processing_handler.py:232, 498-502` | `gui/pyqt_main_window.py:366-371, 381, 397` | READ | A `edac6a3`; B `c258047` (RAN `git blame`) | Messages that name a folder the files did not go to. |
| M36 | A-G5, A-G9 (part) / B-G29 (— / MEDIUM) [GST RULE] ⚠ | The Links cells hand the template folder paths and file names without extensions; GSTR-3B gets only a folder (`B2`). Whatever else sits in those folders (backups M2, duplicates M38, partial copies M3, another client's files H6) may be read. | `core/excel_handler.py:603-705`; map `utils/constants.py:102-126` | `core/excel_handler.py:540-622`; the same map at `utils/constants.py:96-120` | RAN (C910) both: `B2` = the GSTR-3B folder, `B7`/`B8` = folder and stem `…-H1` | `edac6a3`, shared | Depends on the templates' queries (not in either repository). Which form of path they expect is the author's call. |
| M37 | A-G23, A-G56 (part) / B-G44 (MEDIUM / MEDIUM), in part | Shared: the log lands in the start folder and loses ✓/❌ lines on a narrow code page (the same mechanism as M11 and M27; the row is kept because MAP lists it). Only B: no log at all through `launch_extractor.py`. | `main.py:20-27` | `main.py:16-23` | RAN (C910) both (as M11, M27) | `edac6a3`, shared | As M27. |
| M38 | A-G6, A-G38 (part) / B-G48 (— / MEDIUM), in part [GST RULE] ⚠ | Browser duplicate downloads: a GSTR-3B "(1)" becomes another month; a Sales "(1)" becomes a duplicate. IMS "(1)": A accepts it, B refuses it (different). | `utils/constants.py:23, 31` | `utils/constants.py:23, 31` | RAN (C910) both: `GSTR3B-…-Jun (1).xlsx` → month "Jun (1)"; `Sales-…-Mar25 (1).xlsx` → end month "Mar25 (1)"; `ImsReco-…-31032025 (1).xlsx` accepted in A, not recognised in B | `edac6a3`, shared; A's IMS change `5c8ce95` (23-07-2026, A only) | Twelve GSTR-3B months can become thirteen, without remark. |
| M39 | A-G37 (part) / B-G55 (MEDIUM / MEDIUM) | A source folder whose listing fails (share dropped) looks empty instead of failing. | `utils/helpers.py:363-365` | `utils/helpers.py:264-266` | RAN (C910) both, RAN (VERIFY-C910) again: listing made to raise a network error → 0 files, 0 clients, no error raised; the log file gets an "Error finding Excel files in …" line and then "No Excel files found" (READ `A:core/file_parser.py:160`, `B:core/file_parser.py:150`); A's screen says "No valid GST files found" (`A:gui/handlers/file_handler.py:90-91`) | `edac6a3`, shared | "Nothing to do" when the share was simply unreachable. |
| M40 | no gap in A's audit (noted at A-R103) / B-G77 (— / LOW-MEDIUM) | `CoUninitialize` is called while Excel objects are still held, against the author's Standard 12 (`rebuild_pipeline_FINAL.md:872-875`). | `core/excel_handler.py:248-252`; `report_processor.py:410-415, 631-636` | `core/excel_handler.py:213-217`; `report_processor.py:361-366, 582-587` | READ | `edac6a3`, shared | Rarely, an unexplained crash at the end of a report (LOW-MEDIUM, not runnable here). |

### 9.3 LOW

| # | Gaps (audit ratings A / B) | What goes wrong in both | A (code) | B (code) | Evidence | Came in | What it does to the author's work |
|---|---|---|---|---|---|---|---|
| L1 | A-G19 (part) / B-G75 (LOW-MEDIUM for A-G19 as a whole / LOW) | The processing run writes cleaned names back into the scan data, so a Dry Run after a run shows cleaned names. Grouped LOW: A-G19's single LOW-MEDIUM covers all its parts, whose weightier ones have their own rows (H10, M8, M9); by the plain rounding rule this row would be MEDIUM. | `core/file_organizer.py:233-246` | `core/file_organizer.py:216-229` | RAN (C910) both: before `GSTR3B-Client A Private Limited-Delhi-Apr.xlsx`, after `GSTR3B-Client_A_Pvt_Ltd-Delhi-Apr.xlsx` | `edac6a3`, shared | Confusing Dry Run lists after processing. |
| L2 | A-G21 / B-G24 (LOW / MEDIUM), in part | Shared: "Max length" shortens only the client folder (Level 2), never the bracketed names, copied files or report names. Grouped LOW: B's MEDIUM rests on B-only parts (a greyed-out box still applies; migration defects). | `utils/helpers.py:620-644`; `core/file_organizer.py:82-91` | `utils/helpers.py:513-537`; `core/file_organizer.py:78-87` | RAN (C910) both: limit 15 → folder `Alpha Test T-DL`, while names keep `Alpha_Test_Traders_and_Distributors_North` | `edac6a3`, shared | The setting does not shorten the paths that are actually long (M20). |
| L3 | A-G32 / B-G45 (LOW / LOW), in part | The old stand-alone `power_query_extractor/extractor_main.py` cannot be started directly; in B it cannot even be imported (loaded by Python as part of the program). | `power_query_extractor/extractor_main.py:9` (relative import) | `power_query_extractor/extractor_main.py:9` (imports a module deleted in `a62d3cc`) | READ; RAN (A audit DCP §1.1; B audit me r6d) | Separately: A `edac6a3`; B `a62d3cc` | Only if someone starts it that way. |
| L4 | A-G47 / B-G46 (part) (LOW / MEDIUM for B-G46 as a whole) | Upper-case `.XLSX` reports are not found by Step 4 on case-sensitive file systems. A shows the tick but processing fails; B does not list them. Grouped LOW: B-G46's MEDIUM rests on the sub-folder and `.xlsm` parts, which are M16. | `report_processor.py:112-113` | `report_processor.py:76-77`; `gui/pyqt_main_window.py:461-462` | READ; RAN (A audit TRACE T5) | `edac6a3`, shared; B's tree `b0e7d0e` | Only on a Linux-based share; the program itself writes `.xlsx`. |
| L5 | A-G52 / B-G56, B-G59 (LOW / MEDIUM-HIGH, LOW), in part | Shared: log files are appended for ever, never rotated; Step 4 writes about one line per waiting second. Grouped LOW: B's MEDIUM-HIGH rests on its slow HTML (web-page formatted) screen logs (B only). | `main.py:20-27`; `report_processor.py:300-301, 311, 329` | `main.py:16-23`; `report_processor.py:256-258` | READ | `edac6a3`, shared | A log that grows by thousands of lines per client. |
| L6 | A-G58 / B-G65 (LOW / LOW), in part | Text starting with "=" is written into workbooks as a formula. | `core/excel_handler.py:575`; `data_consolidator.py:157` | `core/excel_handler.py:512`; `data_consolidator.py:145` | RAN (C910) both: a variation named `=1+1 notes.xlsx` → a formula cell (type "f") in the summary; an extracted text `=1+1` → a formula in the consolidated workbook | `edac6a3`, shared | Needs an unusual file name; Excel may then show an error. |
| L7 | A-G64 / B-G69 (LOW / LOW) | Dark mode leaves parts light and makes some text unreadable (different code). | `gui/utils/dark_mode_manager.py` | `gui/theme/theme_manager.py` | READ; RAN (A audit START; B audit MW e8) | Separately: A `edac6a3` (rebuilt `aeb83cf`, `725fbaf`, A only); B `c258047`/`7e09e5c` | Readability only. |
| L8 | A-G73 / B-G62 (part) (LOW / HIGH for B-G62 as a whole) | When all three open attempts are rejected, the keys are still sent with no workbook, and the error says "'NoneType' object has no attribute 'Save'". Grouped LOW: the keystroke risk is H20. | `report_processor.py:239-249` | `report_processor.py:195-205` | RAN (C910) both (s9): 9 rejected opens, then `SendKeys` ×3, status "'NoneType' object has no attribute 'Save'" | `edac6a3`, shared | A misleading error; the refreshed copy stays (H2). |
| L9 | A-G74 / B-G79 (LOW / LOW), in part | Code that looks like protection but never acts: the "don't quit" guard, `cleanup()`, the "Open method … failed" branch, `cleanup_on_error`, `validate_results`. | `core/excel_handler.py:56, 241-242`; `report_processor.py:801-804, 167-181`; `core/file_organizer.py:381`; `core/file_parser.py:338` | `core/excel_handler.py:55, 206-207`; `report_processor.py:749-752, 122-137`; `core/file_organizer.py:364`; `core/file_parser.py:366` | READ; grep: no caller of `cleanup_on_error` or `validate_results` in either; `cleanup()` called only at `A:power_query_extractor/gui/extractor_window.py:966` and does nothing | `edac6a3`, shared | A reader believes these cases are handled. |
| L10 | A-G76 / B-G81 (LOW / LOW) | One run is named by two separate clock reads (folders and reports). | `core/file_organizer.py:43`; `gui/handlers/processing_handler.py:186` | `core/file_organizer.py:43`; `gui/controllers/processing_controller.py:44` | READ | `edac6a3`, shared; B's worker `c258047` | A minute change between the two reads gives mismatched names (rare). |
| L11 | A-G77 / B-G16 (part) (LOW / HIGH for B-G16 as a whole) [GST RULE] | The single-word state fallback slices the untrimmed text. Grouped LOW: B-G16's HIGH rests on the parts in H15/H16. | `utils/helpers.py:601` vs `:616` | `utils/helpers.py:494` vs `:509` | RAN (C910) both: " Xyz" → " XY" | `edac6a3`, shared | Needs a caller that passes untrimmed text; the parser trims first. |
| L12 | A-G78 / B-G78 (LOW / LOW) [GST RULE] | Latent: the mapping builder keeps one field name per input cell, so two output columns reading one cell lose the first. | `power_query_extractor/config/cell_mappings.py:255-277` | same file, byte-identical | RAN (C910) both on a modified copy (one extra column reading `Diff GSTR-3B!J15` ⚠): `J15` then maps only to the new field and `GSTR-3B_Difference_Diff_IGST` is gone, so that column would show "Missing" | Builder `edac6a3`; current mappings `47a1492`; shared | Bites only when the author adds mappings. |
| L13 | A-G79 / B-G11 (part) (LOW / MEDIUM for B-G11 as a whole) | Unverified Excel-through-COM assumptions: digits-only text written into a Links cell may become a number; closing workbooks while looping over them may skip some. Grouped LOW: B-G11's MEDIUM rests on the value-type parts, which are M7. | `core/excel_handler.py:193`; `report_processor.py:746-749` | `core/excel_handler.py:163`; `report_processor.py:694-697` | READ (NOT SURE; check on Windows by writing a digits-only stem through Step 3) | `edac6a3`, shared | Unknown until tried on a Windows PC. |
| L14 | A-R60 (part) / B-G42 (— / LOW) | The client look-up and the time estimate sit before the per-client protection, so an error there ends the whole run. | `gui/handlers/processing_handler.py:232-236` (protection starts `:243`) | `gui/controllers/processing_controller.py:85-87` (`:92`) | READ | `edac6a3`, shared; B's worker `c258047` | Unlikely; a whole-run error box when it happens. |
| L15 | no gap in A's audit / B-G76 (— / LOW), in part | Shared: "Skip refresh" travels through a program-wide settings dictionary (a named list of settings held in memory, `EXTRACTION_OPTIONS`), so the tick box and the setting actually used can disagree. Different: in B the value in the configuration file never takes effect, because the box value overwrites it at every Start; in A the file value does take effect, while A's box (which always starts unticked, `power_query_extractor/gui/extractor_window.py:315`) shows "off", until the box is ticked or unticked once. | written only when the tick changes, `power_query_extractor/gui/extractor_window.py:483-487`; read `report_processor.py:134-135`; file value `power_query_extractor/config/cell_mappings.py:249` | written at every Start, `gui/controllers/extraction_controller.py:53-55`; read `report_processor.py:97-98`; file value `power_query_extractor/config/cell_mappings.py:249` | READ both; RAN (VERIFY-C910) A, `skip_probe_A.py` under Xvfb on a copy whose configuration file says `'skip_refresh': True`: box unticked, setting used by Step 4 `True`; after ticking and unticking, both `False` | `edac6a3`, shared; B's per-run write `c258047` | B: editing `skip_refresh` in the file does nothing, although README says it works (`README.md:248`, identical in both). A: a `True` in the file makes Step 4 re-use any earlier refreshed copy (H2) while the box shows unticked. |

### 9.4 How the shared gaps came in, at a glance

| Commit (date, project) | What it brought into both projects | Rows |
|---|---|---|
| `edac6a3` (31-08-2025, shared) | Almost everything: the consolidator look-up, the refresh method and the BB2 test, the Links handling, client keys and name cleaning, file-name patterns, state-code fallback, minute-based names, the processing modes, Excel take-over, the dialog closer, logging set-up, the counting rules | H2-H11, H13-H20; M1-M5, M7-M9, M11, M13-M14, M16-M24, M26-M28, M30, M34, M36-M40; L1-L2, L4-L6, L8-L15 |
| `6237597` (29-10-2025, shared) | The moved import inside `get_state_code` (no change of behaviour) | H16 |
| `bd176bf` (30-10-2025, shared) | The configurable wait and the "Refreshed File Suffix" setting | H3, H18, M6 |
| `47a1492` (30-10-2025, shared) | The output sheets with one-word titles that switch on the H1 mix-up | H1, L12 |
| Separately after the split | The same defect in new code: A's Tk close handler and settings loader (`edac6a3` code kept), B's `c258047`, `7e09e5c`, `b0e7d0e`, `a62d3cc` | H11 (part), H12, M10 (part), M12, M15 (part), M25, M29, M31-M33 (part), M35 (part), L3, L7 |

### 9.5 Checks run for this section

All runs were made in `/root/merge-scratch/run/C910/`, on copies `A/` and `B/` of the two snapshots. The copies were checked blob-for-blob (file content by file content) against `A/main` and `B/main` (RAN `git hash-object`: no difference). The work files are under `tmp*_A`, `tmp*_B`, `startdir_*` and `home_*`.

| Script | What it does | Output |
|---|---|---|
| `probe_step4.py` | Step 4 with fake COM (records calls, never runs Excel): consolidated look-up (s1), BB2 empty (s2), sheet-name case (s3), odd values (s4), one bad value (s5), Skip refresh on this disk (s6) and in name order (s6b), report search (s7), suffix cases (s8), rejected opens (s9), another program's dialog (s10) | `out_step4_A.json`, `out_step4_B.json` (identical apart from paths, times, A's `WIN32_AVAILABLE` flag (a marker A sets to say whether the Windows add-on pywin32 loaded; B has none) and the listing-order cases s6 and s7, explained in M5) |
| `probe_steps123.py` | File names, state codes, folder names, name cleaning, scan with sub-folder / `.xlsm` / listing failure, Fresh and Resume, two spellings of one company, Step 3 reports on the openpyxl path, the summary's Variations sheet, organisation-report totals, name write-back | `out123_A.json`, `out123_B.json` (differences only in the work-folder paths, category folder names, the IMS "(1)" pattern and the "Folders Created" totals, all explained above) |
| `probe_misc.py` | Partial copy, Re-run's choice of run folder, cleaned-name collision in three modes, the mapping builder on a modified copy, a save that fails half-way | `outmisc_A.json`, `outmisc_B.json` (identical) |
| `scripts/start_probe.py` | Loads `main.py` from a start folder holding a folder named `gst_organizer.log`, and from one with an ASCII locale | console output (above) |
| `scripts/settings_probe_A.py`, `settings_probe_B.py` | Settings files of an unexpected shape | console output (above) |
| `scripts/blank_client.py` | Blank client part in a file name | console output (above) |
| `scripts/blame.sh` | `git blame` of the cited lines on `A/main` and `B/main`, printed as commit, date and subject | console output (above) |

**Re-runs while checking this section (RAN (VERIFY-C910), 23-10-2026).** In `/root/merge-scratch/run/VERIFY-C910/`, on fresh copies `A/` and `B/` of `A-copy` and `B-copy`, the three probes and the four small scripts above were run again with `venv-A` and `venv-B`. Every result quoted in 9.1-9.3 came out the same (`out_step4_*.json`, `out123_*.json`, `outmisc_*.json`; `outmisc_A.json` and `outmisc_B.json` byte-identical). Two further runs were added:

| Script | What it does | Output |
|---|---|---|
| `skip_probe_A.py` | Builds A's Step 4 panel under Xvfb (a screen-less display) from a copy `A_skip/` whose configuration file says `'skip_refresh': True`, and compares the tick box with the setting Step 4 reads (L15) | console: "tick shows: False \| setting used by Step 4: True"; after ticking and unticking "False False" |
| `san.py` | A's and B's name cleaning under Python 3.11, 3.12, 3.13 and 3.14 (M19) | console: `Mu_Test_Traders` on 3.11-3.13, `Mu_Test_Traders.` on 3.14, in both |

## 10. Data and confidentiality

This section lists, by location only, client data, personal data and anything secret-looking that the two repositories hold now or held at any point in their histories. No content was opened or copied. Where a file had to be measured, only its size, line count, a checksum or the number of pattern matches was taken. Text that could not be quoted was located by a search that printed only line numbers, hashes or a masked shape (every letter shown as x, every digit as 9). The section states facts only; what to do about them belongs to the decision list.

**Words used.** A *repository* is a project's folder together with its full change history. *History* means the earlier commits: a file deleted in a later commit stays readable in the earlier ones. *Commit metadata* is the name and e-mail address recorded with each commit. A *public* repository on GitHub can be read and downloaded by anyone, without signing in, including every earlier commit. A *fork* is someone else's copy of a repository. A `.pyc` file is the compiled form Python saves next to a source file when it runs it. `.gitignore` is the list of files git is told not to track.

### 10.1 Who can see the two repositories

| Fact | A (`AnnualStatement_oldv1`) | B (`AnnualStatement_oldv2`) | Evidence |
|---|---|---|---|
| Visibility | public | public | Recorded as public on 21-10-2026 in the brief for this comparison; re-checked on 23-10-2026: `"visibility": "public"`, `"private": false` for both (RAN (C910) `gh api repos/goyal-sidhant/<name>`). GitHub keeps no record of earlier visibility that this session can read, so the date 21-10-2026 itself rests on the brief; the B audit's own check (RAN, 23-10-2026) found both public too. |
| Branches on GitHub | `main` only | `main` only | RAN (C910) `gh api …/branches` |
| Public tip = audited snapshot | `main` = `bf2f1dd` | `main` = `b0e7d0e` | RAN (C910) `git ls-remote` |
| Forks | 0 | 0 | RAN (C910) `forks_count` |
| Created | 07-02-2026 05:51:52 UTC | 31-08-2025 16:38:53 UTC | RAN (C910) `created_at` |

Because both are public and their tips are the audited snapshots, every item in 10.2 that is "in history" can be downloaded from GitHub today, as well as every item "in the current files".

### 10.2 The locations

| # | What (kind of data, never the content) | Where | Repository | Commits that carry it | Current files or history only | How checked |
|---|---|---|---|---|---|---|
| D1 | **`gst_organizer.log`**: the program's own run log from real use, 803,193 bytes, 5,689 lines. Judging by the logging statements it matches, it holds client names, state names, source file names, full folder paths and the values written into the Links cells (HIGH by the code; the file was **not opened**). | repository root | **both** | `edac6a3` (31-08-2025) and `6237597` (29-10-2025), the same git object in both (`b1ca290be3…`). Deleted by `cbfa4a5` (31-08-2025) on one line of history; `6237597` was made on a copy that still had it; the merge `9c56a1c` (29-10-2025) dropped it. All shared, before the split. | **History only** (absent from every commit from `9c56a1c` on, in both) | RAN (C910): `git ls-tree` of every commit of both histories; `git rev-parse` (same object); `git cat-file -s`; line count. Both audits' X5 (A item 1, B CD1). |
| D2 | **70 compiled `.pyc` files** in 11 `__pycache__` folders, 35 made by Python 3.12 and 35 by Python 3.13. Each one embeds the absolute Windows path, with a drive letter, of its source file on the computer that compiled it (70 of 70). None contains a `\Users\` segment. The folder names inside the paths were not examined; they could include a personal or client name (LOW-MEDIUM). Two of them are the only trace of an older single-file `gui/app.py` (A audit). | `**/__pycache__/*.pyc` | **both** | `edac6a3`, `cbfa4a5`, `6237597`, `9c56a1c`, `d674386`, `5ea77c2`, `1cb878b` (31-08-2025 … 29-10-2025): the identical set in all seven (RAN: one checksum of the 70 object ids for all seven). Removed by `01862a0` (29-10-2025). All shared. | **History only** | RAN (C910): counts per commit in both histories; a byte search that printed only counts. Both audits' X5 (A item 2, B CD2). |
| D3 | **A company-like business name** that may be a real company. An AI-written document uses it in its worked examples, as the client part of example file names, folder paths and client keys. Possibly an illustrative choice by the AI, possibly a real client; only the author knows. | `PROJECT_AUDIT_ANNUALSTATEMENT.md` lines 130, 133, 134, 158, 160, 161, 162, 177, 179, 180, 182, 227, 228, 231, 232 (line 231 is an example output path built with it) | **A only** | Introduced by `f7eb8f1` (05-03-2026); present in that file in 33 commits (`f7eb8f1` and A's 32 July-2026 commits); absent from the first version `fc279af` (04-03-2026); in no commit message of A or B; in no file of B's history | **Current files** and history | RAN (C910): the name was located by a search that compared word positions and printed only a hash and line numbers; counts per commit. A audit X5 item 3; B audit CD8 (never in B). |
| D4 | **The author's own local project-folder path**: a drive letter, a cloud-sync folder and a year folder, 4 folder levels. A personal environment detail, not client data. | `PROJECT_AUDIT_ANNUALSTATEMENT.md` line 5 (header) and line 663 | **A only** | Line 5 since `fc279af` (04-03-2026); line 663 since `f7eb8f1` (05-03-2026) | **Current files** and history | RAN (C910): a search for drive-letter paths that printed only line numbers and a masked shape. A audit X5 item 4. |
| D5 | **The office's network share path** (a server name with digits, share folders, and a folder named after a financial year), plus figures from real runs (number of files scanned, scan times, path lengths). Infrastructure and operational detail, no client names. | Commit messages of `4f84ba9` (25-07-2026; the share path in its first paragraph) and `7347383` (26-07-2026; in the paragraph on measured path lengths) | **A only** | Those two commits, both A's own (after the split). No other commit message in either history holds a network or drive path. A search hit in `33d0baf` was the text of a word-boundary pattern, not a path. | **History only** (commit metadata is not a file; it cannot be changed without rewriting history) | RAN (C910): a search of every commit message that printed only commit ids and masked shapes. A audit X5 item 5. |
| D6 | **Commit metadata: the author's identity.** A: 2 distinct author e-mail addresses (a personal address on 45 commits, a GitHub "noreply" address on 1, `cbfa4a5`); author names "Sidhant Goyal" (32 commits, A's July-2026 work), "goyal-sidhant" (13), "Sidhant" (1, `47a1492`). B: 2 distinct author e-mail addresses (16 + 1, the same `cbfa4a5`); names "goyal-sidhant" (16), "Sidhant" (1). A "Co-Authored-By" line naming an AI model is on 32 A commits and 3 B commits (`c258047`, `a62d3cc`, `c21f788`). Personal data of the author, not client data. The addresses are not reproduced here. | Every commit | **both** (the 12 shared commits carry identical metadata in both) | All 46 commits of A, all 17 of B | History only (part of every commit) | RAN (C910): `git log --format=%ae / %ce / %an`, counted. Both audits' X5 (A item 6, B CD7). |
| D7 | **`.claude/settings.local.json`**, 9 lines: a setting of the AI coding tool that lets it run `python -m py_compile` without asking. No secret, no path, no client data. | `.claude/settings.local.json` | A: current; B: history only | Added by `bd176bf` (30-10-2025, shared). Still tracked in A at `bf2f1dd`. Deleted in B by `c258047` (07-02-2026), and B's `.gitignore:82` ignores `.claude` since then. | A: **current files**; B: history only | READ (C910) in full; RAN `git log -- .claude/settings.local.json`. Both audits' X5 (A item 7, B CD5). |
| D8 | An ignore rule `/mnt/user-data/` (the name of a web AI sandbox's folder, MEDIUM) in the block headed "User data and sensitive information". Only a rule; nothing in either project reads or writes that folder. | `.gitignore:62` (A), `.gitignore:61` (B) | both | Since `1cb878b` (29-10-2025, shared) | Current files | READ. Both audits' X5 (A item 8, B CD6). |
| D9 | Generic examples, listed only so the author can confirm they are not real clients: a company-style example name (three capital letters followed by "Ltd") used six times with a state name in README's file-naming table; the generic install folder `C:\GST_Organizer`. Both audits judged them obvious placeholders. | `README.md:128-133` and `README.md:70` (byte-identical README in both) | both | `README.md:128-133` since `d674386` (29-10-2025, shared); the README at `edac6a3` has no such phrase (B audit) | Current files and history | READ; RAN (B audit, count). A audit X5.2; B audit CD3, CD4. |
| D10 | **Nothing else found.** In every distinct file version of both histories (A: 146 versions; B: 102 versions; `gst_organizer.log` and the `.pyc` files left out), there were 0 GSTIN-shaped strings, 0 PAN-shaped strings, 0 e-mail-shaped strings and 0 private-key headers. The one "secret word = value" hit is `TIMESTAMP_TOKEN = "{timestamp}"` in A's `power_query_extractor/core/refresh_naming.py:15`, a constant. No Excel workbook, CSV or JSON data file was ever committed to either repository. | all files, all commits | both | — | — | RAN (C910): a pattern count over every file version, printing counts and file names only. Both audits' X5.0 (same result). |

### 10.3 Where the running programs write client data (outside the repositories)

Both audits' Appendix X5.3 list these places. They are not repository content, but one of them is how D1 entered the history. Facts from READ and the cited audit runs:

| File | A | B | Client data in it | Can it reach the repository? |
|---|---|---|---|---|
| `gst_organizer.log` | the folder the program was started from, `A:main.py:24` | the same, `B:main.py:20` | client names, paths, every Links value (first 50 characters, `A:core/excel_handler.py:195` / `B:core/excel_handler.py:165`), every extracted figure (`A:power_query_extractor/core/report_processor.py:590` / `B:…:541`) | Yes, if the program is started from its own folder. That is how D1 was committed, before `.gitignore` existed (`1cb878b`, 29-10-2025). Since then `.gitignore` excludes `*.log` and `gst_organizer.log` (A `:40-41`, B `:39-40`). |
| `pq_extractor.log` | the start folder of the stand-alone extractor, `A:power_query_extractor/extractor_main.py:12-19` | never written (that entry point cannot be imported, L3) | client folder names, paths, figures | Ignored by `.gitignore` (A `:42`, B `:41`). |
| Settings in the user's home folder | `~/.gst_organizer_cache.json`: the four paths and the last five source folders (`A:gui/utils/cache_manager.py:15, 32-43`) | `~/.gst_organizer/settings.json`, `session_state.json`, and `run_history.json`, which holds client names and the whole extraction result (every extracted figure) of each client and Version folder; a repeat extraction of the same Version folder replaces the earlier figures (`B:models/run_history.py:80-85`), but the per-client records are never trimmed, only the list of runs is cut to 50 (`:69-71`), so they are kept indefinitely (`B:models/run_history.py:73-95`, key from `B:gui/pyqt_main_window.py:554-557`) | paths (A); paths, client names and figures (B) | No (outside the project folder). |
| Output in the chosen target folder | client-file copies and backups, ITC and Sales reports, refreshed copies, organisation reports, summary and consolidated workbooks | the same, plus `Extraction_Log_<…>.txt` | all of it | Only if the target is inside the repository folder. `.gitignore` excludes `/Annual Statement-*/` at the root only (A `:57`, B `:56`). No rule covers ordinary `.xlsx` files: only `~$*.xlsx` and `*.test.xlsx` are excluded (A `:50, 69`, B `:49, 68`). |
| `%TEMP%\gst_excel_temp` | a staging copy of each report while it is built, deleted afterwards on a best-effort basis (`A:core/excel_handler.py:147-154, 254-260`, `12d505f`, A only) | created at each run, left empty (`B:core/excel_handler.py:52-53`) | A: the report while it is built; B: none | No. |

---

## Review record

This record belongs to the final step of Phase 2 (task C12-AND-ASSEMBLE, 23-10-2026): Sections 1 and 2 were written, the document was assembled, and the whole of it was then checked for consistency and coverage. The review read every section from start to end. It compared them with each other and with the two final audits, including audit B's final text, which was corrected after Sections 3 to 10 were written. Every correction was made in place and marked "[Corrected during review - …]"; all of them are listed in the second table below. Nothing was shortened or removed. All scripts are in `/root/merge-scratch/run/C12-AND-ASSEMBLE/` (scripts in `scripts/`), and all were run with `python -B` (RAN). They only read the two code copies, the audits and the history; the only files written are the two output files (`/root/merge-scratch/compare/C12.md` and this document) and working files in that folder.

### What was checked, and how

| Check | How | Result |
|---|---|---|
| Assembly | `assemble.py` joined the title block (`title_block.md`) with `C12.md`, `C3.md`, `C4.md`, `C5a.md`, `C5b.md`, `C6.md`, `C78.md` and `C910.md` from `/root/merge-scratch/compare/`, in that order. `sync_c12.py` later replaced Sections 1-2 with the final `C12.md` without touching the review edits. `check_headings.py` skips code blocks; it was run after every change. | One level-1 title. The headings "## 1." to "## 10." each appear once and in order; the only other level-2 heading is this record. Code fences are balanced (one diagram). |
| No contradictions between sections | Every section was read in full. Recurring facts were compared across sections: commit, file, test, error-handler, message-box and gap counts; the 39-second fixed wait per report; the IMS Reco, folder-name, Step 4 batch and Links-cell examples; the relation words and the verdicts. Each section's own list of corrections was compared with the other sections. | Four inconsistencies, all fixed (corrections 4, 7, 13 and 14: when B rebuilds its Step 4 list). The other differences in figures turned out to measure different things; they are explained under "Differences that are not contradictions" below. |
| Against audit B's final text | Audit B's final file was compared line by line with the section files the writers had read. `assemble_b_sections.py` joined `audit/B/S1.md` … `X345.md` the way audit B was assembled; `diff` of that join against the final audit gave `B_final_vs_sections.diff`, 116 changed places; `worddiff.py` turned them into `B_final_changes.txt`. Every changed fact was then searched for in this document. | The final audit's file is byte-identical in `/home/user/…/merge-analysis/` and `/root/merge-scratch/audit/` (RAN `sha256sum`). Six changed facts needed an edit here: the tab dots (correction 6), B's start-up import chain (9, 10), the half-written consolidated workbook (5, 11), the settings-import marker (12), the drawn shapes (15; audit A's completeness check found the same) and the line count (18). The rest are listed under "Checked and found consistent". |
| Against audit A's final text | Audit A was final (22-10-2026) before Sections 3-10 were written, so the review spot-checked the facts Section 1 quotes from it (counts, test results, states, phases) and the places where a section corrects it. | No contradiction in this document. Two places in audit A are wrong (see "Corrections to the audits" below). |
| Feature and rule coverage | `scripts/idcheck_doc.py` searched each section of the assembled document for IDs named one by one; a range such as "A-R80 to A-R82" was not counted. | Section 3 names all of `A-F1`-`A-F31` and `B-F1`-`B-F34`. Section 5 names all of `A-R1`-`A-R157` and `B-R1`-`B-R144`. The whole document also cites every gap `A-G1`-`A-G79` and `B-G1`-`B-G82`. No ID is out of range. |
| GST-rule neutrality | `scripts/neutral.py` flagged every line that holds both a GST-rule mark and a judgement word (better, right, wrong, should, safer, closer …). Every hit was then read, together with every "which works better" verdict in Sections 3, 4 and 8. | Four verdicts or labels needed a change (corrections 2, 3, 8 and 17), and the "At a glance" note of Section 3 was brought into line (correction 1). The other hits describe what the code requires, quote the program, or carry "no judgement" / "the author's call" already. Section 1 names GST-rule differences without a verdict. |
| Confidentiality | `scripts/confcheck.py` and `scripts/namecheck.py` printed counts, hashes and masked shapes only, never the text found. Searched for: GSTIN-shaped and PAN-shaped strings; e-mail addresses; `Users\<name>` paths; the client part of the example file names in lines 130-232 of A's `PROJECT_AUDIT_ANNUALSTATEMENT.md`; the server-like part of the office share path in the commit messages of `4f84ba9` and `7347383`. | 0 GSTIN-shaped, 0 PAN-shaped and 0 e-mail-shaped strings. 0 hits for the company-like name and for the server name. The only drive paths are placeholders (`C:\Users\<name>…`) and the README's example `C:\GST_Organizer`. `gst_organizer.log` is described by kind only. The author's own name appears only in the title block (whom the document is for) and in Section 10, D6, which describes the commit metadata; the GitHub account name `goyal-sidhant` appears as part of the repository names. |
| Plain English | Section 1 explains every technical term it uses, in a word table or in brackets, and Section 2 relies on that table; Sections 3, 4, 6, 7-8, 9 and 10 begin with their own word lists, and Section 5 explains its terms in brackets at first use. The review added brackets in Sections 1-2 where a term had first appeared unexplained (`requirements.txt`, pip, `git merge-base`, `--shortstat`, shebang, `.gitignore`, compiled files, late binding, Excel signatures, bare `except:`, `~$` lock files, hover help). | No unexplained term was found in Sections 1-2 after these additions. Sections 3-10 were read for plain English but not re-checked term by term (see "Limits"). |
| Code and history checked by this review | READ: `B:gui/pyqt_main_window.py:17-27, 430-437, 545-570`; `B:gui/controllers/__init__.py:1-4`; `B:models/cache_manager.py:93-142`; `B:power_query_extractor/core/report_processor.py:11-15, 141, 209-320`; `B:core/excel_handler.py:109-126, 205-210`; `B:utils/constants.py:23, 85-89`; `B:requirements.txt:13-14`; `B:main.py:11-13`; and the A counterparts `A:power_query_extractor/core/report_processor.py:20-24, 282-292`; `A:core/excel_handler.py:122-128, 147-155, 240-245`; `A:utils/constants.py:23, 91-95`; `A:power_query_extractor/gui/extractor_window.py:629-633`; `A:power_query_extractor/core/data_consolidator.py:141-150` (and B `:129-138`); `A:gui/handlers/file_handler.py:123, 135, 139`; `A:gui/handlers/processing_handler.py:45, 83, 224`; `B:gui/controllers/processing_controller.py:73`; `A:main.py:17`. RAN (read-only): `git rev-list --count`, `git merge-base`, `git ls-tree`, `git log` (dates, parents, AI "Co-Authored-By" lines, "Reported" messages), `git diff --shortstat`, and `git ls-remote` of both GitHub repositories (`main` still at `bf2f1dd` and `b0e7d0e` on 23-10-2026); file counts of Tkinter and PyQt5 importers (17 and 14). | Every cited line said what this document says it does, after the corrections below. |

### Corrections made in the text

| # | Where | What was wrong or missing | What it says now, and the evidence |
|---|---|---|---|
| 1 | Section 3, "At a glance" | Named only the nine rows that hold the MAP 2.1 GST-rule differences, and called the GST content of row 6 identical. | Also names rows 6 and 22, whose parts Sections 4 and 5 tag as GST rules "when in doubt". It says that in row 6 only the rule that builds the folder name is identical, while the limit each box lets through differs. Evidence: Section 4, 4.4.3; Section 5, 6.1 and 9.4. |
| 2 | Section 3, row 6 (A-F6 / B-F4) | The verdict "B, for input safety" stood without the GST-rule label, although Sections 4 and 5 treat the length limit as a GST rule through B-R32. | Starts "GST rule - no judgement" on the shortest allowed length. The technical verdict (B) is kept, and it now says it concerns only A's acceptance of text and of values outside its own range. |
| 3 | Section 3, row 22 (A-F22 / B-F21) | The verdict "A" stood without the GST-rule label, although Section 5, 9.4 tags Skip refresh as a GST rule. | Starts "GST rule - no judgement" on which earlier copy Skip refresh may reuse. The technical verdict (A: it follows its own suffix setting; B ignores its own, B-G9) is unchanged. |
| 4 | Section 3, row 18 | "B walks the folders … after extraction." | "after an extraction that ends normally or is stopped, not after one that ends in an error" (READ `B:gui/pyqt_main_window.py:549-565`; Section 4, 4.7 item 2). |
| 5 | Section 3, row 24 | "On a save error, both lose the workbook" only. | Adds that an error raised while the file is already being written leaves a half-written, unopenable workbook under the normal name, in both. Evidence: Section 9, M4 (RAN both); audit B's B-R86 and B-G43, as corrected in its self-review (corrections 24-25, 55). |
| 6 | Section 3, 3.2, row 2 (tab dots) | "Section 7 has no gap number of its own for them." | The four tab-dot defects are now listed under B-G71 in audit B's final text (its self-review, corrections 11-19). |
| 7 | Section 3, 3.2, row 18 | B rebuilds its Step 4 list "after extraction". | Narrowed as in correction 4. |
| 8 | Section 4, row 4A-26 | The consequence read "Technical." only. | Adds that which earlier copy Skip Refresh may reuse is a GST rule for the author (Section 5, 9.4), and that the technical point is B ignoring its own suffix setting. |
| 9 | Section 4, row 4A-32 | B's start-up chain was given as `B:gui/pyqt_main_window.py:27` → the extraction controller. | Line 25 comes first: it loads the package file `B:gui/controllers/__init__.py`, whose line 4 loads the extraction controller (line 27 would load it too). READ both files; audit B's self-review, correction 29 (B-R124). |
| 10 | Section 5, 9.2 | The same `:27` statement. | The same correction as 9. |
| 11 | Section 5, 11.1, row "File name, folder, when written, errors" | Did not say that a failed save can leave a broken file. | Adds the half-written, unopenable workbook in both projects. Evidence as in correction 5. |
| 12 | Section 5, 14.4 | The settings import "runs once (marker `.migrated`)", with no caveats. | Adds that the marker is written even when `settings.json` could not be saved, and that deleting the marker makes the next start import the old settings again. READ `B:models/cache_manager.py:100-101, 119, 127, 135-142`; RAN in audit B (B-F2 bugs 3 and 5, B-G24; its self-review, corrections 42-43). |
| 13 | Section 5, 16.13 | B fills Step 4 "after extraction". | Narrowed as in correction 4. |
| 14 | Section 6, 6.8.1, row U3 | "after each extraction". | Narrowed as in correction 4. |
| 15 | Section 7, 7.3 "Power Query protection" | Did not mention that the openpyxl fallback also drops drawn shapes. | Adds that both audits' completeness checks RAN this with a made-up template holding one text box (A audit Appendix X1.4.1 and A-G7; B audit Appendix X1.4.3 F5 and B-G32). It also notes that Standard 12's table names shapes (`Std:827`), that a `DataMashup` check would not detect this loss, and that whether the real templates hold shapes is NOT SURE. Section 1, 1.4 (c) item 2 says the same. |
| 16 | Section 8.1, "What a test would have caught in B" | "a 28-minute window in which the program could not start". | 27 minutes: `a62d3cc` 11:29:56 → `c21f788` 11:57:02 on 07-02-2026 (RAN `git log --date=iso`). Audit B's Sections 1, 2 and 12A also say 27; its X3.5 says 28 (listed below under "Corrections to the audits"). |
| 17 | Section 8.4, silent-logic table, row title | "Wrong cell or column wiring [GST RULE]": the word "wrong" judges GST-rule items (the nine columns, the first-file rule, the Links cells for missing types). | "Cell or column wiring that differs from the program's own configuration, or that a template may not expect", marked as a technical description only, with "which cell or file is right is the author's call". This follows the rewording both audits' completeness checks made for the same items. |
| 18 | Section 8.7, item 3 | Said that audit B's Section 9.9 gives 7,805 lines. | Adds that audit B's self-review (corrections 37-38) changed Sections 8.2.1 and 9.9 to 7,821, so the final audit and this document now agree. |

### Checked against audit B's final text and found consistent (no edit needed)

- **Two source files of one client whose cleaned names coincide** (audit B's B-F12 correction: in Resume the second file is "Skipped" and the first is kept). Section 9, H7 already gives all three modes this way.
- **Name cleaning on different Pythons** (B-R20, settled HIGH). Section 7.1 and Section 9, M19 already give the 3.14 against 3.11-3.13 result, RAN in both projects.
- **Closing during an extraction** (B-G40 extended): exit code 0, an unrefreshed copy left, and no workbook, log or history. Section 3, rows 2 and 25 and 3.2 row 2 already say so.
- **A processing run ending during an extraction switches the Step 4 list to the new batch** (B-G74). Section 3, 3.2 row 34 and Section 5, 11.2 already say so.
- **A damaged run history aborts the program at the end of processing** (B-R119). Section 3, row 33 already says so.
- **File pickers** (B-G50 narrowed to "first dialog of a session"). Sections 6.4.7 and 9, M10 say only that the pickers do not open at the chosen folder, which stays true.
- **The Qt version** (5.15.19 running, 5.15.14 built against). Section 6.0 already says 5.15.19.
- **Python 3.7** (pip can still install builds for 3.7; running on it is untested). Section 7.1 says the 3.7 floor was never tested and does not claim it cannot be installed.
- **The openpyxl fallback at the split point** (a Windows PC without pywin32 then lost the queries). Section 7.1 already says so.
- **A copy left even when Excel never starts** (B-G7). Section 9, H2 says the copy is made before Excel opens it; consistent.
- **Step 4 options not saved at Start Extraction** (B-G28 extended). Sections 5, 14.1 and 6, D9 list the save moments without Start; consistent.
- **The dialog text reaches the screen only when the `Info!BB2` check passed** (B-F26). No section claims otherwise; Section 6, D56 says that for many failures the reason goes only to the log file.
- **`AppActivate(excel.Caption)`, the README range 221-253 and the second local branch name.** Either not mentioned here, or already right (Section 2).
- **Which repository is the original.** Sections 4.1, 7, 10.1 and 1.2 all follow GitHub's creation dates, as audit B does.

### Differences that are not contradictions

- **How many GST-rule differences there are.** The map counts eleven items (MAP 2.1; its item 8 covers two rule pairs). Section 5's categories 1 to 8 count 14 pairs, because Section 5 also tags A-R36 and A-R52 through their B partners, which B's registry tags. Section 5, 9.4 tags Skip refresh, and Section 4 tags rows 4A-12, 4B-9, 4B-10 and 4B-20. All of these follow the rule "when in doubt, treat it as a GST rule". Section 1, 1.4 (a) lists every one and marks the extra ones with *.
- **The fixed wait per report.** Sections 3, 5 and 7 give 39 seconds at the default setting. Section 9, H18 says "about 36-39 s", because its 36 seconds are the waits inside the refresh session and the remaining 3 seconds come before the figures are read (36 + 3 = 39; READ `B:power_query_extractor/core/report_processor.py:141, 209-319`, the same waits as A). The waits also shrink by up to 5 seconds when the dialog check finds a dialog (`:266-275`). Section 7.2, rule 20 gives 33-39 seconds for the time the desktop must not be touched. All three describe the same code.
- **"Only B" gaps.** The map pairs 9 gaps as "only B", while audit B counts 26 gaps that "come only from B's own commits". Audit B's count is about where a defect came from, and A can have the same defect through code of its own (Section 9 marks those "separately"). The two numbers answer different questions.
- **Registry tags.** The two registries tag some matching pairs differently (for example A-R48 against B-R57). Sections 4, 5 and 9 tag a pair whenever either registry does, so a "[GST RULE]" here does not always appear in both audits.

### Corrections to the audits found during this review

1. **Audit A, Section 1 (its line 7) and Section 12A (its line 10582)** say "A is probably the original repository, renamed (MEDIUM)", reasoning from the old name `goyal-sidhant/AnnualStatement` alone. GitHub's creation dates point the other way. B's repository was created on 31-08-2025, three minutes before `edac6a3`, and A's on 07-02-2026. Audit A's own Section 12 (its line 11155) and audit B (its Section 1 and Section 12A) record these dates. B is therefore probably the original (MEDIUM). The map flagged this too.
2. **Audit B, Section 12A, Phase 7 (its line 10537)** says Step 4's list is rebuilt "after every extraction". The code rebuilds it only in the normal-end handler (`B:gui/pyqt_main_window.py:549-561`); the error handler (`:563-565`) does not (READ). Section 4, 4.7 item 2 found this first.
3. **Audit B, Appendix X3.5 (its line 12705)**, and the consistency note of its self-review, say the program could not start for "28 minutes". The commit times give 27 minutes and 6 seconds (11:29:56 → 11:57:02), as audit B's Sections 1, 2 and 12A say.
4. **Audit A, rule A-R152** (its Section 5) says the four `MESSAGES` texts are "shown as `Warning` boxes before Dry Run and Start". Three of them are "Error" boxes (`A:gui/handlers/file_handler.py:123, 135, 139`); only "Please select at least one client to process" is a "Warning" box (`A:gui/handlers/processing_handler.py:45, 83`). Section 5, 16.4 found this; the review re-read the lines.

### Limits of this review

- Nothing that needs Windows or Excel was run. The review repeated no RAN result of Sections 3 to 10; those rest on each section's own runs and on its independent checker's re-runs, cited in place.
- Plain English in Sections 3 to 10 was checked by reading them through and by their own word lists, not term by term. The neutrality search finds words, not meanings: a judgement phrased without any of the searched words could remain.
- Sections 3 to 10 still point to the section files of audit B (`/root/merge-scratch/audit/B/…`) in places. Their content equals the final audit except for the 116 changes compared above.
