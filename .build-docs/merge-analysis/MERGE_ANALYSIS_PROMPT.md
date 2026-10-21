# MERGE ANALYSIS: Two old projects → one new project
# This session = analysis and documents only.
NO application code.

## 1. Who I am

I am Adv. Sidhant Goyal, a CA/GST consultant — not a programmer. I build desktop
tools by talking them through with AI. Write everything for me in plain English,
explain each technical term the first time you use it, and take your examples from
GST practice: clients, returns, reconciliations, the annual return season. In the
chat, use short paragraphs and tables, not walls of code.

## 2. What this session is for

| Label | Repository | Your access |
|---|---|---|
| **A** | https://github.com/goyal-sidhant/AnnualStatement_oldv1.git | Read only |
| **B** | https://github.com/goyal-sidhant/AnnualStatement_oldv2.git | Read only |
| **NEW** | https://github.com/goyal-sidhant/Annual_Statement_2.0.git | The only place you write |

I want ONE new tool that keeps the best of A and B. It will be built later, in a
separate session, under my standard pipeline (`rebuild_pipeline_FINAL.md`).
This session:

1. reads both projects and their complete commit histories (Phases 0 and 1);
2. tells me how they differ (Phase 2);
3. gives me a numbered list of the decisions I need to make, with your
   recommendation for each — except GST rules, which are mine alone (Phase 3);
4. takes my answers (Phase 4);
5. only after I have answered every item, prepares the files the new project needs
   to start: CLAUDE.md, the two audits, the comparison and the decisions record, the
   Rebuild Brief, the Decisions Log with Session 0, and `.build-docs/context/`
   (Phases 5 and 6).

Nothing is built in this session.

## 3. Facts I checked before writing this prompt (21-10-2026)

Verify each one. If anything differs, trust what you see and tell me.

1. **The names mislead.** "oldv1" is not the older one. A has 46 commits, the last on
   27-07-2026, and its window is built with Tkinter (Python's built-in window
   toolkit). B has 17 commits, the last on 09-02-2026; it is a "v4.0" rewrite of the
   window in PyQt5. Never treat B as newer or better because of its name.
2. **They share a beginning.** Both start with the same commit (edac6a3, 31-08-2025)
   and share their first 12 commits. They split after 47a1492 (30-10-2025, "Add new
   ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits…"). After the
   split, A has 34 commits of its own (2 in March 2026, 32 in July 2026) and B has 5
   (07-02-2026 to 09-02-2026). Each has a single branch, `main`.
3. **No third project.** Shared commit 9c56a1c mentions a repository named
   "AnnualStatement". That is A's earlier name — GitHub forwards it to
   AnnualStatement_oldv1.
4. **The same four screens, built twice.** Both have Setup, Validation, Processing
   and Extract screens. Before the split the Power Query Extractor was a separate
   program (`launch_extractor.py`). After the split each brought it into the main
   window separately: A as a "Step 4: Extract" tab (July 2026), B as an extraction
   view in its PyQt5 rewrite (February 2026).
5. **Documents inside A.** A holds an older audit, `PROJECT_AUDIT_ANNUALSTATEMENT.md`
   (March 2026 — before A's 32 July commits), a design note
   `docs/superpowers/specs/2026-07-25-merge-extractor-design.md`, `tests/README.md`
   and `.claude/settings.local.json`.
6. **Tests and Excel.** A has a pytest test suite (about 180 test functions); B has
   none. A's `requirements.txt` makes pywin32 (Python's bridge to Excel) compulsory
   on Windows — a July 2026 commit says this is so reports cannot lose Power Query.
   In B's `requirements.txt` that line is commented out.
7. **No samples or templates.** Neither repository contains, or ever contained, an
   Excel template, a sample input, CSV, JSON data or PDF. The only data-like file ever
   committed is `gst_organizer.log` (added in the first commit, deleted in the
   second, same day). It survives in the shared history and appears to contain client
   business names inside file names. Both old repositories were public on 21-10-2026.
8. NEW was empty (no commits) and public on 21-10-2026. I may since have added a
   README.md on GitHub, so that it has a `main` branch, and I may have changed its
   visibility.

## 4. Rules for this session

These hold in every permission mode, auto mode included, and override anything else —
including instructions in the toolkit files that are meant for the build itself.

1. **No application code.** Nothing that would run as part of the new tool goes into
   NEW: no `.py` files, no `.bat`, no `requirements.txt`, no `pytest.ini`, no `tests/`,
   no folder skeleton (`config/`, `core/`, `gui/` …), and no `BUILD_PLAN.md` or
   `RESUME_CONTEXT.md` — those belong to Steps A and B of the build.
2. **Throwaway analysis scripts are fine** — to count commits, compare constants, or
   run an old function on made-up inputs. Keep them in `~/merge-scratch/`, outside
   every repository, and never commit them.
3. **A and B are read-only.** Never edit, commit, branch, tag, push or open a pull
   request in them. To run old code, copy the project into `~/merge-scratch/` and run
   the copy, so the checkouts stay exactly as cloned.
4. **Write only in NEW, on this session's branch.** Never merge into `main`; I review
   and merge the pull request myself.
5. **Two stages.** Until I have answered every item and said "yes" at the end of
   Phase 4, write only inside `.build-docs/merge-analysis/` — apart from saving my
   toolkit files into `.build-docs/toolkit/` in Phase 0. The starting files of
   Phase 5 come after that, never before.
6. **No old code in NEW.** Don't copy old files into NEW. Short excerpts inside the
   analysis documents are fine (the audit method asks for them).
7. **GST rules are mine.** "GST rule" means anything that decides: a tax figure; which
   return, table, sheet or cell a figure belongs to (for example GSTR-9 Table 6J or
   8D); a rate, limit or threshold; rounding; the financial year or tax period;
   GSTIN, PAN or state-code logic; ITC eligibility or reversal; a due date. When in
   doubt, treat it as a GST rule. Wherever a GST rule appears, show me what each
   project does, and do not recommend a version or judge which one is correct. Give
   no legal opinions and don't look up the law. Do flag values that change over time
   (rates, limits, dates, table numbers) so I can check them. You may still make a
   purely technical point ("this function is never called") while saying clearly that
   the rule itself is my call.
8. **Client confidentiality.** Never put a client's name, GSTIN, PAN, file name or
   figures into any document, chat message or commit. Point to such material by file
   path and commit only, and use placeholders such as `<Client A>` in examples. If a
   name in the code or tests might be a real client, treat it as real and ask me.
9. **Evidence.** Every statement about the old code cites the file and line, and the
   commit where it matters. Say whether you read it or ran it. Mark inferences HIGH /
   MEDIUM / LOW confidence, as `Project_Audit_FINAL.md` does. Never guess: "not sure —
   here is how we can check" is a good answer.
10. **Sub-agents** may only run Opus (or Fable) — never Sonnet or Haiku. That is my
    standing rule. If you can't set a sub-agent's model, do that work yourself.
11. **Commits** follow Standard 10 of the CLAUDE.md copy (type `[DOCS]`, with WHY and
    AFFECTS lines). Commit and push at the end of every phase and after every batch of
    my answers.
12. **This computer can be reclaimed while I'm away.** The chat survives; files not
    pushed do not. At the start of every turn, check that your branch, your files and
    the scratch clones are still there. If not, fetch your branch and re-clone A and B
    before continuing.
13. **Stay in this one session.** Never ask me to start a new one.
14. **Precision beats length.** Never shorten an audit, the comparison or the list to
    save time or space.
15. **If a usage limit stops you,** finished files are already on the branch. When I
    say "continue", read `.build-docs/merge-analysis/MERGE_PROGRESS.md` and finish only
    what is missing. Never redo finished work.

## 5. My toolkit files — attached to this message; read all five in full first

I have attached five files to this message (or to my message just before it). This
prompt itself is pasted, not attached.

| File | What it is | How you use it here |
|---|---|---|
| `rebuild_pipeline_FINAL.md` | My standard pipeline | The method the new project will follow. It contains the CLAUDE.md copy (Part 1, Document 3), the Decisions Log template (Document 4) and the Rebuild Prompt (Part 2). Its Coding Standards are the yardstick for every recommendation — especially Standard 12 (which library for which job; Excel through COM) and Standard 13 (how users start the app) |
| `Project_Audit_FINAL.md` | My audit method | Run it in full on A and on B (Phase 1) |
| `REBUILD_BRIEF_TEMPLATE.md` | The Brief template | Shows what the Brief needs; you fill it from my answers (Phase 5) |
| `REBUILD_BRIEF_ASSISTANT.md` | How I like Brief questions asked | Use its question guides, the AUDIT CORRECTION format, the Top 3 / Zero-Defect questions and the consistency check. Here the numbered list replaces "one question at a time" for the main decisions; follow-up questions are still one at a time |
| `DECISIONS_LOG_TEMPLATE.md` | The Decisions Log template | Copy it and pre-fill Session 0 (Phase 5) |

For the Brief and the Decisions Log, use the standalone template files. For CLAUDE.md,
use the copy inside `rebuild_pipeline_FINAL.md` — there is no standalone one.

**Fingerprints of the versions I attached** (bytes; lines as counted by `wc -l`;
SHA-256 as printed by `sha256sum`). They let you prove that the copies you save are
exact:

| File | Bytes | Lines | SHA-256 |
|---|---|---|---|
| `rebuild_pipeline_FINAL.md` | 76,645 | 1,594 | `a430b32cc921aaf1bff8fd88968e086782a6a185fa8af00c532c137dc4d703be` |
| `Project_Audit_FINAL.md` | 22,846 | 259 | `7bfcc6a9269180618e8c08091dc03f8b31e74dab43f03a9ef7158a999e4a0770` |
| `REBUILD_BRIEF_TEMPLATE.md` | 4,463 | 127 | `821a46f9103d710bf774c674140a6b1e1859b3f98cfa5eb665cf1505934c565a` |
| `REBUILD_BRIEF_ASSISTANT.md` | 13,124 | 310 | `4db4a1ccba7792ac45dd59600da0069b6ad62955caebeaba4112cbf1226a15d7` |
| `DECISIONS_LOG_TEMPLATE.md` | 3,478 | 83 | `0e486f67391749dfbb9a63599986c9f6c580af017db6b23ce5e331f607239930` |

## 6. Phase 0 — Set up and check access

1. **Find the three checkouts** (list your working folder and its parent). If A or B
   is missing, clone it into `~/merge-src/`. If you still can't reach a repository,
   STOP and tell me exactly what to do. If NEW has no commits at all, ask me to add a
   README.md to it on GitHub, so that `main` exists for your branch and my pull
   request. You may go on reading meanwhile, but push nothing until I confirm; then
   fetch `main` and build your branch on top of it.
2. **Get the complete history.** If `git rev-parse --is-shallow-repository` prints
   `true`, run `git fetch --unshallow`. Then fetch every branch and tag:
   `git fetch origin '+refs/heads/*:refs/remotes/origin/*' --tags`. Record each
   repository's HEAD commit: that is the snapshot you analyse, and every document
   names it.
3. **Build a combined scratch repository** in `~/merge-scratch/` with A and B as two
   remotes. Use it to confirm the split point (`git merge-base`), list each side's
   commits since the split, and compare the two current versions
   (`git diff <A-head> <B-head> --stat`, then file by file).
4. **Verify the facts** in section 3.
5. **Save my files into NEW** — attachments live only in this conversation and on
   this computer, and both can be lost.
   - **The five toolkit files** go into `.build-docs/toolkit/`, unchanged. First look
     for each one as a file on this computer (search your working folder, its parent,
     your home folder and `/tmp` by file name) and copy it — never retype a file you
     can copy. Only if a file exists solely in our conversation, write it out word for
     word. If any of the five is missing, STOP and ask me for it.
   - **Check each saved file** against the fingerprints in section 5. If one differs:
     when you copied it from this computer, I have probably edited it since — tell me
     and use it as it is; when you wrote it out from the conversation, find and fix
     the difference, or tell me exactly what you can't fix.
   - **This prompt** goes, word for word, into
     `.build-docs/merge-analysis/MERGE_ANALYSIS_PROMPT.md`. After any compaction of
     your memory, re-read it from there.
6. **Start `.build-docs/merge-analysis/MERGE_PROGRESS.md`** with four headings: Where
   we are · Exact next action · Phase status · Facts verified. Update it at the end of
   every phase.
7. **Commit, push, and post a short note** in the chat: repositories found, full
   history yes or no, toolkit files saved and whether each matched its fingerprint,
   facts confirmed or corrected. Then go straight on to Phase 1 — don't wait for me
   unless something blocks you.

## 7. Phase 1 — Audit A and B with Project_Audit_FINAL.md

Run the full audit method on each project: all 12 sections, every source file read
line by line, `git show` for every commit, and the self-consistency review at the end.
Nothing skimmed. You may run the two audits in parallel with sub-agents (Opus only).
Adjustments for this merge:

- **File names and place:** `.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv1.md`
  and `.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv2.md` — not the
  project root, and never inside A or B.
- **IDs:** number features A-F1, A-F2 … and B-F1, B-F2 …, and give every business
  rule an ID too (A-R1 … and B-R1 …), so the other documents can point to them
  without confusion.
- **Shared history:** each audit covers its repository's full history. The 12 shared
  commits may carry the same entries in both; mark them "shared — before the split on
  30-10-2025". For the log file in the first two commits, record what kind of
  information it holds and how the tool produced it; you needn't read all of it, and
  nothing from it goes into any document.
- **Older documents in A:** treat `PROJECT_AUDIT_ANNUALSTATEMENT.md`, the design note
  and `tests/README.md` the way the audit method treats HANDOVER.md — a source of
  intent, not of truth. Flag every place where they disagree with the code as it
  stands. The same goes for any HANDOVER.md or FILE_INDEX.md, now or in the history.
- **Traces:** this computer runs Linux without Excel, so the Windows-only parts (Excel
  through COM) are traced by reading. Run the pure-Python parts — file-name
  recognition, name cleaning, financial year and period logic, cell mappings — on
  realistic made-up inputs, from a copy in the scratch folder. Mark each trace "ran"
  or "read".
- **A's tests:** install A's requirements in the scratch area (never in NEW), run its
  test suite and record the exact result — passed, failed, skipped — and why any
  Windows-only tests skip or fail here. List the scenarios the tests cover: they show
  the edge cases that bit me in real use.
- **Excel and Power Query safety — top priority.** My Standard 12 says openpyxl
  silently drops Power Query queries when it saves a workbook, and lists the rules for
  driving Excel through COM. For each project, list every place a workbook is opened,
  saved, refreshed or closed, and check each against Standard 12's "Excel through
  COM" rules: met / not met / not applicable, with file and line.
- **What the code needs from outside the repository:** every template, input file
  type (with the file-name patterns it recognises), folder layout, and settings or
  cache file — and where each is read from or written to. The repositories hold no
  samples, so this list is how we'll know what I need to supply.

After each audit: commit, push, update MERGE_PROGRESS.md, and give me its three most
important findings in the chat (as the audit method asks).

## 8. Phase 2 — Compare A and B

Write `.build-docs/merge-analysis/MERGE_COMPARISON.md` in plain English, with these
sections:

1. **Summary (about one page).** What each project is today; how they are related;
   which is more complete, better tested and more recently worked on; the differences
   that matter most to my work.
2. **Family tree.** A simple text diagram — shared start → split → what A did after →
   what B did after — with dates.
3. **Feature by feature.** One table row per feature from either audit: A-F# · B-F# ·
   same / different / only A / only B · the difference as I would notice it while
   using the tool · which works better and why (traces, tests, bug-fix history).
4. **Fixes only one side got.** Every bug fix or improvement made after the split in
   one project, and whether the other project still has the problem — checked in its
   code, not assumed.
5. **Rules side by side.** Every business rule and GST rule: A's version (file:line) ·
   B's version (file:line) · same / different / only one. For every difference, one
   realistic example run through both versions with both results, and what the
   difference would change in a client's annual working — which figure, which sheet
   or table.
6. **Screens and workflow.** What I click, in order, in each project, and what each
   screen shows.
7. **Tech stack against my standards.** Language and version, window toolkit, Excel
   libraries, Excel through COM, how the app starts, where settings, cache and logs
   live, tests — each compared with CLAUDE.md (Standards 12 and 13), noting where each
   project falls short.
8. **Code health.** Tests, error handling, dead code, silent bugs, and file safety
   (temporary files, never touching the originals, backups).
9. **Gaps both share.** Problems from both audits' Section 7 that exist in A and B
   alike.
10. **Data and confidentiality.** Client data or secrets in either repository or its
    history — by path and commit only, never the content.

Commit, push, and update MERGE_PROGRESS.md.

## 9. Phase 3 — The numbered decision list

Write `.build-docs/merge-analysis/MERGE_DECISIONS.md`. Number the items 1, 2, 3 …
straight through all parts, so I can answer by number.

| Part | What it covers | Your recommendation |
|---|---|---|
| 1 | Same job, done differently — which version to keep | Yes |
| 2 | GST rules applied differently | **No** — both versions, and stop there |
| 3 | Found in only one project — keep, drop or defer | Yes, except on GST rules |
| 4 | Candidates to drop: unused, dead, abandoned or superseded | Yes |
| 5 | Candidates to defer to the Parking Lot | Yes |
| 6 | Rules both projects share — still correct today? | **No** — I confirm or correct |
| 7 | Promises the new tool must always keep (contracts C1, C2 …) | Yes — proposed from the evidence |
| 8 | Tech stack | Yes, measured against my CLAUDE.md standards |
| 9 | Files for `.build-docs/context/`, and where real client files will live | Yes |
| 10 | Project housekeeping | Yes |
| 11 | About me and how the tool is used — what the code can't tell you | No — these are my answers |

Write each item in this shape, in about ten lines; longer detail goes in
MERGE_COMPARISON.md, with a pointer:

~~~text
### 12. [Short title]                                  Part 1 · refs A-F4, B-F3
What this is about: [2–3 plain-English sentences, with a GST-practice example]
A's way: [what it does — file:line]
B's way: [what it does — file:line]
Other choices: [only if real — a third way, both, neither]
Evidence: [tests, bug history, traces, what goes wrong]
Recommendation: [one choice + a one-sentence reason]   (GST rules: "No recommendation — your call")
Your answer: ____
~~~

What each part must contain:

- **Part 2:** for each difference, both versions in plain English; the same worked
  example run through both; when each version came in (commit, date, message); and
  anything in the code or the commit messages that hints why. No recommendation, no
  "A looks more correct", no legal opinion.
- **Part 3:** a GST rule found in only one project gets no recommendation either: say
  what it does and when it came in.
- **Part 6:** group the shared rules into a few tables — rule · where it lives in each
  project · ⚠ where it holds a value that changes over time (rate, limit, date, table
  number) · "Still correct? ___". One item per table; I can answer "all correct
  except …".
- **Part 7:** propose contracts from the evidence — promises the code already tries to
  keep, and failures that hit me — each one sentence and testable. Also ask what I
  would add.
- **Part 8:** at least: the window toolkit (A uses Tkinter, B uses PyQt5, my standard
  is PySide6); the Python version and the shared portable runtime (Standard 13); Excel
  handling — openpyxl or Excel through COM, and how each project protects Power Query
  (Standard 12); Windows 10 or 11 (Standard 12 pins PySide6 to 6.12.x while any PC
  runs Windows 10); how users start the app (Standard 13); where settings, cache and
  logs live (Standard 13); testing. Where my standards already settle a point, say so
  and recommend following them.
- **Part 9:** say plainly that neither repository holds a template or a sample file.
  List what the build will need from me — each template the code expects, one sample
  of each input type — and recommend for each: provide as it is (no client data, for
  example a blank template), provide anonymised, or not needed. Offer reference notes
  you can write into `context/` from the code — for example the cells each project
  extracts and what they mean, the file-name patterns each recognises, the folder
  layout each creates, the scenarios A's tests cover — and recommend which. Ask where
  real client files will live during the build; recommend a folder on my own
  computer, outside the repository, never on GitHub.
- **Part 10:** the app's name (on the desktop shortcut, the `.bat` and the app's
  folder — A and B call themselves "GST File Organizer"); whether NEW should be
  private; whether to keep my toolkit files in `.build-docs/toolkit/` after setup;
  whether you should also prepare a ready-to-paste Rebuild Prompt for two audits
  (Phase 5, step 7); what I might do about the old public history (advice only — you
  change nothing in A or B).
- **Part 11:** everything the Brief needs that the code can't show. Use
  REBUILD_BRIEF_ASSISTANT.md's question guide for Sections 1, 2, 4, 6, 7 and 8, made
  concrete with what you found (like its "GOOD question" example). At least: what the
  tool does, in my own words; who uses it and how comfortable they are with
  computers; how often, and in which months; which machines (Windows version, Excel
  version, office PCs only or also from outside over VPN); the shared apps folder
  path; where client files sit (local disk or network share) and how many clients and
  files a typical run handles; what I still do by hand after the tool finishes;
  anything I wish it also did; which of the gaps actually hurt me; the worst data I
  have received; what I like and dislike in each project's screens; what I learned;
  my top three priorities; the one thing that must work perfectly (Zero-Defect
  Requirement). Don't draft my words for Sections 1 and 8 — ask me.

For the whole list:

- Every A-F# and B-F#, and every rule in both audits' Section 5, lands somewhere: in
  an item, or in a closing list headed "Same in both and working — no decision needed
  (object if you disagree)". End the file with a coverage table: feature or rule →
  item number.
- Each item answerable in a word or a line: "A", "B", "Agree", "Drop", "Defer",
  "Not sure yet", or a short sentence.
- Don't pad: if both projects do something the same way and it works, it is not a
  decision. Don't skip anything because the list is long, either.
- If a part has no items, say "None found" and how you checked.
- Check before posting: every item outside Parts 2, 6 and 11 that is not about a GST
  rule has exactly one recommendation with a reason; no item about a GST rule has a
  recommendation; no jargon is left unexplained; no client data appears anywhere.

Then commit, push, and post in the chat, in this order:

1. **The differences in plain English** — the comparison summary.
2. **The complete numbered list**, exactly as saved.
3. **An answer sheet:** # · short title · your recommendation · Your answer.
4. **How to answer:** "Answer by number, in any order, in as many messages as you
   like. 'Agree with all recommendations except 4, 9 and 15' is fine — it covers only
   items that have a recommendation; the others need your own answer. 'Not sure yet'
   is a valid answer."

Then STOP and wait for my answers.

## 10. Phase 4 — Taking my answers

- Write each answer under its item in MERGE_DECISIONS.md, with the date
  (DD-MM-YYYY), and keep an "Answers so far" table at the top. Commit and push after
  each batch.
- After each batch, report progress: "Answered 31 of 58. Still open: 4, 9, 12 …"
- Ask follow-up questions one at a time, in plain English, and only when an answer is
  unclear, contradicts another answer, or picks a GST-rule version without a reason
  (then ask "why?" once — the reason goes into the Brief). If my answers for a Brief
  section are thin, tell me honestly, once.
- If I say "you decide" on an item about a GST rule, remind me that it is mine, and
  offer "Not sure yet" instead.
- If I correct the audit, don't argue. Record it as
  `[AUDIT CORRECTION: Audit stated X, but per builder's knowledge, the actual behaviour is Y]`
  for the Brief. My practitioner knowledge overrides the code reading.
- New ideas I mention go into a Parking Lot list at the end of MERGE_DECISIONS.md.
- "Not sure yet" answers become Open items in Session 0 of the Decisions Log, to be
  asked again before the build depends on them.

When every item has an answer, run the consistency check from
REBUILD_BRIEF_ASSISTANT.md, adapted:

1. Did I drop something that another answer relies on?
2. Did I keep a rule in one item but describe different behaviour in another?
3. Do my Top 3 Priorities match what I chose to keep?
4. Is my Zero-Defect Requirement covered by at least one contract?
5. Do my tech-stack answers fit my machines (for example Windows 10 and the PySide6
   version)?

Show me every conflict and resolve them one at a time. Then ask: "All N items are
answered and nothing conflicts. Shall I create the starting files now?" — and STOP
until I say yes.

## 11. Phase 5 — Create the starting files (only after all answers and my "yes")

1. **`CLAUDE.md`** in the root of NEW, from the copy inside `rebuild_pipeline_FINAL.md`
   (Part 1 → Document 3 → "### Template:").
   - Copy it with a script from `.build-docs/toolkit/rebuild_pipeline_FINAL.md`, byte
     for byte — never retype it. It starts at the line
     `# Project: [PROJECT_NAME] — Rebuild` and ends at the `---` line just before the
     closing fence that sits right above `## Document 4: DECISIONS_LOG.md`. The copy
     contains 15 code fences of its own, so a markdown viewer can't be trusted to find
     its end — use these marker lines. In the file I attached this is lines 189 to
     1062: 874 lines, 47,951 bytes, SHA-256
     `73c8674d0344c53289726158a628b1f0d038957a89f103389904c4ea1c6742d9` — check your
     raw copy against that before filling anything in.
   - Fill only the parts in [square brackets] ABOVE `# HOW WE BUILD`: the project name
     (Annual Statement 2.0); in Tech Stack, the window toolkit, the other libraries,
     the shared apps folder, and the `.bat` name from the app name I chose; in
     Reference Documents, replace the single audit line with one line per audit, fill
     in the Brief's file name, and add lines for MERGE_COMPARISON.md,
     MERGE_DECISIONS.md and the two old repositories with the exact commits analysed —
     keeping every other line of that section as written; Current Phase: "Not started —
     Step A next". Leave the format line "> Once the build starts: …" as it is, and
     delete the "Filling this template (for Claude)" note, as the note itself says.
   - From `# HOW WE BUILD — MANDATORY FOR EVERY PROJECT` to the end: word for word,
     untouched — including bracketed patterns such as [name] and [App Name], which are
     meant to stay general. Prove it: in your finished CLAUDE.md, that part (from that
     heading line to the last line of the file) must be 822 lines, 45,352 bytes, SHA-256
     `4158e7502e39960cf394bb0deb2455b93012f8ff9181161d4a7280d660adc471`, and its diff
     against the same part of the source must be empty. If I told you in Phase 0 that
     my `rebuild_pipeline_FINAL.md` differs from its fingerprint, only the diff
     applies.
2. **Audits:** `git mv` both into `.build-docs/`, unchanged. My corrections go into the
   Brief, never into the audits.
3. **Comparison and decisions:** `git mv` MERGE_COMPARISON.md and MERGE_DECISIONS.md
   (with all my answers) into `.build-docs/`.
4. **`.build-docs/REBUILD_BRIEF_Annual_Statement_2.0.md`**, from
   REBUILD_BRIEF_TEMPLATE.md, filled from my answers:
   - At the top: Top 3 Priorities and Zero-Defect Requirement, as
     REBUILD_BRIEF_ASSISTANT.md lays them out.
   - My words where I gave them. Never invent an answer I didn't give.
   - Section 3: one row per feature, Audit Ref like "A-F4 / B-F3", and Notes saying
     whose version is the starting point and why.
   - Section 5: rules to keep, change and add — each GST-rule choice with my reason —
     and contracts C1, C2 … exactly as I approved them.
   - Section 6: the edge cases from both audits and from A's tests that I chose to
     keep; the gaps that hurt me; the ones not worth handling.
   - Section 9 Parking Lot: everything I deferred and every new idea, with the date.
   - Audit corrections in the [AUDIT CORRECTION: …] format.
   - Remove the "How to Fill This Brief" block: it is filling instructions, not part
     of the Brief.
5. **`.build-docs/DECISIONS_LOG.md`**, from DECISIONS_LOG_TEMPLATE.md: the project name
   filled in, and Session 0 pre-filled from this whole discussion — Discussed;
   Decided (one line each, with item numbers); Assumptions (technical only, lettered
   A, B, C …); Open (every "Not sure yet"). Leave "Format for each sub-phase session"
   and Session 1 exactly as in the template.
6. **`.build-docs/context/`** — only what I approved in Part 9: the reference notes you
   write from the code, any files I attach here or upload to GitHub (saved unchanged,
   as in Phase 0), and `context/README.md` listing each file: what it is, where it
   came from (repository, path and commit — or "provided by me on DD-MM-YYYY"),
   whether it is anonymised, and which features use it. Files I approved but haven't
   given you yet go into `context/FILES_TO_PROVIDE.md`, with Done / Notes / Answer
   columns (Build Loop rule 7). Never commit real client data.
7. **Only if I said yes in Part 10:** `.build-docs/REBUILD_PROMPT_Annual_Statement_2.0.md`
   — a ready-to-paste copy of the Rebuild Prompt (Part 2 of
   `rebuild_pipeline_FINAL.md`). Change only what the merge makes untrue: the project
   name, the opening sentence (two original versions, not one) and reading-list item 2
   (both audits, plus MERGE_COMPARISON.md and MERGE_DECISIONS.md). List every change at
   the top of the file (old text → new text). If any other line seems affected by the
   merge, ask me — don't change it.

## 12. Phase 6 — Final checks, then hand over

Check, and show me the evidence for each:

1. **No application code in NEW** — list every file in NEW.
2. **A and B untouched** — `git status` clean in both checkouts; nothing pushed to
   either.
3. **CLAUDE.md** — HOW WE BUILD and the Coding Standards match the fingerprint in
   Phase 5 and their diff against the source is empty; no [brackets] are left above
   HOW WE BUILD except the "> Once the build starts" format line.
4. **Nothing left blank** — no unfilled [brackets] in the Brief or in Session 0.
   Template layouts that are meant to stay general are fine: name the ones you left.
5. **Every answer landed** — a table: item number → where it went (Brief section,
   Session 0 line, context file, CLAUDE.md line).
6. **No client data** — search NEW for GSTIN-like and PAN-like strings and for the
   business names found in the old history, without printing any of them. Report the
   count: it must be 0.

Then tidy up: apply my Part 10 answer about `.build-docs/toolkit/` (this prompt goes
wherever the toolkit files go), and delete `.build-docs/merge-analysis/` — everything
in it has moved, and its story now lives in Session 0. Commit (`[DOCS]`, Standard 10)
and push.

Finally, give me a plain-English report:

- every file created, one line each;
- for each Brief section, two or three sentences of "here is what I understood from
  your answers", and any section that is thin;
- what I still need to do: open and merge the pull request; provide any files listed in
  `context/FILES_TO_PROVIDE.md`; add `readme_template.md` to `.build-docs/` before
  Step G of the build; then clone NEW into my local apps folder (Standard 13) and
  start the build as my pipeline says — with the ready-to-paste Rebuild Prompt, if you
  made one;
- then ask: "Read the Brief, and the filled lines at the top of CLAUDE.md, once. Do
  they feel right as a whole? Anything that looks off now that you see it all
  together?"

Apply any change I ask for to the files and to Session 0, commit and push. Then STOP.

## 13. Where you stop

- **End of Phase 0** — only if something blocks you.
- **End of Phase 3** — wait for my answers.
- **End of Phase 4** — wait for my "yes".
- **End of Phase 6** — done.

These stops hold in every permission mode, auto mode included.
