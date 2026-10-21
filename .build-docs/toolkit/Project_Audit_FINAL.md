# Project Archaeology — Complete Audit Prompt v2.0

I need you to perform a complete "Project Archaeology" — a deep, line-by-line audit of this entire project and produce a single comprehensive Markdown document that captures everything someone would need to understand, maintain, and rebuild this tool from scratch.

## Who I Am

I am a non-coder professional (CA/GST consultant) who built this tool through AI-assisted conversations over time. The original chat history is lost. I need you to reverse-engineer not just WHAT the code does, but WHY certain decisions were made — infer the reasoning from the code patterns, comments, variable names, and structure.

## Additional Context Documents

This project may contain HANDOVER.md and FILE_INDEX.md files in the project root or bundled alongside the source code. These are not part of the application itself — they are structured knowledge-transfer documents generated at the end of the original AI-assisted development conversation, capturing the full decision history, rejected approaches, domain rules, open questions, and file-by-file status as of that conversation's end.
If these files exist, read them in full before beginning the audit. They are the closest thing to a developer diary this project has. Treat them as a primary source of intent and reasoning — but NOT as a source of truth for what the code actually does. The code may have diverged from what HANDOVER.md describes (due to subsequent changes, incomplete implementation, or errors in the handover itself). Your job is to audit the code as it exists now, use HANDOVER.md and FILE_INDEX.md to inform your understanding of why things were built a certain way, and flag any discrepancy between what these documents claim and what the code actually does. Such discrepancies are high-value findings — they reveal either post-handover drift or original documentation errors, both of which matter for reconstruction.

## Your Process

1. First, scan the entire project directory tree — every file, every folder, **including hidden directories and dotfiles** (`.claude/`, `.github/`, `.vscode/`, `.env`, `.gitignore`, etc.).
2. Read every source file line by line. Do not skip or skim.
3. **CRITICAL: Do NOT assume what a function, method, or variable does based on its name alone.** A function called `validate_gstin()` might do far more (or far less) than validation. Read the actual implementation — every line, every condition, every return path. Document what the code ACTUALLY does, not what its name suggests it does. Names can be misleading, incomplete, or outdated from earlier versions. The code itself is the only source of truth.
4. **CRITICAL: After documenting any feature that involves shared state (files, caches, configs, database entries, environment variables, registry keys), trace the WRITE path and the READ path across every file that touches that state. Confirm they are consistent — same file paths, same formats, same keys, same serialization. Mismatches between writer and reader are among the most common bugs in multi-module projects. Do not assume two files agree on a path or format — verify by reading both.**
5. For configuration files (requirements.txt, package.json, .env, etc.), note every dependency and its likely purpose.
6. For data files (JSON, CSV, Excel templates), note their structure and role.
7. Look for comments, docstrings, TODOs, and FIXMEs — these are clues to original intent. But verify comments against actual code behaviour — comments can be stale or wrong.
8. Identify patterns — repeated validation logic, error handling approaches, naming conventions — these reveal design philosophy. **For every pattern you identify, verify it is applied consistently across ALL files that face the same concern. If File A handles platform-specific imports with a graceful fallback but File B imports the same module unconditionally, flag the inconsistency — these are often bugs or incomplete refactoring, not intentional design choices.**
9. **Scan for common silent bugs** — code that runs without errors but produces incorrect output. Specifically look for: lambda/closure variable capture in loops (variables captured by reference, not value), mutable default arguments, off-by-one errors in range/slice operations, string comparison where numeric comparison was intended (or vice versa), `==` vs `is` confusion with `None`/`True`/`False`, variables used after a loop that hold only the last iteration's value, and thread-unsafe shared state access without locks.
10. **Run `git log --oneline --all` to get the full commit list, then run `git show <commit_hash>` for EVERY SINGLE COMMIT — no exceptions.** Read every diff, every line added, every line removed, every file touched. Do not summarize from commit messages alone — commit messages can be vague or misleading. The actual diff is the truth. The git history is the chronological diary of this project's entire life — it reveals the order in which features were built, why files were restructured, what bugs were discovered in real use, which approaches were tried and abandoned, and most importantly, the REASONING behind every change. This step is not optional and not skimmable. Every commit matters. **For every commit, note which branch it belongs to. Clearly separate the main/production branch timeline from unmerged/feature branches.**

## Output Structure

Produce a single Markdown file called `PROJECT_AUDIT_[PROJECT_NAME].md` with the following sections. Be thorough. Write in plain English that a non-programmer can understand. Where technical terms are unavoidable, add a brief parenthetical explanation.

---

### SECTION 1: EXECUTIVE SUMMARY

Write 5-8 sentences covering:

- What this tool does in plain language
- Who it's built for and what problem it solves
- The core technology stack (e.g., "Built with Python, uses PyQt5 for the desktop interface, openpyxl for Excel file handling")
- Current state: does this appear complete, work-in-progress, or abandoned?

### SECTION 2: PROJECT STRUCTURE MAP

- Full directory tree with annotations explaining what each folder and file does
- **Include ALL files, including hidden directories and dotfiles** (`.claude/`, `.github/`, `.vscode/`, `.env`, `.gitignore`, etc.). These reveal development tooling, CI/CD, IDE settings, and AI-assisted development patterns. Tag them as [TOOLING] or [CONFIG] as appropriate.
- Mark files as: [CORE] (essential logic), [UI] (interface), [CONFIG] (settings), [DATA] (templates/samples), [UTIL] (helper/utility), [TEST] (testing), [DOCS] (documentation), [TOOLING] (development tools/IDE/CI), [GENERATED] (auto-created, safe to delete)
- Identify the entry point — which file do you run to start the application?
- **Identify dead code at every level**: dead files (never imported), dead classes (never instantiated), dead methods (defined but never called from anywhere), dead constants/variables (defined but never referenced), and dead parameters (accepted but never used). For methods, verify by searching for every call site — a method that only appears in its own definition is dead. Do not stop at file-level analysis.

### SECTION 3: FEATURE INVENTORY

For EACH distinct feature the tool provides, document:

- **Feature name**: A clear, descriptive name
- **What it does**: Plain English explanation
- **Where it lives**: Which files/functions implement this
- **Input → Output**: What goes in, what comes out
- **User interaction**: How does a user trigger or use this feature (button click, file selection, automatic, etc.)
- **Concrete trace**: For each feature, mentally trace one realistic example through the full code path. Pick a concrete input (e.g., "user scans a folder with 3 clients, 2 complete and 1 missing files") and walk through every function call, every variable assignment, every conditional branch, every loop iteration. Document what the output would actually be. This step catches bugs that reading code "in general" misses — the devil is always in a specific execution path. If the traced output differs from what the feature description claims, flag it as a bug.

Number each feature (F1, F2, F3...) for easy reference in later sections.

### SECTION 4: DATA FLOW

Trace the complete journey of data through the application:

- Where does data enter? (file upload, user input, API call, clipboard, etc.)
- What transformations happen at each step?
- Where does data exit? (file saved, displayed on screen, copied to clipboard, sent somewhere, etc.)
- Draw this as a simple text-based flow diagram using arrows (→)

### SECTION 5: BUSINESS RULES AND DOMAIN LOGIC

This is the most critical section for rebuilding. Document every business rule embedded in the code:

- Validation rules (e.g., "GSTIN must be exactly 15 characters", "tax rate must be one of: 0, 5, 12, 18, 28")
- Calculation logic (e.g., "IGST applies when supplier state ≠ recipient state")
- Formatting rules (e.g., "invoice numbers are zero-padded to 6 digits")
- Conditional logic (e.g., "if return type is B2B, include GSTIN; if B2C, use state code instead")
- Thresholds and limits (e.g., "B2C invoices above ₹2.5 lakh go in B2CL table")
- Any hardcoded values, constants, or lookup tables with their purpose — **including operational/safety constants** (retry counts, timeout values, size limits, temp file prefixes, etc.) that affect runtime behaviour

For each rule, note the exact file and line number where it's implemented.

### SECTION 6: EDGE CASES AND DEFENSIVE CODE

List every edge case the code explicitly handles:

- What happens with empty files?
- What happens with missing columns or unexpected data formats?
- What happens with special characters, Unicode, or encoding issues?
- What happens with very large files or very small datasets?
- What happens with duplicate entries?
- What happens with null/blank/zero values?
- Any try-except blocks — what errors are caught and how are they handled?

Rate the defensive coding as: ROBUST (handles most edge cases), MODERATE (handles common cases), or FRAGILE (minimal error handling, likely to break with unexpected input).

### SECTION 7: EDGE CASES NOT HANDLED (GAPS AND VULNERABILITIES)

Equally important — what does the code NOT handle that it probably should?

- Missing validations that could cause silent data corruption
- Error conditions that would crash the application
- Input variations that real-world users might provide but aren't accounted for
- Performance bottlenecks with large datasets
- Security concerns (hardcoded passwords, unvalidated file paths, etc.)
- Accessibility or usability gaps in the interface
- **Silent logic bugs**: Code that runs without errors but produces incorrect output. Report any instances found during the antipattern scan (Step 9 in Your Process) and during the concrete feature traces (Section 3). Include: lambda/closure bugs, race conditions in threaded code, variables that silently hold stale values, and any place where the code's actual output differs from its apparent intent.
- **Scenario-based failure analysis**: For each major feature, ask "what would break if...": the working directory is different from the development directory, the tool is run by a different OS user, files are on a network drive that disconnects mid-operation, the tool is run for the second time on the same data, two instances run simultaneously on the same target, the OS is different from the development machine, or the language runtime / library version changes. These "what if" scenarios catch environment-dependent bugs that static code reading misses.

### SECTION 8: DESIGN DECISIONS (INFERRED)

Based on code patterns, infer and document the likely reasoning behind key decisions:

- Why was this framework/library chosen over alternatives?
- Why is the code structured this way (modular vs monolithic)?
- Why are certain features implemented in a complex way when simpler options exist? (Often there's a domain-specific reason)
- Any evident iterations — places where code was clearly refactored, patched, or evolved from an earlier approach
- Commented-out code blocks — what do they suggest about abandoned approaches?
- **Pattern consistency**: For every design pattern you identify (error handling style, platform checks, import guards, logging approach, config access method, API design), verify it is applied consistently across ALL files that face the same concern. Flag any file that deviates from the project's own established pattern — note what the established pattern is, which file(s) follow it, and which file(s) break it.
- **Re-export / API patterns**: Note any `__init__.py` re-export patterns, facade classes, or similar API design choices that simplify imports but create maintenance requirements (e.g., new functions must be added to `__all__` lists).

Mark each inference with a confidence level: HIGH (clear from code), MEDIUM (reasonable inference), or LOW (speculation based on limited evidence).

### SECTION 9: DEPENDENCIES AND ENVIRONMENT

- Complete list of external libraries with version numbers and purpose
- Python version requirements (if applicable)
- OS-specific dependencies or behaviours
- File system assumptions (paths, folder structures, network drives)
- Any external services, APIs, or tools the application depends on
- Setup steps needed to run this on a fresh machine
- **For OS-specific modules**: Note which files import them and whether the import uses a graceful fallback pattern (try/except, platform check) or crashes on unsupported platforms. Flag any inconsistency in import handling across files.

### SECTION 10: UI/INTERFACE DOCUMENTATION

If the tool has a user interface:

- Describe every screen, window, or dialog
- Document every button, input field, dropdown, checkbox — what it does and what triggers it
- Note the layout approach and any styling choices
- Identify the user workflow — the typical sequence of actions a user performs
- Screenshot descriptions (describe what each screen looks like since you can't take actual screenshots)

### SECTION 11: RECONSTRUCTION BLUEPRINT

This is the actionable section — a step-by-step plan for rebuilding this tool from scratch:

1. **Foundation**: What to set up first (environment, dependencies, project structure)
2. **Core Logic**: Build order for features (which depends on which)
3. **Feature Priority**: Rank features by importance — what's essential vs nice-to-have
4. **Known Improvements**: Based on gaps identified in Section 7, what should be done differently in a rebuild
5. **Estimated Complexity**: Rate each feature as SIMPLE (< 1 hour), MODERATE (1-3 hours), or COMPLEX (3+ hours) for AI-assisted rebuilding
6. **Suggested Skill**: For each major feature, suggest what a Claude Code Skill (SKILL.md) could encode to make this pattern reusable across future projects

> **IMPORTANT CONSTRAINT FOR ITEM 4**: Before recommending any alternative approach or technology, check Section 8. If Section 8 explains why the current approach was chosen over the alternative you are about to recommend, do NOT recommend that alternative — you would be reintroducing the exact problem the original developer already solved. Instead, recommend a THIRD option that addresses the current approach's fragility without reintroducing the original problem. Every improvement recommendation must respect the constraints that shaped the original design. If you cannot find a third option, state the trade-off honestly rather than recommending a known-bad alternative.

### SECTION 12: CHANGELOG ARCHAEOLOGY (FROM GIT HISTORY)

This section must be built from reading EVERY commit diff — not summaries, not samples, not just "important" ones. Every commit exists because a decision was made. Your job is to extract that decision.

> **BRANCH AWARENESS**: For every commit, note which branch it belongs to (main, feature branch, unmerged PR, etc.). Clearly separate the main/production branch timeline from unmerged branches throughout all sub-sections below. When attributing a feature or fix to a commit, verify the commit is on the active/production branch. If a feature exists in the current code but the only commit that appears to "add" it is on an unmerged branch, the attribution is wrong — the feature must have been added in a different main-branch commit or was present from the initial commit. Flag and resolve such discrepancies.

**12A. Complete Development Timeline**

For EVERY commit, chronologically, document:

- **Commit hash and date** (and branch name if not main)
- **What changed**: Files touched, lines added/removed — describe the actual code changes in plain English
- **Why it changed**: Infer the reasoning — was this a new feature, a bug fix, a refactor, a dependency update, a response to a real-world problem?
- **What it reveals**: What does this change tell us about the project's rules, logic, data flow, edge cases, or design philosophy?

Group commits into natural development phases as they emerge (e.g., "Phase 1: Initial scaffold and core data model", "Phase 2: Validation logic", "Phase 3: UI development", "Phase 4: Real-world bug fixes", etc.).

**12B. File Evolution Map**

Using the git history, trace how each core file evolved over time:

- When was it created? What was its original purpose?
- How did it change across commits? Did its role expand, shrink, or shift?
- Was it ever renamed, split into multiple files, or merged from other files?
- This reveals the architecture's evolution — how the project grew from a simple script into a structured application.

> **VERIFICATION RULE**: For any file marked as "never changed" or "stable since creation," verify this claim against the current code. If the file imports a function, constant, or module that was added or modified in a later commit, the file MUST have been changed — update the evolution map accordingly. The current code is the ground truth; git log descriptions and commit messages can be incomplete (e.g., a commit may touch a file without the message mentioning it).

**12C. Logic and Rule Evolution**

Track how business rules, validation logic, and calculations changed across commits:

- Were rules added incrementally (suggesting they were discovered during real-world testing)?
- Were rules modified (suggesting the original understanding was wrong or regulations changed)?
- Were rules removed (suggesting they became obsolete or were replaced)?
- This is critical for rebuilding — it shows which rules were foundational vs. which were discovered through trial and error.

**12D. Edge Cases Discovered in Production**

Bug-fix commits are gold. For every commit that fixes a bug or handles a new edge case:

- What broke? What input or scenario caused the failure?
- How was it fixed? Was it a patch or a proper solution?
- Does the fix suggest other similar edge cases that might still be unhandled?

**12E. Abandoned Approaches and Reversals**

Identify commits where:

- Code was added and later removed or replaced in a subsequent commit
- A different library or approach was swapped in
- A feature was started but never completed (partial implementation still in codebase)
- **Dead methods or commented-out code** that remain in the current codebase as remnants of earlier approaches
- Document what was tried, why it likely failed, and what replaced it — these are expensive lessons that should not be repeated in a rebuild.

**12F. Dependency Evolution**

Track changes to requirements.txt, package.json, or any dependency files:

- When were dependencies added? Why? (Correlate with feature commits)
- Were any dependencies removed or replaced? What triggered the switch?
- Version upgrades — were they deliberate (to fix something) or routine?
- **Dependencies used in code but missing from dependency files** (e.g., `pywin32` used via `import win32com` but absent from `requirements.txt`) — flag these as documentation gaps.

**If no git history exists**: State this clearly and fall back to: version numbers or date stamps in code/comments, TODO/FIXME comments suggesting planned work, multiple approaches to the same problem (suggesting iteration), commented-out code blocks, and any README, HANDOVER.md, or documentation files.

---

## Self-Consistency Review (MANDATORY FINAL STEP)

**After completing all 12 sections, you MUST perform a full self-consistency review before saving.** Re-read the entire document from Section 1 through Section 12 and verify:

1. **No cross-section contradictions**: No claim in any section contradicts a claim in another section. Pay special attention to:
   - Section 8 (Design Decisions) vs. Section 11 (Reconstruction Blueprint) — never recommend replacing something you have justified with HIGH confidence
   - Section 3 (Feature Inventory) vs. Section 7 (Gaps) — if you described a feature as working in Section 3, but your concrete trace or gap analysis shows it is actually broken, update Section 3 and move the issue to Section 7
   - Section 12B ("never changed" / "stable" claims) vs. features/imports documented in Sections 3-5 — if a file imports something added in a later commit, it was changed
   - Section 2 ("no dead code") vs. actual dead methods/constants found during analysis

2. **All features verified**: Every feature described in Section 3 has been verified through a concrete trace, not just described from reading function signatures.

3. **All shared state verified**: Every cache, config file, temp file, or shared resource has had its write path and read path compared across all files that touch it (as required by Process Step 4).

4. **All patterns verified for consistency**: Every design pattern identified in Section 8 has been checked for consistent application across the full codebase (as required by Process Step 8).

5. **Flag all corrections**: If the self-consistency review reveals that an earlier section's claim needs to be corrected, update the claim in place AND add a note: `[Corrected during self-review — originally stated X, but Section Y revealed Z]`.

---

## Formatting Rules

- Use clear Markdown headers (##, ###) for sections
- Use tables for structured comparisons
- Use code blocks with syntax highlighting for any code references
- Use > blockquotes for important warnings or notes
- Cross-reference sections (e.g., "See Feature F3 in Section 3")
- Keep language accessible — explain technical terms in parentheses on first use
- Be specific — "Line 45 of main.py" not "somewhere in the main file"

## Final Instruction

Save the completed audit as `PROJECT_AUDIT_[PROJECT_NAME].md` in the project root directory. After saving, give me a brief verbal summary of the three most important findings — things I'd want to know first before diving into the full document.
