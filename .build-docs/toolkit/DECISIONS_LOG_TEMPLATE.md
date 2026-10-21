# Decisions Log: [PROJECT_NAME]

> **What is this file?** This is your project's memory. After EVERY message
> exchange during the build, Claude Code appends a 2-3 line summary below.
> When a chat session ends and the history is gone, this file preserves
> what happened — what was discussed, what was decided, and what's still open.
>
> **How it is organised:**
> - **Session 0** — what we decided while discussing the project, before the
>   build. Claude pre-fills it from that discussion.
> - **Session 1** — Steps A and B of the Build Prompt: confirming understanding
>   and writing the Build Plan.
> - **Session 2 onwards** — one session per sub-phase of the Build Plan (the
>   Build Loop in CLAUDE.md), laid out as in *Format for each sub-phase session*
>   below.
>
> **Rules:**
> - Auto-maintained by Claude Code — do not edit manually unless correcting an error
> - Claude Code appends after EVERY response, no exceptions
> - Each new sub-phase gets a new session header — not each new chat or usage window
> - The "Open" field carries forward unresolved questions
> - P-flags (technical choices Claude made without asking) are numbered across the
>   whole project — P1, P2, … — so continue from the last number used in this log
> - Lettered assumptions (A, B, C…) are for technical choices only; they stand
>   unless the author objects — silence = consent. Business questions are always asked.

---

## Format for each sub-phase session

Copy this layout when a sub-phase starts. Add a step status line as each step
ends, an exchange entry after every response, and the session entry at .9 Close.

~~~markdown
## Session [N] — Sub-phase [N.M] [name] — [DATE]

### Step status
- [DATE] | [N.M].[K] [step name] | [done / skipped — why] | [one line]

### Exchanges
#### [DATE] | Session [N] | Message [N]
- **Discussed:** [What we talked about in this exchange]
- **Decided/Built:** [What was decided or what code was written]
- **Open:** [Any unresolved question or next step]

### Session entry (written at [N.M].9 Close)
- **Built:** [What now exists that did not before]
- **Decisions:** [What was decided, and why]
- **P-flags:** [From the plan and the coders' saved choices]
- **Assumptions:** [Lettered — A, B, C…]
- **Disputes:** [Each dispute, with Fable's ruling]
- **Break-it tests:** [What Codex's tests found, and what was done]
- **Code review:** [Codex's findings, and the answers]
- **Adversarial review:** [Findings and answers — or "declined"]
- **Superseded:** [Earlier decisions this sub-phase replaced]
- **Open items:** [Carried forward]
- **Status:** [Suite count, e.g. "412 passed, 3 skipped"] · version [x.y]
~~~

---

## Session 0 — Project discussion — [DATE]

> Pre-filled by Claude from our project discussion, before the build starts.

- **Discussed:** [What the tool is for and how it will be used, as we talked it through]
- **Decided:** [Each decision we reached — one line each]
- **Assumptions:** [Technical choices only — A. …  B. … — they stand unless the
  author objects. Anything that changes what the app does with client work goes
  under Open, as a question.]
- **Open:** [Business questions still to answer — asked one at a time before the plan depends on them]

---

## Session 1 — Build Plan (Steps A and B) — [DATE]

### Step status

(One line when Step A is confirmed; one when BUILD_PLAN.md is saved)

### Exchanges

(Entries will be added automatically below)
