# A-T — Working notes: project A's test suite (snapshot bf2f1dd, 27-07-2026)

Task: read every line of A's tests (`tests/conftest.py`, all 18 `tests/test_*.py`, `tests/README.md`, `pytest.ini`), run the suite twice (with and without a virtual display), explain every failure and skip, map each test to the scenario it pins and the commit it guards, map what has no tests, check the README and commit-message claims, and note tests that write outside a temp folder.

All paths are repository-relative to project A (goyal-sidhant/AnnualStatement_oldv1, HEAD bf2f1dd56f088d019bed6022428aab9bf0eb423d) unless marked otherwise. Commits are from `git -C /root/merge-scratch/combined`, branch A/main.

Tags used: **[READ]** = from reading the code or the git output. **[RAN]** = I executed it on this Linux machine. **[RAN-sim]** = I executed a small simulation (for example Windows path rules faked with `PureWindowsPath`). Inferences carry **HIGH / MEDIUM / LOW** confidence. **[GST RULE]** marks logic that decides a tax-related classification, period, state code and so on. I only describe what the code does with these. I make no judgement on whether they are right. **TIME-VARYING** marks values that can change over time.

---

## 0. Plain-English glossary (terms used below)

- **test / test suite**: a small program that runs part of the app with known inputs and checks the result. The "suite" is all of them together.
- **pytest**: the tool that finds and runs the tests and reports passed / failed / skipped.
- **collected**: the number of individual tests pytest found before running anything.
- **passed / failed / skipped / error**: *passed* means the check held. *Failed* means a check did not hold, or the test crashed. *Skipped* means the test chose not to run, for example because no screen was available. *Error* means a problem while preparing or collecting the test. There were no errors at HEAD.
- **fixture**: a helper that prepares something before a test runs, such as a temporary folder or a fake Excel file.
- **parametrize**: one test function run several times with different inputs. Each input counts as a separate test.
- **tmp_path / temp folder**: a fresh, throw-away folder that pytest creates for each test, under the system temp area (here `/tmp/pytest-of-root/...`).
- **monkeypatch**: temporarily replacing a real function with a fake one during a test, for example a fake `pip` run.
- **subprocess / probe**: a test that starts a separate Python program, which runs a small script and reports back as text. A's tests use this for windows that need their own Tk root.
- **Tk / Tk root / tkinter**: the GUI toolkit A uses. The "root" is the main window object. Most Tk work needs a display (a screen).
- **display / xvfb**: Linux needs a display server to open windows. `xvfb-run` provides an invisible ("virtual") one. Windows PCs always have a display.
- **COM / pywin32 / pythoncom / win32com**: COM is the Windows mechanism by which Python drives Microsoft Excel. pywin32 is the package that provides it, and `pythoncom` is one of its modules. None of it exists on Linux.
- **mutation check**: deliberately re-introducing a bug in a scratch copy and running the tests to see whether they notice. "CAUGHT" means a test failed. "SURVIVED" means no test noticed.
- **characterization test**: a test that records what the code does now, so that later changes cannot alter it silently. It does not claim the behaviour is right.
- **bytecode / `__pycache__`**: cached compiled Python files that Python writes next to the source unless told not to (`-B`).

---

## 1. Headline findings (details and evidence in later sections)

1. **Run results confirmed exactly** [RAN]. Both runs collected 288 tests.
   - With a virtual display: `1 failed, 287 passed` (3.56 s, repeated 3 more times with the same result).
   - Without a display: `14 failed, 241 passed, 33 skipped` (1.48 s, repeated twice with the same result).
2. **The single failure under the virtual display** is `tests/test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only`. It fails because Linux does not treat `\` as a folder separator: `Path(r"C:\gst\itc.xltx").name` returns the whole string. It would pass on Windows (HIGH, [RAN-sim]).
3. **Without a display, the 13 "window in a subprocess" tests FAIL instead of skipping.** These are in `test_extractor_standalone.py` (6) and `test_main_window_tabs.py` (7). The probe scripts catch the "no display" error themselves and still print a result, so the fixture's skip branch is never reached. The 33 skips are the in-process GUI tests that use the shared `tk_root` fixture.
4. **On the author's Windows PC** I expect all 288 to pass with the interpreter that has pywin32 (MEDIUM-HIGH, cannot run here). With the `py` interpreter, which commit messages say lacks pywin32, `test_requirements_check.py::TestRealRequirements::test_openpyxl_is_actually_installed_here` would fail with a misleading message (HIGH, [RAN-sim]).
5. **One test writes outside the temp folder** [RAN]. `test_main_window_tabs.py` builds the real main window. That:
   - creates `<system temp>/gst_excel_temp` (core/excel_handler.py:53-54);
   - always reads the user's real settings file `~/.gst_organizer_cache.json`;
   - rewrites that file when it says dark mode is on (gui/main_window.py:388-389 → :196 → gui/utils/cache_manager.py:45-46).
   On the author's PC this test can modify his real app settings.
6. **Coverage is uneven** [RAN, function-level recorder].
   - Never executed by any test: the whole folder-organising engine (core/file_organizer.py, 0 of 12 functions), Excel report creation (core/excel_handler.py, 1 of 15), the Power Query refresh and extraction (report_processor.py, 3 of 15), the consolidator (data_consolidator.py, 0 of 6) and Step 3's worker (processing_handler.process_files_thread).
   - Also untested: `main.py`, `launch_extractor.py` and `extractor_main.py`.
   - The [GST RULE] cell-mapping table power_query_extractor/config/cell_mappings.py has no tests.
7. **Mutation checks** [RAN, 30 mutations]. The tests do catch the original bugs for:
   - 33d0baf: sanitise abbreviation and the re-entry guard;
   - 4219ffb, 5c8ce95, bfcb24a, 2aaa2b1 (name matching), c463256;
   - d64d3a9: both scroll bugs;
   - bf2f1dd, in its plain form.

   They do NOT catch a regression of:
   - c89c1d4: the dark-mode recompute. The source-text test is defeated by the word `refresh_status` inside a `hasattr` check;
   - 725fbaf, 169c7b5, 7347383, c1583d6 (version folder), c1f425e, 2acffed (wiring);
   - the third fix of 33d0baf: the folder-setting key;
   - the skip-refresh wiring of 2aaa2b1 in report_processor;
   - be6418c, on Linux only.
8. **README and commit-message claims vs reality** [RAN/READ].
   - tests/README.md is stale. It lists 4 of 18 files. It says "No GUI … is exercised" and "runs fast (<1s) and anywhere". Its example quirk ("Privateer" being mangled) was reversed by 33d0baf.
   - The "81 tests" in 6c1e1c6 is exact [RAN].
   - Two later counts are wrong: d64d3a9 says "140 tests pass" but 132 were collected; 322717f says "21 unit tests" plus "9 smoke tests" but there are 19 and 8, although its total of 127 is right.
   - Several commit messages describe verification harnesses that were never committed.
9. **Test-hygiene issues** [RAN/READ]:
   - The dark-mode tests leave the shared Tk root and global styles dark for later tests.
   - Debounced Setup refreshes fire after their widgets are destroyed: 5 "invalid command name" tracebacks appear in the captured output.
   - One test is vacuous: `test_does_not_import_the_module_to_check_it`.
   - One test's name promises a behaviour its body never exercises: `test_state_change_while_dark_is_reflected_after_restore`.
   - The scroll guard is text-based and can be bypassed (`unbind_all (` with a space).
10. **B has no tests at all** [READ], in any branch or commit. Running A's pure-logic tests against B's code [RAN] gives 7 failures (sanitise abbreviation and IMS trailing "(N)"). `test_file_ops.py` cannot even import, because B has no `scan_excel_files`. A's tests therefore record exactly the behaviours where A and B differ.

---

## 2. What I read, and how

[READ] Every line of these files, with the Read tool, at HEAD bf2f1dd in /root/merge-scratch/A-copy (line counts are exact; every test file ends with a newline):

| File | Lines | First added / last changed (A/main) |
|---|---|---|
| pytest.ini | 6 | 6c1e1c6 23-07-2026 |
| tests/README.md | 34 | 6c1e1c6 23-07-2026 (never updated) |
| tests/conftest.py | 52 | 6c1e1c6; a677670 25-07-2026 |
| tests/test_dark_mode.py | 198 | aeb83cf 26-07-2026; c89c1d4 26-07-2026 |
| tests/test_dependency_installer.py | 115 | a5ad487 26-07-2026 |
| tests/test_excel_com.py | 32 | ba2a753 23-07-2026 |
| tests/test_extractor_panel.py | 137 | a677670 25-07-2026; 2aaa2b1 25-07-2026 |
| tests/test_extractor_standalone.py | 106 | a677670 25-07-2026 |
| tests/test_file_ops.py | 169 | 6c1e1c6; 56aaad9 25-07-2026; 4f84ba9 25-07-2026 |
| tests/test_file_patterns.py | 94 | 6c1e1c6 23-07-2026 |
| tests/test_helpers.py | 181 | 6c1e1c6; 33d0baf 23-07-2026; 4219ffb 25-07-2026 |
| tests/test_main_window_tabs.py | 140 | bfcb24a 26-07-2026 |
| tests/test_optional_pywin32.py | 101 | c463256 26-07-2026 |
| tests/test_processing_handler.py | 36 | 33d0baf 23-07-2026 |
| tests/test_refresh_naming.py | 121 | 2aaa2b1 25-07-2026 |
| tests/test_requirements_check.py | 95 | f363946 26-07-2026; be6418c 26-07-2026 |
| tests/test_scan.py | 80 | 6c1e1c6 23-07-2026 |
| tests/test_scroll_bindings.py | 64 | bf2f1dd 27-07-2026 |
| tests/test_setup_tab_smoke.py | 210 | 322717f; d64d3a9; a677670 (all 25-07-2026); 510f7b4 26-07-2026 |
| tests/test_setup_validation.py | 112 | 322717f 25-07-2026 |
| tests/test_workflow_mode.py | 120 | 510f7b4 26-07-2026 |

[READ] Every test-touching commit in full with `git show`: 6c1e1c6, 33d0baf, ba2a753, 4219ffb, 56aaad9, 4f84ba9, 322717f, d64d3a9, a677670, 2aaa2b1, bfcb24a, 510f7b4, aeb83cf, c463256, f363946, be6418c, a5ad487, c89c1d4, bf2f1dd. I also read the messages and stats of related fixes without tests: 064c634, 725fbaf, 169c7b5, 7347383, 2acffed, c1f425e, b276542, 5c8ce95, c1583d6, 0890022, a603c02, 275d9a4, 12d505f.

[READ] Non-test code, read only as far as needed to explain the tests:
- In full: utils/helpers.py, core/file_parser.py, utils/constants.py, gui/utils/setup_validation.py, gui/utils/dark_mode_manager.py, utils/requirements_check.py, utils/dependency_installer.py, utils/excel_com.py, power_query_extractor/core/refresh_naming.py, gui/utils/cache_manager.py, gui/tabs/extract_tab.py, gui/main_window.py, gui/tabs/setup_tab.py, power_query_extractor/extractor_main.py, launch_extractor.py, main.py, and all `__init__.py` files under gui/, utils/, core/ and power_query_extractor/.
- In part: power_query_extractor/gui/extractor_window.py (lines 1-200, 420-509, 540-849, 960-1063), gui/handlers/processing_handler.py (1-200, 470-530), power_query_extractor/core/report_processor.py (1-175), core/excel_handler.py (1-120), gui/handlers/client_handler.py (18-50, 155-224), gui/handlers/file_handler.py (115-151), power_query_extractor/config/cell_mappings.py (1-30, 245-278, plus a grep of its sheet names).

These other modules belong to other readers. I record only what bears on the tests.

**Side observations for the orchestrator** [RAN]:
- **A-copy has extra bytecode folders.** /root/merge-scratch/A-copy contains 12 untracked `__pycache__` folders, dated 21-10-2026 03:21. They were written before this task started; the names (`*-pytest-9.1.1.pyc`) show an earlier pytest run inside A-copy without `-B`. Apart from them, A-copy is byte-identical to A/main (`git archive A/main` compared with `diff -rq`). The original checkout /home/user/AnnualStatement_oldv1 has no `__pycache__`.
- **`lines_A.txt` undercounts by one line for 37 of A's tracked files.** These files end without a final newline, so `wc -l` misses their last line. For example, power_query_extractor/gui/extractor_window.py really has 1063 lines (line 1063 is `return getattr(panel, name)`), and the three "0-line" `power_query_extractor/{config,core,gui}/__init__.py` files each contain a one-line docstring. Affected files include core/*.py, gui/**/*.py except setup_tab, extract_tab, dark_mode_manager, setup_validation and tooltip, main.py, utils/{__init__,constants,helpers}.py, requirements.txt, README.md and .gitignore. None of the test files is affected.

---

## 3. The runs

### 3.1 Commands (exactly as specified, plus `PYTHONDONTWRITEBYTECODE=1` so no bytecode was written) [RAN]

```
mkdir -p /root/merge-scratch/run/A-T && cp -r /root/merge-scratch/A-copy/. /root/merge-scratch/run/A-T/
# I then deleted the copied stale __pycache__ folders in run/A-T only
cd /root/merge-scratch/run/A-T
PYTHONDONTWRITEBYTECODE=1 xvfb-run -a /root/merge-scratch/venv-A/bin/python -B -m pytest -p no:cacheprovider -rA --color=no   # run 1
PYTHONDONTWRITEBYTECODE=1 /root/merge-scratch/venv-A/bin/python -B -m pytest -p no:cacheprovider -rA --color=no                # run 2 (DISPLAY empty)
```

Environment: Python 3.14.6, pytest 9.1.1, openpyxl 3.1.5, lxml 6.1.3, Tk 9.0, no pywin32, Linux. The full outputs are saved in /root/merge-scratch/run/A-T/run1_xvfb.txt and run2_nodisplay.txt. After the runs, run/A-T contained no `__pycache__` and no `.pytest_cache` [RAN].

### 3.2 Final counts [RAN]

| Run | Collected | Passed | Failed | Skipped | Errors | Time |
|---|---|---|---|---|---|---|
| 1, virtual display (xvfb) | 288 | 287 | 1 | 0 | 0 | 3.56 s (repeats: 3.48, 3.71, 3.71) |
| 2, no display | 288 | 241 | 14 | 33 | 0 | 1.48 s (repeats: 1.38, 1.48) |

The earlier quick run's numbers ("1 failed, 287 passed" and "14 failed, 241 passed, 33 skipped") are **confirmed exactly**. The results were stable across repeats. No "warnings summary" appeared in either run.

Collected per file [RAN]: test_dark_mode 11, test_dependency_installer 13, test_excel_com 3, test_extractor_panel 10, test_extractor_standalone 6, test_file_ops 22, test_file_patterns 21, test_helpers 47, test_main_window_tabs 7, test_optional_pywin32 4, test_processing_handler 2, test_refresh_naming 19, test_requirements_check 15, test_scan 5, test_scroll_bindings 64, test_setup_tab_smoke 13, test_setup_validation 19, test_workflow_mode 7. Total 288.

Slowest steps under xvfb [RAN, `--durations=8`]:
- setup of `test_main_window_starts` (the main-window subprocess): 1.34 s
- standalone extractor subprocess: 0.55 s
- pywin32 probe: 0.30 s
- first SetupTab build: 0.17 s

The four original files alone (file_patterns, helpers, file_ops, scan: 95 tests) took 0.32 s [RAN].

### 3.3 Noise in run 1 that is not a failure [RAN]

The `-rA` output shows 5 "Exception in Tkinter callback … _tkinter.TclError: invalid command name …" tracebacks. They are raised at gui/tabs/setup_tab.py:553 inside `refresh_status` and appear in the "Captured stderr setup" of the 4 `TestScrolling` tests in test_setup_tab_smoke.py.

Cause [READ, HIGH]:
1. Setting a path variable triggers a trace (`_attach_traces`, setup_tab.py:506-510).
2. The trace schedules `refresh_status` 350 ms later (`_schedule_refresh`, :512-519; `REVALIDATE_DELAY_MS`, :135).
3. The tests then call `refresh_status(force=True)` directly. That clears `_revalidate_job` (:533) but does **not cancel** the job already scheduled.
4. The fixture destroys the widgets (test_setup_tab_smoke.py:64).
5. The stale job fires later, during the next test's `root.update()`, against widgets that no longer exist.

Five earlier tests set a source folder and so leave a pending job: lines 80, 87, 96, 110 and 118. In the real app the tab is never destroyed, so the effect there is only a redundant second refresh (MEDIUM).

Captured log lines that are expected behaviour:
- "Could not parse filename: …": core/file_parser.py:131
- "Unknown state …, using fallback": utils/helpers.py:612, :617
- "Error finding Excel files in …/nope": utils/helpers.py:364
- "No Excel files found": core/file_parser.py:160
- "Could not open the extraction tab: scan exploded": gui/handlers/processing_handler.py:174

---

## 4. Every failure and every skip

### 4.1 Run 1 (virtual display): 1 failure

**`tests/test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only`** (test at :65-66)
- Assertion [RAN]: `AssertionError: assert 'C:\\gst\\itc.xltx' == 'itc.xltx'` at tests/test_setup_validation.py:66.
- Why [READ+RAN]:
  - The test feeds Windows-style paths (`ALL_GOOD`, test_setup_validation.py:12-17).
  - `_check_template` returns `Path(value).name` as the message (gui/utils/setup_validation.py:64).
  - On Linux, `Path` is a POSIX path in which `\` is an ordinary character, so `.name` is the whole string.
  - [RAN]: `PurePosixPath(r'C:\gst\itc.xltx').name == 'C:\\gst\\itc.xltx'` but `PureWindowsPath(...).name == 'itc.xltx'`.
- On Windows: **would pass**, HIGH. [RAN-sim]: with `setup_validation.Path` replaced by `PureWindowsPath`, the message is `'itc.xltx'`, and all four suffix tests still pass. The other tests in that file use only string checks (`_short`, :67-72) or the injected `exists` function, so they pass on both systems [RAN].
- Side note [READ]: the suffix check at :62 works on Linux by coincidence, because `PurePosixPath(...).suffix` is still `.xltx` [RAN].

### 4.2 Run 2 (no display): 14 failures

All 13 subprocess-window failures have the same root cause: the probe script cannot create a Tk window without a display. Each probe wraps everything in `try/except` and still prints its `@@RESULT@@` line, so the fixture never reaches its skip branch:
- test_extractor_standalone.py:30-54 with the skip at :75-77;
- test_main_window_tabs.py:24-80 with the skip at :99-101.

The first test in each file then fails on `probe['ok']`, and the others fail with `KeyError` because the result dictionary lacks their keys.

| Test id | Exact error [RAN] | Why here | On Windows |
|---|---|---|---|
| test_extractor_standalone.py::test_standalone_window_starts | `AssertionError: TclError: no display name and no $DISPLAY environment variable` (:82) | `PowerQueryExtractorApp.__init__` → `tk.Tk.__init__` (power_query_extractor/gui/extractor_window.py:1040) needs a display | pass, HIGH (always a display) |
| …::test_owns_the_window_level_concerns | `KeyError: 'title'` (:86) | same | pass, MEDIUM-HIGH (asserts title "Power Query Extractor" and geometry "1000x700", set at extractor_window.py:1042-1043; Windows display scaling could in theory change the reported geometry: not sure, check by running on the PC) |
| …::test_standalone_keeps_its_banner_and_status_bar | `KeyError: 'show_header'` (:92) | same | pass, HIGH |
| …::test_panel_supplies_its_own_two_tabs | `KeyError: 'tabs'` (:97) | same | pass, HIGH |
| …::test_attribute_forwarding_keeps_old_usage_working | `KeyError: 'forwards_client_vars'` (:101) | same | pass, HIGH |
| …::test_unknown_attribute_still_raises | `KeyError: 'unknown_attr_raises'` (:106) | same | pass, HIGH |
| test_main_window_tabs.py::test_main_window_starts | `AssertionError: … File ".../gui/main_window.py", line 48, in __init__  self.root = tk.Tk() … _tkinter.TclError: no display name and no $DISPLAY environment variable` (:106) | `GSTOrganizerApp.__init__` creates `tk.Tk()` at gui/main_window.py:48 | pass, MEDIUM-HIGH. Caveat: on the PC it reads the real `~/.gst_organizer_cache.json`, and may check real (network) paths from it; see §9 |
| …::test_has_four_steps_ending_with_extract | `KeyError: 'tabs'` (:110) | same | pass, HIGH |
| …::test_selecting_any_tab_does_not_raise | `KeyError: 'tab_select_errors'` (:121) | same | pass, HIGH |
| …::test_startup_does_not_scan_the_network | `KeyError: 'scans_at_startup'` (:125) | same | pass, HIGH |
| …::test_step4_prefills_the_target_folder | `KeyError: 'folder_prefilled'` (:129) | same | pass, MEDIUM-HIGH (depends on the folder-sync trace, gui/tabs/extract_tab.py:121-131, behaving with whatever target folder the real cache held first) |
| …::test_step4_suppresses_its_own_chrome | `KeyError: 'panel_header_suppressed'` (:135) | same | pass, HIGH |
| …::test_scanning_on_demand_finds_clients | `KeyError: 'clients_after_scan'` (:140) | same | pass, HIGH |
| test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only | same as §4.1 | Linux path rules | pass, HIGH |

### 4.3 Run 2 (no display): 33 skips

Every skip has the reason **`no display available for Tk`**, raised by the shared session fixture `tk_root` at tests/conftest.py:43-46 when `tk.Tk()` raises `TclError`. A skip in a session-scoped fixture (one created once for the whole run) skips every test that asks for it. All 33 would run on Windows; whether they pass is in the right-hand column.

| File:line (test) | On Windows |
|---|---|
| test_dark_mode.py:65 TestRoundTrip::test_restores_every_colour_exactly | pass, MEDIUM-HIGH (Windows default colour names such as `SystemButtonFace` are captured and replayed as strings, gui/utils/dark_mode_manager.py:199-215, 255-268) |
| :76 TestRoundTrip::test_repeated_toggling_does_not_drift | pass, MEDIUM-HIGH |
| :89 TestRoundTrip::test_coloured_surfaces_are_not_flattened_to_white | pass, HIGH |
| :100 TestRoundTrip::test_button_foreground_is_not_forced_to_white | pass, HIGH |
| :112 TestDarkApplication::test_widgets_actually_go_dark | pass, HIGH |
| :123 TestDarkApplication::test_entry_uses_the_field_colour | pass, HIGH |
| :129 TestDarkApplication::test_widget_created_after_the_toggle_is_captured | pass, HIGH |
| :148 TestTreeTags::test_light_mode_restores_tags | pass, HIGH |
| :162 TestTreeTags::test_widget_without_tag_configure_is_ignored | pass, HIGH |
| :182 TestDynamicColoursSurviveTheToggle::test_state_change_while_dark_is_reflected_after_restore | pass, HIGH. Note: it skips only because it requests `tk_root` (:182) without using it; its body needs no display |
| test_extractor_panel.py:49 TestEmbeddedPanel::test_builds_inside_an_arbitrary_parent | pass, HIGH |
| :54 …::test_finds_clients_from_the_target_folder | pass, HIGH |
| :58 …::test_host_chrome_is_suppressed | pass, HIGH |
| :64 …::test_has_its_own_setup_and_processing_tabs | pass, HIGH |
| :70 …::test_logging_and_progress_work | pass, HIGH |
| :82 TestRefreshStatusUpdates::test_status_starts_as_never | pass, HIGH |
| :86 …::test_status_updates_in_place_after_processing | pass, HIGH |
| :97 …::test_custom_suffix_is_picked_up_from_the_setting | pass, HIGH |
| :114 …::test_legacy_files_still_recognised_after_changing_the_suffix | pass, HIGH |
| :126 …::test_selections_survive_the_status_update | pass, HIGH |
| test_setup_tab_smoke.py:68 test_tab_builds_without_error | pass, HIGH |
| :75 test_scan_button_starts_disabled_when_nothing_selected | pass, HIGH |
| :80 test_scan_button_enables_once_source_is_valid | pass, HIGH |
| :87 test_scan_button_disabled_again_for_a_bad_path | pass, HIGH |
| :96 test_checklist_reflects_states | pass, HIGH |
| :110 test_summary_updates | pass, HIGH |
| :118 test_scan_command_is_wired | pass, HIGH |
| :126 test_length_hint_uses_the_configured_limit | pass, HIGH |
| :154 TestScrolling::test_wheel_over_a_child_widget_scrolls | pass, MEDIUM-HIGH. `_make_scrollable` (:140-145) asserts that the content is taller than the 900x600 window; Windows fonts and scaling change heights, but the tab is long |
| :168 TestScrolling::test_small_delta_still_scrolls | pass, MEDIUM-HIGH |
| :181 TestScrolling::test_wheel_outside_the_tab_is_ignored | pass, MEDIUM-HIGH |
| :195 TestScrolling::test_no_scroll_when_everything_fits | pass, HIGH |
| :205 TestScrolling::test_scrolling_up_at_the_top_stays_put | pass, MEDIUM-HIGH |

Supporting evidence for the "pass on Windows" verdicts [READ]:
- Commit a677670 says "Verified: zero skips across six consecutive runs".
- d64d3a9 and 322717f report passing totals.
- The author ran tests on Windows: c463256 mentions "Tests ran under the latter [python.exe]".

### 4.4 Tests that depend on Windows, on the installed packages or on a display (summary)

- **Display needed (in-process, skip without one):** test_dark_mode (10 of 11; `test_toggle_recomputes_rather_than_replaying` needs none), test_extractor_panel (10), test_setup_tab_smoke (13). Total 33.
- **Display needed (subprocess, FAIL without one):** test_extractor_standalone (6), test_main_window_tabs (7). Total 13.
- **Platform-dependent assertions:**
  - test_setup_validation.py:65-66: needs Windows path parsing; fails on Linux.
  - test_requirements_check.py:52-58 (`test_pywin32_is_required_on_windows`): asserts `essential is (sys.platform == 'win32')`, so it expects different values on each system.
  - test_requirements_check.py:69-71 (`test_openpyxl_is_actually_installed_here`): calls the real `check_requirements()`. On Windows, pywin32 counts as essential (utils/requirements_check.py:17, :57), so this test **fails under any Windows Python without pywin32**, with the misleading message "openpyxl is needed to run the app". [RAN-sim]: with pywin32 marked essential, `missing_essential` returns `['pywin32']`.
  - test_helpers.py:171-172 (`clean_windows_path`): the result is identical on both systems (utils/helpers.py:418-422) [READ].
- **Simulates Windows-without-pywin32 on any machine:** test_optional_pywin32.py blocks `pythoncom`, `win32com`, `win32api`, `win32con` and `win32gui` inside a subprocess (:29-41). It passed here; it should pass on Windows with or without pywin32 (HIGH).
- **No test drives real Excel or COM.** test_excel_com.py uses fake objects only (:9-16) [READ]. The coverage recorder confirms that no COM code path ran (§7).

---

## 5. File-by-file

Each "scenario" line says what the test pins down, which shows the edge cases the author hit in real use. "Guards" names the bug-fix commit a test protects, where one can be identified.

### 5.1 tests/conftest.py (52 lines), shared setup [READ]

- :9 puts the project root first on Python's search path, so tests can `import utils…` and `import core…` without installing anything.
- :14-15 hold the real Excel file signatures (the first bytes of a file on disk): `PK\x03\x04` for .xlsx/.xlsm/.xltx (ZIP-based) and `\xd0\xcf\x11\xe0` for .xls (OLE-based).
- `_write_fake_excel(path, header=PK_SIG, size=2048)` (:18-23) writes the signature followed by zero bytes, up to `size`. "Not a real workbook" (:19-21). It is just enough to pass `validate_excel_file`, which needs a file of at least 1024 bytes (utils/helpers.py:303) with a matching signature (:284).
- The `make_excel` fixture (:26-29) returns that helper.
- The `tk_root` fixture (:32-52) is session-scoped: one Tk root for the whole run. It skips if Tk cannot start (:45-46), hides the window (:47) and destroys it at the end (:49-52).
  - Docstring (:34-41): one root is shared because per-test and per-module roots "made GUI tests skip intermittently". This was added in a677670, which replaced a module-level root from 322717f (test_setup_tab_smoke).
  - [RAN] Consequence of sharing: state leaks between tests (§10.1).
- Fake data: all synthetic (made-up). No real workbook, client or figure.

### 5.2 pytest.ini (6 lines) [READ]

`testpaths = tests`, `python_files = test_*.py`, `python_classes = Test*`, `python_functions = test_*`, `addopts = -q` (quieter output by default).

There is no `pythonpath = .` line. The author's Standard 13 asks for one (rebuild_pipeline_FINAL.md:944-947), because the portable Python never adds the project folder. A relies instead on conftest.py:9 and on each probe's own `sys.path.insert` (for example test_main_window_tabs.py:21). Whether the suite works unchanged under the portable `_runtime` Python is MEDIUM: probably yes because of conftest.py:9, but not run.

### 5.3 test_file_patterns.py (21 tests): filename classification

Exercises `FileParser.parse_filename` (core/file_parser.py:56-134) with `FILE_PATTERNS` (utils/constants.py:13-62). Uses no temp files. Added 6c1e1c6 (23-07-2026); never changed. Module docstring (:1-6): "the most important logic in the app … a silent misclassification misfiles a client's tax data."

[GST RULE] The patterns decide which return or report type a file is: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco or Annual Report. They also decide which part of the name is the client, the state and the period. **TIME-VARYING**: the set of portal returns and reports, and their file names, can change over time. The patterns, as the code has them (constants.py:15, 23, 31, 39, 47, 55):

| Type | Pattern (all `re.IGNORECASE`) | Groups |
|---|---|---|
| GSTR-2B Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | client, state, period |
| IMS Reco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` | client, state, date (8 digits, described as DDMMYYYY at :27) |
| GSTR-3B Export | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | client, state, month |
| Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | client, state, start_month, end_month |
| Sales Reco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | client, state, period |
| Annual Report | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | client, state, year |

Tests:
- `test_parses_each_type` ×8 (:19-39): one happy name per type; `.xls` accepted (`Sales-XYZ-Maharashtra-Apr-Jun.xls`); case-insensitive (`gstr3b-abc-delhi-jan.xlsx`). Each checks `parsed`, `type`, `client` and `state`.
- `test_trailing_duplicate_marker_still_parses` ×6 (:46-60): a trailing " (1)", "(2)" or " (10)" (the copy number Windows adds to duplicate downloads) still parses, and the client is still "ABC". Covers GSTR-2B Reco, IMS Reco ×3, Sales and Annual Report. Not covered: GSTR-3B or Sales Reco with a trailing marker. **Guards 5c8ce95** (23-07-2026, IMS trailing-suffix fix; named in the comment at :44-45). [RAN] M03 (removing the optional `(N)`) → 4 failures.
- `test_ims_suffix_keeps_clean_date` (:63-67): the "(N)" must not leak into the IMS 8-digit date (metadata `date == "01012024"`). [GST RULE] tax-period metadata. Guards 5c8ce95.
- `test_gstr3b_leading_prefix_parses` (:71-75): a leading "(1) " before GSTR3B parses. The comment calls this a portal-export convention, handled only for GSTR-3B. This behaviour dates from the initial commit edac6a3 [READ, `git log -S`]. [RAN] M04 → caught.
- `test_unrecognized_names_do_not_parse` ×5 (:80-94): random names, too few or too many dash groups, and the wrong extension give `parsed False`, `type "Unknown"`, `client ""`.

Untested behaviour noticed [RAN probe]:
- For every type except IMS, the trailing "(N)" ends up inside the period/month/year metadata: `period 'Apr24 (1)'`, `end_month 'Jun (2)'`, `year '2024 (1)'`, `month 'Jan (1)'`, `period 'Q1 (1)'`. [GST RULE] tax-period metadata, described only.
- A leading "(1) " on any type other than GSTR-3B makes the file unrecognised (for example `(1) Sales-…` and `(1) ImsReco-…` give `parsed False`).
- A one-letter client parses, but a warning is attached (core/file_parser.py:109-119).

These outputs were not checked for correctness.

### 5.4 test_helpers.py (47 tests): pure helpers in utils/helpers.py

Added 6c1e1c6; changed 33d0baf and 4219ffb. Imports at :10-19.

**TestSanitizeFilename** (`sanitize_filename`, helpers.py:68-126; builds file and folder names):
- :25-28: "ABC Private Limited" inside a GSTR-2B name becomes `GSTR-2B-Reco-ABC_Pvt_Ltd-Delhi-Apr24.xlsx`, with spaces turned into underscores.
- :30-31: invalid characters `<>:*` are removed, so the name becomes `Report.xlsx`.
- :33-34: repeated spaces collapse and are trimmed, giving `multiple_spaces`.
- :36-37: an empty name becomes `unnamed` (helpers.py:72-73).
- :39-41: the duplicate marker "(1)" survives.
- :43-45: `test_abbreviation_uses_word_boundaries`: "Privateer Unlimited" is NOT abbreviated. **Guards 33d0baf**, fix 1, which reversed the original 6c1e1c6 test that pinned `Pvter_Unlimited` [READ diff].
- :47-48: case-insensitive: "ABC PRIVATE LIMITED" becomes `ABC_Pvt_Ltd`. Guards 33d0baf.
- :50-67: the abbreviation happens whatever the separator: spaces, underscores, hyphens or repeated whitespace (×4). **Guards 4219ffb** (the `_` underscore regression introduced by 33d0baf). The test's docstring cites Excel's ~218-character path limit.
- :69-74: NOTE: a run-together "ABCPrivateLimited" is left alone. Pinned as known behaviour. Added 4219ffb.
- :76-79: 250 x's + ".xlsx" is cut to 185 x's + ".xlsx" (the limit is 200 − extension − 10, helpers.py:119-120).
- [RAN] M01 (substring replace) → 3 failures. M02 (no underscore normalisation) → 1 failure.

**TestGetStateCode** (`get_state_code`, helpers.py:592-618; `STATE_CODE_MAPPING`, constants.py:251-292). [GST RULE] state-code logic. **TIME-VARYING**: the list of states and union territories, and their abbreviations, can change. Described only:
- ×5 (:85-93): Delhi → DL; delhi → DL; "  Maharashtra " → MH (trimmed and lower-cased first, :601); Tamil Nadu → TN; West Bengal → WB.
- :95-97: NOTE: an unknown two-word name takes the initials of its first two words. "Unknown Place" gives "UP", which the comment says collides with Uttar Pradesh (helpers.py:609-613).
- :99-100: an unknown single word takes its first three letters upper-cased: "Xyz" gives "XYZ" (:615-618).
- :102-103: an empty string returns empty (:594-595).
- Only 5 of the 40 mapping entries are pinned. [RAN] M23 (changing one unpinned entry) → no test noticed.
- [RAN probe] Untested quirk: the single-word fallback slices the *unstripped* input (`state_name[:3]`, :616), so " Xyz" gives " XY".
- These are 2-letter abbreviations, not the numeric GSTIN state codes. No GSTIN logic exists in the tested code [READ].

**TestCreateClientStateKey** (`create_client_state_key`, helpers.py:620-644). [GST RULE], because it embeds the state code:
- :109-112: "ABC Private Limited" + Delhi gives `ABC Pvt Ltd-DL`.
- :114-115: "XYZ" + Tamil Nadu gives `XYZ-TN`.
- :117-122: a long name is capped at 35 characters and ends with `-MH`.
- Stale comment at :110-111: "distinct from sanitize_filename's underscore/substring behaviour". `sanitize_filename` has used word boundaries since 33d0baf [READ].
- The function is used for folder naming at core/file_organizer.py:83. That link is untested.

**TestValidateClientName** ×6 (:127-137; helpers.py:428-449): empty → "Client name cannot be empty"; "A" → "Client name too short"; "ABC" valid; "Valid Name" valid; 101 x's → "Client name too long"; "AB<C" → "Client name contains invalid characters". `parse_filename` also runs this function on state names (file_parser.py:111) [READ].

**TestFormatSize** ×7 (:142-153; helpers.py:379-392): 0 → "0.0 B"; 1023 → "1023.0 B"; 1024 → "1.0 KB"; 1536 → "1.5 KB"; 1 MB; 1 GB; −5 → "Unknown".

**TestFormatDuration** ×6 (:156-166; helpers.py:512-523): 0s, 30s, 59s, "1m 30s", "59m 59s", and 3661 → "1h 1m". NOTE: the hours format drops the seconds.

**Path helpers**: `clean_windows_path("a/b/c") == r"a\b\c"` (:171-172; helpers.py:418-422). `extract_filename_without_extension` ×3 (:175-181; helpers.py:553-555): only the last extension is removed.

[RAN-sim] I checked every `sanitize_filename` input with Windows path parsing (`PureWindowsPath.stem/suffix`). The results are the same, so the whole file passes on Windows (HIGH).

### 5.5 test_file_ops.py (22 tests): file discovery and validation against real temp files

Added 6c1e1c6; extended 56aaad9 and 4f84ba9. Uses `make_excel` and `tmp_path`.

**TestValidateExcelFile** (`validate_excel_file`, helpers.py:240-260, built from `_passes_metadata_checks` :290-306 and `_has_excel_signature` :275-287):
- Accepted: valid xlsx; valid xls (OLE signature); a `.xltx` template.
- Rejected:
  - a 500-byte file (helpers.py:303);
  - a wrong signature;
  - a `.txt` file;
  - a missing file;
  - a folder named `adir.xlsx`.
- [RAN] M25 (size threshold 1 byte) → 3 failures.

**TestFindExcelFiles** (`find_excel_files`, helpers.py:368-373):
- :47-64: in a mixed folder only valid `.xlsx/.xls/.xlsm` files are found. "UPPER.XLSX" counts (case-insensitive, helpers.py:327-331). A `.xltx` file is valid but not listed. Tiny, fake, `.txt` and folder entries are excluded.
- :66-70: results are sorted.
- :72-73: an empty folder gives an empty list.

**TestScanExcelFiles** (`scan_excel_files`, helpers.py:309-365). Added 56aaad9 and 4f84ba9:
- Results are identical to `find_excel_files`.
- The returned stat (the file's size and date details, cached from the directory listing) is real (`st_size == 4096`).
- A folder named like an .xlsx file is rejected using the cached stat.
- An empty folder gives an empty list.
- `test_parallel_branch_matches_sequential` (:106-128): 30 valid plus 3 invalid files exercise the thread pool (16 workers, used when there are 4 or more files, helpers.py:271-272, 351-357). The result must be the same files, in sorted order. **Guards 4f84ba9**. [RAN] M26 (sorting removed) → 2 failures.
- A missing folder gives `[]`. The error is logged and swallowed (:363-365).

**TestCachedStatEquivalence** (:134-149): `validate_excel_file` and `get_file_info` give the same result with and without a pre-supplied stat. **Guards 56aaad9**.

**TestGetFileInfo** (`get_file_info`, helpers.py:129-166): metadata of a real file; a folder reports `is_dir True`; a missing file returns `{"error": "File not found"}`.

Fake data: names such as `GSTR3B-ABC-DL-Jan.xlsx` and `GSTR3B-Client{00..29}-Delhi-Jan.xlsx` (:115) are synthetic.

On Windows: pass, HIGH. No two names in one folder differ only by letter case, and Windows sorts paths without regard to case, but all names here share the same case.

### 5.6 test_scan.py (5 tests): end-to-end `FileParser.scan_folder`

Exercises core/file_parser.py:136-177: discover, then validate, parse, group by client and check completeness. Added 6c1e1c6; never changed.

- :18-38: `GSTR3B-ABC-Delhi-Jan.xlsx` + `Sales-ABC-Delhi-Apr-Jun.xlsx` give one client, key `ABC-DL` (client "ABC", state "Delhi", code "DL"; file_parser.py:229-236), two file types and file_count 2. Missing = the other 4 expected types; status "Missing 4 files" (:281-283).
- :41-49: a valid-signature file whose name matches no pattern goes to `variations` (:212-220).
- :52-58: a bad-signature `.xlsx` is skipped entirely and does not appear in `scanned`.
- :61-64: an empty folder gives no clients. This is the early return at :159-161.
- :67-80: one file of each of the six types gives `missing_files == []` and status "Complete" (:278-280).

[GST RULE] **TIME-VARYING**: `EXPECTED_FILE_TYPES` (constants.py:68-75) lists the six inputs a client needs to count as "Complete". The tests pin the count through the set difference at :36 and `len(missing) == 4` at :37.

Not tested:
- the "Has duplicates" branch (:284-286);
- [GST RULE] the rule that several GSTR-3B files per client are allowed while duplicates of other types are flagged (:262-271; comment "monthly filings"). [RAN] M24 (that exception removed) → no test noticed;
- `get_statistics` (:308-336) and `validate_results` (:338-371). The latter treats 'Annual Report' and 'GSTR-3B Export' as "critical types" at :359 [GST RULE];
- the warning branch (:113-119);
- the progress callback.

### 5.7 test_excel_com.py (3 tests): `find_open_workbook`

Exercises utils/excel_com.py:19-29. Fakes stand in for Excel's workbook list (:9-16). Added ba2a753 (23-07-2026), which shared the "already open in Excel?" loop between `ExcelHandler._safe_open_workbook` (core/excel_handler.py:324) and `ReportProcessor._safe_open_workbook` (report_processor.py:687).

Scenarios: the matching open workbook is returned; `None` when the file is not open; `None` when nothing is open.

Comparison is `Path(wb.FullName).absolute() == abs_path` (excel_com.py:27). Not tested: the same file reached through a mapped drive letter versus a UNC (`\\server\share`) path, and long `\\?\` paths. MEDIUM relevance, not verified.

### 5.8 test_processing_handler.py (2 tests): Step 3 re-entry guard

Exercises `ProcessingHandler.start_processing` (gui/handlers/processing_handler.py:69-80). Uses a fake app with no Tk. Added **33d0baf**, fix 2.

- :24-29: while `is_processing` is True, a click returns before validation.
- :32-36: when idle, validation runs.

Why the guard exists [READ]: the Start buttons are still wired with `.bind('<Button-1>')` (gui/tabs/processing_tab.py:74; gui/tabs/validation_tab.py:142), which fires even when the button is disabled. [RAN] M05 → caught.

Importing the module pulls in the whole GUI (§10.4).

### 5.9 test_setup_validation.py (19 tests): Setup checklist logic

Exercises gui/utils/setup_validation.py (all 10 functions ran). Fake existence checker `make_exists` (:7-9); Windows-style sample paths (:12-17). Added **322717f** (25-07-2026, live checklist).

**TestEvaluateSetup:**
- All four inputs present gives all OK, in the order source, itc, sales, target.
- Empty → MISSING "Not selected yet".
- Whitespace only → MISSING (setup_validation.py:49, :58).
- A folder that does not exist → INVALID "Folder not found".
- A template that does not exist → INVALID "File not found".
- A `.txt` template → INVALID "Not an Excel template".
- Accepted template suffixes `.xlsx`, `.xltx`, `.xltm`, `.xlsm` (×4; `TEMPLATE_SUFFIXES`, :21).
- An OK template shows only its file name. This is the Linux-only failure (§4.1).
- A long UNC path (a `\\server\share\…` network path) is shortened to keep its end: it starts with "...", ends with "level19" and is at most 46 characters (`_short`, :67-72).

**TestGating:**
- Scanning needs only the source folder; processing needs all four (:88-97).
- `blocking_labels` for scanning gives `['Source folder']`; for processing it lists the pending items.

**TestSummary:** "Ready - all 4 items set" and "2 of 4 items set" (:109-115).

**`test_default_exists_checker_handles_bad_input`** (:108-112): `""`, a missing Windows path and an embedded NUL character all give False (the `ValueError` is caught at :43).

Docstring claim checks [READ]:
- The module docstring says "no Tkinter involved". Importing `gui.utils.setup_validation` runs `gui/__init__.py:7`, which imports the whole main window. [RAN] 222 project and library modules load, including tkinter, all tabs and report_processor. No window is created, so there is no display dependency [RAN: passes without a display].
- `test_scan_needs_only_the_source_folder` says it "mirrors FileHandler.validate_scan_inputs". No test checks that the two agree, and they differ [READ]:
  - `FileHandler.validate_processing_inputs` (gui/handlers/file_handler.py:132-151) does not check template suffixes, and checks only that the target folder is not empty, not that it exists;
  - the checklist's state does not gate the Start button: `_update_actions` (setup_tab.py:585-608) sets only Scan and Re-scan;
  - there are three different "Excel/template extension" sets: setup_validation.py:21 `{.xlsx,.xltx,.xltm,.xlsm}`, utils/helpers.py:299 `{.xlsx,.xls,.xlsm,.xltx,.xltm}` and constants.py:197 `TEMPLATE_EXTENSIONS {.xlsx,.xltx,.xltm}`. The last is never used [READ grep].

### 5.10 test_setup_tab_smoke.py (13 tests): builds the real Setup tab in Tk

Exercises gui/tabs/setup_tab.py. 20 of its 22 functions ran; `_on_workflow_change` (:347-351) and the inner `on_length_change` (:434-438) did not. Uses `FakeApp` (:21-44), which has exactly the variables and callbacks SetupTab touches; it has no `dark_mode`, so `_is_dark()` returns False through its exception path (:204-209). The fixture `built` (:47-65) lays the tab out in a 900x600 root, because the canvas must have a real size for scrolling. History: 322717f added it (8 tests), d64d3a9 (scroll tests plus layout), a677670 (shared root), 510f7b4 (`workflow_mode` added to `FakeApp`).

- :68-72: the tab builds; `scan_btn` exists; `client_name_check` exists ("kept for compatibility").
- :75-77: Scan is disabled when nothing is chosen.
- :80-84: Scan is enabled once the source folder is valid. `refresh_status(force=True)` bypasses the 350 ms delay.
- :87-93: Scan is disabled again when the path becomes invalid.
- :96-107: the checklist icons show ✓ for source and ITC, ○ for sales, and the target reads "Not selected yet". The ITC template is a 1-byte file `itc.xltx`; the template check looks only at existence and suffix.
- :110-115: the readiness label equals `summary_line(evaluate_setup(...))`.
- :118-123: `scan_btn.invoke()` reaches `app.scan_files`, proving it is wired through `command=` (fix carried in 322717f).
- :126-130: the length hint shows the configured limit ("50"). This fixes the hint that used to hard-code "10 chars" (322717f).

**TestScrolling** (`_on_mousewheel` :220-248, `_is_in_scroll_area` :211-218). Each test calls the handler directly with a `FakeWheel` event (:133-137). **Guards d64d3a9**:
- :154-166: the wheel over a child label scrolls. This was the old `<Leave>`-on-children bug. [RAN] M11 → caught.
- :168-179: six deltas of −40 add up and scroll. This was the old `int(delta/120)` truncation. [RAN] M10 → caught.
- :181-193: the wheel over a widget outside the tab is ignored and returns None.
- :195-203: when nothing needs scrolling, the handler returns None, so the event is not swallowed.
- :205-210: wheel-up at the top stays at 0.0.

Not tested: real delivery of `<MouseWheel>` events by the operating system. Windows sends them to the focus or pointer window depending on the Tk version; this machine runs Tk 9.0, while the author's Tk version is unknown. Also untested: the debounce timer itself, tooltips showing and hiding (only `Tooltip.__init__` and `set_text` ran), and the workflow and processing-mode radio buttons.

### 5.11 test_dark_mode.py (11 tests): dark mode

Exercises gui/utils/dark_mode_manager.py (all 14 functions ran). `scene` fixture (:21-47): a container of coloured Frames, Labels, a Button, an Entry and a Checkbutton, using the app's real colours. `snapshot` (:50-61) records background and foreground. Added **aeb83cf** (26-07-2026, dark-mode rebuild); extended c89c1d4.

**TestRoundTrip:**
- dark then light restores every colour exactly;
- three round trips do not drift;
- coloured frames are not flattened to "white" (the old bug);
- the button's foreground stays "black" and is not forced to white.

**TestDarkApplication:**
- coloured labels really go dark. Before, only labels whose background was in a four-colour whitelist were darkened;
- an Entry gets the field colour `#3b3b3b`;
- a widget created after the toggle is still restored to its own colour (`#123456`).

**TestTreeTags** (`update_tree_tags`, :271-294):
- light mode restores the tag colours (in dark mode, the 'complete' tag is `#1a3d1a`);
- a widget without `tag_configure` is ignored. This fixed an `else` attached to the wrong `if`.

[RAN] Swapping in the pre-aeb83cf `dark_mode_manager.py` (from `aeb83cf^`) makes **8 of 11** tests fail. The commit said "fails five of them", but it meant a narrower mutation, so this is not a contradiction.

**TestDynamicColoursSurviveTheToggle** (added **c89c1d4**):
- `test_state_change_while_dark_is_reflected_after_restore` (:182-188): despite its name, it only asserts two constants, `STATE_STYLE[OK][2] == '#E8F5E8'` and `STATE_STYLE[INVALID][2] == '#FDECEA'` (gui/tabs/setup_tab.py:30, :32). It requests `tk_root` and `tmp_path` but uses neither, so it never exercises a state change in dark mode [READ, HIGH].
- `test_toggle_recomputes_rather_than_replaying` (:190-198): checks that the *source text* of `GSTOrganizerApp.toggle_dark_mode` contains the word `refresh_status`. [RAN] M08 deleted the actual call `setup_tab.refresh_status(force=True)` (gui/main_window.py:191), and no test noticed, because `hasattr(setup_tab, 'refresh_status')` on :189 still contains the word. [RAN] The coverage recorder shows `toggle_dark_mode` itself is **never executed** by any test.

Not tested, with [RAN] mutation evidence:
- **725fbaf**: an empty ttk option is recorded so the client-list background can be restored (M13 survived), and the "padding [20, 10]" repaint hack is gone (M14 survived). No test checks ttk style values after a round trip.
- **169c7b5**: the dark variants of the checklist colours (M15 survived).
- The dark-mode toggle through the real title-bar button and the main window.

### 5.12 test_extractor_panel.py (10 tests): the Step 4 panel embedded in a frame

Exercises power_query_extractor/gui/extractor_window.py `ExtractorPanel`: constructor, `scan_folder` (:611-668), `display_clients` (:670-757), `refresh_status_display` (:759-782), `get_refresh_status`, `_configured_suffix`, `format_refresh_status`, `log_message`/`_ui`/`_append_log`, `update_progress`/`_apply_progress`, and `_latest_refresh_times`/`_classify_version_reports` (:20-82).

Fixture `output_tree` (:24-31): a fake organiser output folder, `Annual Statement-250726 1000/ABC-DL/Version-250726 1000/` containing `ITC_Report_ABC_DL.xlsx` and `Sales_Report_ABC_DL.xlsx` (fake bytes). Added a677670; extended 2aaa2b1.

**TestEmbeddedPanel:**
- the panel builds inside any parent;
- it finds client `ABC-DL` from the target folder (`auto_load` defaults to True, so it scans at construction, :135-137);
- the header and status bar are suppressed when embedded;
- it has its own sub-tabs "Setup" and "Processing";
- logging and progress work once the scheduled calls are drained with `root.update()`.

**TestRefreshStatusUpdates.** The class docstring says "the 'last refreshed' column never updated after a run", so it **guards 2acffed**:
- the status starts as "Never refreshed";
- after `_Refreshed_` file appears, `refresh_status_display()` shows "ITC: …";
- **guards 2aaa2b1**: a custom suffix `_Updated_{timestamp}` is recognised from the setting;
- files with the old `_Refreshed_` suffix are still recognised after the setting changes;
- the user's ITC/Sales ticks survive the update, because the rows are updated in place and not re-scanned.

[RAN] M07 → 1 failure in this file. M19 (the log written directly instead of being scheduled) → no failure: the test drains the event loop before checking, so the c1f425e "schedule onto the UI thread" property is not pinned. M20 (`_reset_after_processing` no longer refreshing) → survived: the tests call `refresh_status_display()` directly, and `_reset_after_processing` (:977-986) never runs, so the 2acffed *wiring* is untested.

Not tested:
- the per-client tick selection helpers (:810-836);
- processing (`start_processing`, `_process_clients_thread`);
- `load_from_cache` (:587-599);
- the skip-refresh toggle (:483-492), which changes the module-level `EXTRACTION_OPTIONS` dictionary;
- the client-list wheel handler (:450-478; it never ran);
- picking the newest version folder by modification time when there are several (c1583d6; M18 survived). Every fixture has a single version folder.

### 5.13 test_extractor_standalone.py (6 tests): the standalone extractor window, in a subprocess

The probe (:21-56) builds `PowerQueryExtractorApp(target)` (extractor_window.py:1032-1063) in a separate Python, reads its state, destroys it and prints JSON. Module-scoped fixture (:59-78) with the same fake output tree; subprocess timeout 180 s. Added a677670.

Scenarios:
- the window starts;
- its title is "Power Query Extractor" and its size "1000x700";
- its `panel` is an `ExtractorPanel`;
- banner and status bar are shown when standalone;
- the two sub-tabs exist;
- attribute forwarding works (`app.client_vars` and `app.log_message` reach the panel, :1051-1063);
- an unknown attribute still raises `AttributeError`.

Claim check [READ, HIGH]:
- The docstring (:6-8) and a677670's message say this "exercises the real startup path, the same one launch_extractor.py uses". It does not. The real path is launch_extractor.py:12-18 → `power_query_extractor/extractor_main.py` (whose module-level `logging.basicConfig` writes `pq_extractor.log`, :12-19) → `PowerQueryExtractor()` → `load_cached_folder` (:36-63) → `PowerQueryExtractorApp(...).mainloop()` (:65-69). The probe skips the launcher, extractor_main, the cache lookup and `mainloop`. The coverage recorder shows that neither file was imported [RAN].
- The docstring says the tests "skip" if the window cannot start. Without a display they fail (§4.2).
- [READ] Untested side effect, LOW: `PowerQueryExtractorApp.__getattr__` replaces tkinter's own `Tk.__getattr__`, which hands unknown attributes to the Tcl interpreter (tkinter/__init__.py:2569-2571 in this Python). Such lookups now go to the panel instead.

### 5.14 test_main_window_tabs.py (7 tests): the real main window, in a subprocess

The probe (:18-82) wraps `ExtractorPanel.scan_folder` to count scans. It then builds `GSTOrganizerApp()`, patches `save_cache` to do nothing (but only *after* construction, :38), sets the target folder, selects every tab with a `report_callback_exception` hook, scans on demand and prints JSON. Subprocess timeout 300 s. Added **bfcb24a** (26-07-2026, Step 4 tab).

Scenarios:
- the window starts;
- it has 4 tabs: Setup, Validation, Processing, Extract;
- **guards bfcb24a's crash fix**: selecting any tab raises nothing. `on_tab_changed` used to index a 3-name list; the fix is at gui/main_window.py:331-342. [RAN] M06 → caught. The hook is needed because Tk swallows callback errors (:53-59);
- startup performs no scan, i.e. no network access at launch (`auto_load=False`, extract_tab.py:74-80);
- Step 4 is pre-filled with the target folder and its client list is still empty;
- Step 4's banner and status bar are suppressed;
- scanning on demand finds `['ABC-DL']`.

Not tested:
- [RAN] M28: the rule that Step 4 never overwrites a folder the user chose there (extract_tab.py:127-128) survived mutation;
- the "pywin32 unavailable" notice (extract_tab.py:84-105; never ran).

It writes outside the temp folder (§9).

### 5.15 test_optional_pywin32.py (4 tests): the app must start without pywin32

The probe (:24-67) runs in a subprocess. It replaces Python's import function so that the 6 win32 modules raise `ImportError` (:29-41). It then imports report_processor, calls `ReportProcessor().process_client({'name': 'ABC-DL'})`, imports `gui.tabs.extract_tab` and imports `gui.main_window`. Unlike the other probes, the fixture uses `pytest.fail` (not skip) when there is no result (:78-80). Added **c463256** (26-07-2026).

Scenarios:
- the probe ran;
- `WIN32_AVAILABLE` is False (report_processor.py:20-36);
- extract_tab and main_window import fine;
- `process_client` refuses clearly: `success False` and the error mentions "pywin32" (report_processor.py:66-77).

[RAN] M09 (one unconditional `import pythoncom`) → all 4 fail. This matches the commit's claim "fails all four".

It does not build the window, so the Step 4 "unavailable" card is never shown in a test.

### 5.16 test_requirements_check.py (15 tests): the startup dependency check

Exercises utils/requirements_check.py (all 7 functions ran). Sample requirements at :13-17: `sys` (present), `no_such_module_xyz` (absent, essential) and `no_such_module_abc` (absent, optional). Added **f363946**; changed **be6418c**.

**TestDetection:**
- an installed module is detected;
- a missing module is detected;
- `test_does_not_import_the_module_to_check_it` (:27-32) is **vacuous** (it cannot fail). A module that does not exist can never appear in `sys.modules`, whatever method the check uses. [RAN] M27 (`find_spec` replaced by a real `import_module`) → all 15 tests still pass.

**TestSeverity:** essential and optional are separated; nothing is missing when everything is present.

**TestRealRequirements:**
- openpyxl is essential;
- `test_pywin32_is_required_on_windows` asserts `essential is (sys.platform == 'win32')`. **Guards be6418c** on Windows only. [RAN] M22 (`essential=False`) passes on Linux, but would be caught on Windows (HIGH);
- pywin32's consequence text mentions "Power Query";
- pywin32 is checked through `pythoncom`;
- `test_openpyxl_is_actually_installed_here` really checks *every* essential package (§4.4).

**TestReport:**
- the report names `pip install optional-pkg` and what breaks;
- it shows "Running:" and "'py' and 'python' can be different" (requirements_check.py:125-131);
- a clean report has no ⚠️ or ❌;
- the interpreter hint includes `sys.executable`.

`main.py` (the code that actually uses this at startup, main.py:57-127) is never run by a test [RAN].

### 5.17 test_dependency_installer.py (13 tests): the offer to install missing packages

Exercises utils/dependency_installer.py: `build_pip_command`, `install`, `describe_missing` and `offer_install`'s first lines; `restart_application` never ran. `subprocess.run` is monkeypatched, so **no real pip runs** [READ]. Added **a5ad487**.

- **TestPipCommand**: the command uses the running interpreter, `[sys.executable, '-m', 'pip', 'install', …]` (the py-vs-python fix); named packages come last; `-r reqs.txt` is used when given; both together work.
- **TestInstall**: success with output; failure reports the output (a fake "could not reach proxy"); a timeout is reported, not raised ("did not finish"; `INSTALL_TIMEOUT_SECONDS = 600`, :23); a missing pip is reported ("Could not run pip"); progress lines reach the callback.
- **TestDescribeMissing**: names the package and purpose; includes the consequence; handles several.
- **TestNeverBreaksTheApp**: with tkinter imports blocked, `offer_install` returns False (:95-99).

The dialog path (the Yes/No question, the streaming window and the error box, :101-158) is untested.

### 5.18 test_refresh_naming.py (19 tests): recognising refreshed report files

Exercises power_query_extractor/core/refresh_naming.py (all 3 functions) and `_latest_refresh_times` (extractor_window.py:45-82). Added **2aaa2b1** (25-07-2026). That commit says "Tests: 20 new"; my count is 21 (19 here plus 2 in test_extractor_panel) [RAN].

- **TestRefreshMarkers**: the default `_Refreshed_{timestamp}` gives `['_refreshed_']`; `_Updated_{timestamp}` gives `['_updated_']`; `_v{timestamp}_final` gives `['_v', '_final']`; a bare `{timestamp}` falls back to the legacy marker; `None` or `""` falls back to the legacy marker.
- **TestIsRefreshedName**:
  - default suffix recognised;
  - custom suffix recognised (the reported bug);
  - legacy files still match after the setting changes;
  - the original report is not counted as refreshed;
  - matching is case-insensitive;
  - `.txt` is rejected;
  - for a split pattern both literal parts are required.
  - Note: only `.xlsx` counts (refresh_naming.py:50-51).
- **TestIsRefreshedCopyOf**: a match requires the original's name as a prefix; a different report and the original itself are rejected.
- **TestStatusDetectionUsesTheConfiguredSuffix** (end to end through `_latest_refresh_times`): a custom suffix gives a time, not "Never"; the same folder read with the default suffix finds nothing (the old behaviour); an original with no refreshed copy gives `(None, None)`.

[RAN] M07 → 5 failures here. [RAN] M21 survived: hard-coding `'_Refreshed_'` in report_processor's skip-refresh lookup (report_processor.py:142-146) went unnoticed, so the second caller named in 2aaa2b1 is untested.

[READ, for the report_processor reader; not verified here, MEDIUM] `_process_report` takes `report_files[0]` from `glob(f"{report_type}_Report_*.xlsx")` (report_processor.py:112-125). That pattern also matches refreshed copies (`…_Refreshed_…xlsx`). No test covers which file is chosen.

### 5.19 test_workflow_mode.py (7 tests): organise-only versus the full pipeline

Exercises `ProcessingHandler._maybe_continue_to_extraction` (processing_handler.py:136-159), `_open_extraction_tab` (:161-175) and `WORKFLOW_MODES` (constants.py:163-172). Uses fakes only; `FakeRoot.after` runs immediately (:46-49). Added **510f7b4**.

- the two modes `organize` and `full` exist and both have a name and description;
- `full` selects tab index 3, scans once and copies the target folder `C:\out`;
- `full` logs "press Process", meaning it does not start the refresh. This checks a log message only: a call that started the refresh after the log line would also pass (MEDIUM);
- `organize` does not switch tabs or scan;
- `organize` still mentions "Step 4";
- with no extract tab, nothing is raised and the tab is not switched;
- a scan that raises logs "Could not open Step 4" instead of failing processing.

The fakes cannot show thread problems. The real call happens on the worker thread (processing_handler.py:519) and reads a Tk variable there (:145) [READ].

### 5.20 test_scroll_bindings.py (64 tests): source-text guard against scroll bindings that destroy each other

Collects every `.py` under gui/ and power_query_extractor/: 31 files [RAN] (:16-19).

- `test_sources_were_found`: the list is not empty.
- `test_no_unbind_all` ×31: no non-comment line contains `unbind_all(`.
- `test_bind_all_is_additive` ×31: every non-comment line with `bind_all(` also contains `add='+'` or `add="+"`.
- `test_wheel_handlers_check_the_pointer_is_over_them`: setup_tab.py and extractor_window.py, if they use `bind_all`, define a function matching `def _?(is_)?in_scroll_area`. That holds for setup_tab.py:211 and extractor_window.py:450.

Added **bf2f1dd** (27-07-2026). Guards bf2f1dd (extractor `unbind_all`) and d64d3a9 (setup tab). [RAN] M12a (add `canvas.unbind_all('<MouseWheel>')`) → caught.

Limits (they are text checks, not behaviour checks):
- [RAN] M12b: `unbind_all (` with a space passed all 64 tests;
- a legitimate multi-line `bind_all(` call with `add='+'` on the next line would be falsely flagged (reasoned from :42-47; M12c shows that the first line alone is judged);
- the regex test only checks that a function with that name exists, not that it is used;
- the real wheel handler in extractor_window.py (:458-476) never ran in any test [RAN].

The current `bind_all` uses are setup_tab.py:201 and extractor_window.py:478, both with `add='+'` [READ].

---

## 6. Mutation checks [RAN]

Harness: /root/merge-scratch/run/A-T/mutate.py, with results in /root/merge-scratch/run/A-T/mut_results_all.jsonl. For each mutation it:
1. copies A-copy (minus `__pycache__`) into run/A-T/mut/<id>/;
2. applies one exact text replacement (the old text must occur exactly once);
3. runs the listed tests, or the whole suite under xvfb, with HOME and TMPDIR pointing into the copy.

Whole-suite baseline (M00): 1 failed, 287 passed, the known Linux-only failure. "SURVIVED" below means no failure beyond that baseline.

| Id | Re-introduced bug (file) | Guards | Result |
|---|---|---|---|
| M01 | substring Private/Limited replace (utils/helpers.py:92-93) | 33d0baf | CAUGHT (3) |
| M02 | no `_` normalisation (helpers.py:91) | 4219ffb | CAUGHT (1) |
| M03 | IMS pattern without "(N)" (constants.py:23) | 5c8ce95 | CAUGHT (4) |
| M04 | GSTR3B without leading "(N)" (constants.py:31) | edac6a3 behaviour | CAUGHT (1) |
| M05 | re-entry guard removed (processing_handler.py:75-76) | 33d0baf | CAUGHT (1) |
| M06 | hard-coded 3-name tab list (main_window.py:337) | bfcb24a | CAUGHT (1) |
| M07 | only the legacy marker (refresh_naming.py:52) | 2aaa2b1 | CAUGHT (6) |
| M08 | `refresh_status` call deleted (main_window.py:191) | c89c1d4 | **SURVIVED** |
| M09 | unconditional `import pythoncom` (report_processor.py:20) | c463256 | CAUGHT (4) |
| M10 | `int(delta/120)` truncation (setup_tab.py:243-247) | d64d3a9 | CAUGHT (1) |
| M11 | canvas-only scroll test, children ignored (setup_tab.py:213-218) | d64d3a9 | CAUGHT (2) |
| M12a | `unbind_all('<MouseWheel>')` added (extractor_window.py:478) | bf2f1dd | CAUGHT (2) |
| M12b | `unbind_all ('<MouseWheel>')` with a space | bf2f1dd | **SURVIVED** |
| M12c | multi-line `bind_all` without add on line 1 | bf2f1dd | CAUGHT (1) |
| M13 | empty ttk values not recorded (dark_mode_manager.py:101) | 725fbaf | **SURVIVED** |
| M14 | `TNotebook.Tab padding=[20,10]` hack back (main_window.py:206-207) | 725fbaf | **SURVIVED** |
| M15 | no dark checklist colours (setup_tab.py:543) | 169c7b5 | **SURVIVED** |
| M16 | folder-setting key by state NAME (client_handler.py:181) | 33d0baf fix 3 | **SURVIVED** |
| M17 | long category folder 'Other ITC related files ({client})' (constants.py:93) | 7347383 | **SURVIVED** |
| M18 | newest version folder chosen by name (extractor_window.py:648) | c1583d6 | **SURVIVED** |
| M19 | log written directly, not scheduled (extractor_window.py:1025) | c1f425e | **SURVIVED** |
| M20 | no refresh after a run (extractor_window.py:986) | 2acffed | **SURVIVED** |
| M21 | skip-refresh hard-codes `_Refreshed_` (report_processor.py:144-145) | 2aaa2b1 | **SURVIVED** |
| M22 | pywin32 never essential (requirements_check.py:57) | be6418c | SURVIVED on Linux (would be caught on Windows, HIGH) |
| M23 | one unpinned state-table entry changed (constants.py:258) [GST RULE] | edac6a3 behaviour | **SURVIVED** |
| M24 | GSTR-3B duplicates flagged like the other types (file_parser.py:267-268) [GST RULE] | edac6a3 behaviour | **SURVIVED** |
| M25 | size threshold 1 byte (helpers.py:303) | edac6a3 behaviour | CAUGHT (3) |
| M26 | scan results unsorted (helpers.py:361) | 4f84ba9 | CAUGHT (2) |
| M27 | dependency check really imports (requirements_check.py:68) | f363946 | **SURVIVED** (vacuous test) |
| M28 | Step 4 overwrites the user's folder (extract_tab.py:127-128) | bfcb24a | **SURVIVED** |

These results confirm the commit claims "Mutation-checked" in c463256 (all four fail), bfcb24a, 2aaa2b1 and d64d3a9 ("both bugs").

---

## 7. Coverage map

Method [RAN]: a scratch `sitecustomize.py` (run/A-T/covsite/) using Python's `sys.monitoring` recorded every project function that started at least once. It covered the main pytest process and the 3 subprocess probes, during a full xvfb run with HOME and TMPDIR redirected. The report script is run/A-T/covreport.py. "Ran" means the function was entered, not that every line in it was checked.

| Module | Functions entered / total | Never entered (selection) |
|---|---|---|
| core/file_organizer.py | **0 / 12** | everything: creating client folders, copying files, picking the level-1 folder for fresh/rerun/resume, cleanup, organisation report |
| core/excel_handler.py | 1 / 15 (`__init__` only) | report creation (COM and openpyxl), `_safe_open_workbook`, summary report sheets, `prepare_template_data` |
| core/file_parser.py | 9 / 11 | `get_statistics`, `validate_results` |
| gui/handlers/client_handler.py | 1 / 14 | the whole client tree, selection, `get_selected_clients`, folder-name toggles, export |
| gui/handlers/file_handler.py | 1 / 10 | browse, `scan_files`, `rescan_files`, `validate_scan_inputs`, `validate_processing_inputs` |
| gui/handlers/processing_handler.py | 5 / 10 | `process_files_thread` (the Step 3 worker), `dry_run`, `stop_processing`, `_progress_async` |
| gui/main_window.py | 10 / 39 | `toggle_dark_mode`, `save_cache`, `log_message`, `update_progress`, `update_summary`, `on_closing`, `run`, all delegate methods |
| gui/tabs/extract_tab.py | 5 / 6 | `_create_unavailable_notice` |
| gui/tabs/setup_tab.py | 20 / 22 | `_on_workflow_change`, `on_length_change` |
| gui/tabs/validation_tab.py, processing_tab.py | 6/6, 2/2 (built only, nothing asserted) | none (they are only built) |
| gui/utils/cache_manager.py | 2 / 5 | `save_cache` (but see §9: it runs when the real cache has dark mode on), get/set |
| gui/utils/ui_helpers.py | 1 / 5 | log/progress/status helpers |
| gui/widgets/tooltip.py | 2 / 6 | show/hide |
| gui/widgets/collapsible_frame.py | 1 / 2 | `toggle` |
| gui/utils/dark_mode_manager.py, setup_validation.py, status_bar.py; gui/widgets/title_bar.py | all | — |
| power_query_extractor/core/report_processor.py | 3 / 15 | `_process_report`, `_refresh_power_query_simple`, error-window handling, `_extract_data`, `_validate_refresh`, `_safe_open_workbook`, `cleanup` |
| power_query_extractor/core/data_consolidator.py | **0 / 6** | everything, including `_calculate_value` (subtract/add/multiply/divide of extracted cells) |
| power_query_extractor/config/cell_mappings.py | imported only (data) | — |
| power_query_extractor/gui/extractor_window.py | 24 / 38 | wheel handlers, skip-refresh toggle, log-wheel, `load_from_cache`, browse, select helpers, `start_processing`, `_process_clients_thread`, `_offer_to_open_report`, `_reset_after_processing` |
| power_query_extractor/extractor_main.py, launch_extractor.py, main.py | **never imported** | all |
| utils/helpers.py | 15 / 30 | timestamps, `safe_path_join`, `safe_copy_file`, `create_backup`, `truncate_path`, `get_unique_path`, `ProgressTracker`, `calculate_file_hash`, `get_relative_excel_path`, `get_short_path` |
| utils/dependency_installer.py | 4 / 6 | `restart_application`, the dialog's inner `log` |
| utils/requirements_check.py, utils/excel_com.py, refresh_naming.py | all | — |

Features with **no** tests (from the table and §5) [RAN/READ]:
- **Step 3, the actual organising and report building:** copying files into the client/version/category folders, the fresh/rerun/resume modes, Excel report creation from templates, writing the Links-sheet cells (`EXCEL_TEMPLATE_MAPPING`, constants.py:102-126), the summary report, building on local disk and copying to the network once (12d505f), the long-path and short category folder names (7347383), the "include client name in folders" settings (33d0baf fix 3), and the dry run.
- **Step 4 processing:** the Power Query refresh, data extraction, consolidation, the skip-refresh option, and Excel error-window handling.
  - [GST RULE] **TIME-VARYING**: the extraction map in cell_mappings.py has no tests. It covers the sheets 'GSTR-3B Difference' :35, 'GSTR-1 Difference' :61, 'Turnover GSTR-1' :92, 'Reco Diff' :105, 'Table 6J Match' :130, 'Table 8D Limits' :143, 'IMS Reco' :156, 'Extra In Purchase' :181 and 'Extra In GSTR2B' :206. It decides which report cell feeds which output column, and the calculations in data_consolidator.
- **Step 2, the client tree:** selecting clients, marking complete clients, the per-client folder-name toggle and export.
  - [RAN] A Treeview hands back the stored value '007' as the number 7. `get_selected_clients` (client_handler.py:167-171) would therefore build key "7-DL" for a client named "007". Hint for the client_handler reader: check how that key is used later.
- **Startup:** `main.py` (requirements gate, install offer, restart, `_pause`), launch_extractor.py and extractor_main.py.
- **Saved settings:** writing and reading `~/.gst_organizer_cache.json` (cache_manager.py; main_window.py:358-389), and the extractor reading it (extractor_window.py:587-599; extractor_main.py:36-63). No test compares the writer's keys with the readers' keys.
- **Standard 12 "Excel through COM" rules** (rebuild_pipeline_FINAL.md:844-878): none is exercised. For the COM reader, without judging: the code uses `win32com.client.Dispatch("Excel.Application")` at core/excel_handler.py:127 and report_processor.py:218, :536, while the standard names `DispatchEx`. There are `DisplayAlerts` settings, `Quit()` calls and `CoInitialize`/`CoUninitialize` pairs at the lines grep found (excel_handler.py:120-250; report_processor.py:215-412, 533-633, 781). No test exists for any of these, and none could run on Linux.
- Dead constants noticed in passing [READ grep]: `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB` and `EXCEL_SIGNATURES` (constants.py:197-202) are never used. `EXCEL_EXTENSIONS` is only re-exported (utils/__init__.py:9). `DarkModeManager._original_root_bg` is set at dark_mode_manager.py:83 but never read.

---

## 8. Checking the claims

### 8.1 tests/README.md (34 lines, never updated since 6c1e1c6) [READ/RAN]

| Line(s) | Claim | Reality |
|---|---|---|
| 3-4 | "Regression tests for … data-integrity logic" | Stale scope. The suite now also covers GUI, dark mode, dependency checks and the extractor. |
| 8-14 | run `pytest` or `python -m pytest`; needs pytest from requirements.txt | True: requirements.txt:14 has `pytest>=7.0.0`. Note that Standard 13 says an app's requirements.txt lists "the packages it imports (not pytest)" (rebuild_pipeline_FINAL.md:937-938). |
| 16-23 | table of 4 files | Stale: 14 of 18 files are missing (dark_mode, dependency_installer, excel_com, extractor_panel, extractor_standalone, main_window_tabs, optional_pywin32, processing_handler, refresh_naming, requirements_check, scroll_bindings, setup_tab_smoke, setup_validation, workflow_mode). The row for test_file_ops omits `scan_excel_files` and the cached-stat tests (added 56aaad9 and 4f84ba9). The other three rows are accurate. |
| 27-31 | characterization tests; NOTE comments mark quirks, "e.g. the case-sensitive Private→Pvt substring replace mangling Privateer" | The idea holds; the example is **stale**. 33d0baf reversed it, and test_helpers.py:43-45 now asserts that Privateer is *not* mangled. The NOTE comments at HEAD are at test_helpers.py:70, :96 and :163. |
| 33-34 | "No GUI, Excel COM, or network access is exercised, so the suite runs fast (<1s) and anywhere." | **False at HEAD.** About 46 tests build Tk GUIs (33 in-process, 13 in subprocesses). The suite takes about 3.5–4 s [RAN]. It does not run "anywhere": without a display there are 14 failures and 33 skips, and 1 test fails on Linux. COM: still not exercised (fakes only), true. Network: the tests touch none themselves, but the main-window test reads the user's real settings file and may check paths stored in it (§9). The original 4 files do run in 0.32 s [RAN]. |

### 8.2 Commit 6c1e1c6 (23-07-2026), "Add pytest suite for data-integrity logic (81 tests)"

- "81 tests": **exact** [RAN]. Exporting 6c1e1c6 and collecting gives 81 (file_ops 14, file_patterns 21, helpers 41, scan 5), all passing in 0.25 s.
- "First automated tests for the repo": true [READ: `git log --all` shows no test file before it, in either A or B].
- "zero GUI/COM/network coupling … runs in <1s anywhere": true when it was written [RAN]. One nuance: importing `core.file_parser` runs `core/__init__.py:6-8`, which imports excel_handler. On Windows that tries `import win32com.client` (core/excel_handler.py:21-28) and sets `warnings.filterwarnings('ignore')` (:17). That is an import-level coupling only. [RAN] The global warnings filter does not hide warnings in A's suite as written, because the modules are imported during collection; it would hide warnings inside a test that first imports core.
- File descriptions: accurate for that commit.
- "Known quirks (substring 'Private'->'Pvt' mangling, 'UP' fallback collision) are pinned with NOTE comments": true then. The first was reversed in 33d0baf.
- "Tooling: pytest.ini …, conftest …, README …, pytest uncommented in requirements, .pytest_cache/ gitignored": all true [READ diff].
- "mutation-checked": no artefact in the repo. My own mutations of the same logic (M01–M04, M25, M26) are caught [RAN].

### 8.3 Test counts in later commit messages [RAN: each commit exported and collected]

| Commit | Claim | Collected on Linux | Verdict |
|---|---|---|---|
| ba2a753 | "3 tests" | 84 → 87 | correct |
| 4219ffb | "92 tests pass" | 92 | correct |
| 56aaad9 | "99 tests pass (7 new …)" | 99 | correct |
| 4f84ba9 | "100 tests pass" | 100 | correct |
| 322717f | "21 unit tests; 9 smoke tests … 127 tests pass" | 127 (setup_validation 19, smoke 8) | total correct; breakdown **wrong** (19 and 8, not 21 and 9; 100+21+9 would be 130) |
| d64d3a9 | "5 new scroll tests … 140 tests pass" | 132 (= 127 + 5) | **"140" does not match**. Maybe the author's run included uncommitted test files (LOW). |
| 2aaa2b1 | "Tests: 20 new" | 21 new (19 refresh_naming + 2 panel) | off by one |
| a677670 … aeb83cf | — | on Linux, collection errors (`ModuleNotFoundError: No module named 'pythoncom'`): 1 error at a677670, 2aaa2b1 and 7347383; 2 at bfcb24a; 3 at 510f7b4; 4 at aeb83cf | Shows the c463256 bug: at those commits the test files that import GUI modules could not even be collected on a Python without pywin32 |

[RAN/READ, MEDIUM] At those historical commits, *which* files failed collection depended on the order of imports. A module half-imported by an earlier failed import stays cached, and some later files then imported fine. For example, at aeb83cf test_setup_tab_smoke and test_setup_validation collected while test_processing_handler did not. This no longer matters at HEAD.

### 8.4 Other claims about tests

- **Design spec** docs/superpowers/specs/2026-07-25-merge-extractor-design.md:107-120:
  - "`_classify_version_reports` and `_latest_refresh_times` are unit-tested (`tests/test_extractor_helpers` equivalents live in the suite)": **false when written** (b276542). No test referenced either helper, and `tests/test_extractor_helpers*` never existed in any branch [RAN `git log --all`]. At HEAD, `_latest_refresh_times` is tested (since 2aaa2b1); `_classify_version_reports` only runs indirectly through scans, and nothing asserts on it.
  - "same approach as the existing `_log_async` tests": no `_log_async` test exists [READ grep].
  - "Unit tests: workflow-mode gating … as pure logic, no Tkinter": the logic is tested with fakes, but importing it loads tkinter (§10.4).
  - The planned smoke tests (panel, shell, main app with four tabs) exist.
- **Verification evidence named in commit messages but never committed** [READ; RAN `git log --all -- tests …`]:
  - 0890022: "Verified with a test that reproduces the bug";
  - 275d9a4: "proven equivalent … across 9 folder layouts … with coverage asserted";
  - a603c02 and 12d505f: "characterization snapshot";
  - 4219ffb: "Folder-structure harness";
  - 56aaad9 and 4f84ba9: "scan characterization snapshot".
  None of these tests or harnesses is in any commit of A or B. The fixes in 0890022, 275d9a4 (level-1 folder cache) and 12d505f have no committed test today.
- **Docstring claims in the tests:**
  - test_setup_tab_smoke.py:1-6, "Skipped automatically where no display/Tk is available": true [RAN].
  - test_extractor_standalone.py:3-8 and test_main_window_tabs.py:3-5, skip when the window cannot start: not true without a display (§4.2).
  - test_setup_validation.py:1, "no Tkinter involved": tkinter is imported but no window is made (§5.9).
  - test_processing_handler.py:1, "no GUI/threads exercised": true for execution; importing it loads the whole GUI.

---

## 9. Tests that write (or read) outside a temp folder [RAN]

1. **tests/test_main_window_tabs.py**, through `GSTOrganizerApp()` in its subprocess:
   - **Creates `<system temp>/gst_excel_temp`** (`ExcelHandler.__init__`, core/excel_handler.py:53-54, called from gui/main_window.py:116). This is the folder the real app uses to build reports locally (12d505f). [RAN exp1: with TMPDIR redirected, `gst_excel_temp` appeared there.] The same folder already existed in /tmp, from the earlier quick run (21-10-2026 03:21:36).
   - **Reads the real user settings file `~/.gst_organizer_cache.json`** (gui/utils/cache_manager.py:15, :18-27; main_window.py:121, :373-389). Its stored source, template and target paths are put into the Setup tab, whose existence checks (setup_tab.py:521-529) can then touch those real, possibly network, paths during the test (MEDIUM: the 350 ms debounce versus the probe's `update()` calls).
   - **Rewrites `~/.gst_organizer_cache.json` when it says dark mode is on.** `apply_cached_values` → `toggle_dark_mode()` (main_window.py:388-389) → `self.save_cache()` (:196) → cache_manager.py:45-46. The probe's `app.save_cache = lambda …` comes too late (test_main_window_tabs.py:38). [RAN exp2: a fake home cache with `"dark_mode": true` was rewritten during the test and gained `"recent_folders"` plus the defaults; 7 tests passed.] With no cache file, nothing is written [RAN exp1]. On the author's PC this test can therefore change his real app settings file (HIGH).
   - Standard 13 says settings belong in `%LOCALAPPDATA%\[App Name]\` (rebuild_pipeline_FINAL.md:1039-1040). A keeps them in the home folder as a dotfile. Recorded only, for the audit.
2. **pytest itself and the subprocess probes**, when run the normal way (without `-B`, `PYTHONDONTWRITEBYTECODE` or `-p no:cacheprovider`), write `__pycache__/` folders into the project's source folders (tests/, core/, gui/…, power_query_extractor/…, utils/) and a `.pytest_cache/` in the project root. Both are git-ignored (.gitignore:2-3). The probes start `sys.executable` without `-B` (test_extractor_standalone.py:69, test_main_window_tabs.py:94, test_optional_pywin32.py:73) [READ]. This is how A-copy acquired its stray `__pycache__` (§2).
3. pytest's own temp folders sit under the system temp (`/tmp/pytest-of-<user>/`; on Windows `%TEMP%\pytest-of-<user>`). This is standard behaviour. Every other test writes only into `tmp_path` or `tmp_path_factory` [READ].
4. No test runs pip for real (`subprocess.run` is patched; test_dependency_installer.py:46-84). No test starts Excel. No test writes a log file: `main.py` and `extractor_main.py`, which configure `gst_organizer.log` and `pq_extractor.log`, are never imported [RAN coverage].

---

## 10. Test-suite hygiene issues (technical, no GST content)

1. **State leaks from the dark-mode tests into later GUI tests** [RAN]. Several tests apply dark mode and never restore it: test_dark_mode.py:112-121, :123-127, and :129-141 (dark applied twice, restored once). The manager walks and themes the *shared* session root and the global ttk styles (dark_mode_manager.py:105-115). Measured afterwards in a scratch probe test: root background `#2b2b2b`, `TFrame` `#2b2b2b`, `TButton` `#404040`; without test_dark_mode they are `#d9d9d9`. No later test checks colours, so nothing fails.
2. **Pending `after` callbacks outlive their widgets** (§3.3). The 5 "invalid command name" tracebacks come from here.
3. **Each SetupTab and ExtractorPanel built in a test adds one more application-level `<MouseWheel>` binding** to the shared root (setup_tab.py:201; extractor_window.py:478; `add='+'`, never removed) [READ]. They pile up across the session. The handlers cope with destroyed canvases (setup_tab.py:224-228; extractor_window.py:459-463).
4. **Importing almost any `gui.*` module loads the whole app.** gui/__init__.py:7 imports main_window, which imports every tab and the extractor (report_processor, data_consolidator, cell_mappings). [RAN] 222 modules load for `from gui.utils import setup_validation`. Consequences: "pure logic" GUI tests depend on everything importing cleanly. This is how the pywin32 import bug broke collection at a677670–aeb83cf (§8.3).
5. **Weak or vacuous tests:**
   - test_requirements_check.py:27-32 can never fail (M27);
   - test_dark_mode.py:182-188 does not test what its name says;
   - test_dark_mode.py:190-198 checks source text, defeated by M08;
   - test_workflow_mode.py:85-90 checks a log string, not that the refresh was not started;
   - test_scroll_bindings is text-based (M12b).
6. **Stale comments in tests:** test_helpers.py:110-111 (about `sanitize_filename`'s "substring behaviour"); tests/README.md:28-30.
7. **Misleading failure message:** test_requirements_check.py:71 says "openpyxl is needed", but the assertion covers every essential package, including pywin32 on Windows.

---

## 11. Relation to project B [READ/RAN]

- **B has no tests at all.** `git log --all` over tests/, `test_*` and `pytest.ini` lists only A commits. `git ls-tree -r B/main` shows no test files. The side branches `A/claude/pensive-feynman-0mz1aa` (= bf2f1dd) and `B/claude/pensive-feynman-0mz1aa` (= b0e7d0e) add nothing. All A tests were written after the split (first one 6c1e1c6, 23-07-2026; the split point is 47a1492, 30-10-2025).
- [RAN] A's four pure-logic files, run against B's code in a scratch copy (run/A-T/AtestsOnB/, venv-B, offscreen Qt):
  - test_file_ops.py cannot be collected: `ImportError: cannot import name 'scan_excel_files' from 'utils.helpers'`. B has the older `find_excel_files` and `validate_excel_file` without the cached-stat parameter (B utils/helpers.py:210, :249).
  - test_file_patterns + test_helpers + test_scan give **7 failed, 66 passed**:
    - 4 IMS trailing-"(N)" tests fail: B's IMS pattern has no "(N)" (B utils/constants.py:23);
    - 3 sanitise tests fail: `test_abbreviation_uses_word_boundaries`, `test_abbreviation_is_case_insensitive` and `test_run_together_name_is_left_alone`. B still uses `base_name.replace('Private', 'Pvt')` (B utils/helpers.py:78).
  - So A's tests mark exact behaviour differences between A and B. In a merge, whichever version is kept, these tests decide which behaviour is "right" for file naming and IMS file recognition. IMS recognition is a [GST RULE] classification (§5.3): the author decides.
- The rest of A's suite is Tkinter-specific (B uses PyQt5) and cannot be reused for B's GUI as it stands [READ, HIGH].

---

## 12. GST-rule items the tests touch (described only; no judgement) [READ]

| Item | Where in code | Pinned by | Notes |
|---|---|---|---|
| [GST RULE] File → return/report type classification (6 types), with client/state/period groups. **TIME-VARYING** | utils/constants.py:13-62 | test_file_patterns.py (21 tests) | GSTR-3B is the only type that tolerates a leading "(N) "; for all types except IMS the trailing "(N)" lands in the period metadata (§5.3) |
| [GST RULE] IMS Reco date = 8 digits, described as DDMMYYYY | constants.py:23, :27 | test_file_patterns.py:63-67 | the 8 digits are not checked as a real date |
| [GST RULE] Six expected inputs for a "Complete" client. **TIME-VARYING** | constants.py:68-75; core/file_parser.py:252-286 | test_scan.py:18-38, 67-80 | "Missing N files" / "Complete" |
| [GST RULE] Several GSTR-3B files allowed per client; duplicates of other types flagged | file_parser.py:262-271 | **none** (M24 survived) | comment "monthly filings" |
| [GST RULE] 'Annual Report' and 'GSTR-3B Export' treated as critical missing types | file_parser.py:359 | **none** | `validate_results` never runs |
| [GST RULE] State name → 2-letter code table; fallback rules; 'Unknown Place'→'UP'. **TIME-VARYING** | utils/helpers.py:592-618; constants.py:251-292 | test_helpers.py:84-103 (5 of 40 entries) | not GSTIN codes; M23 survived |
| [GST RULE] Client key "<client>-<STATECODE>", 35-character cap with Pvt/Ltd shortening | helpers.py:620-644; also file_parser.py:229-230 (raw client, no cap) | test_helpers.py:108-122; test_scan.py:25 (`ABC-DL`) | two different key builders exist; the link to folder names (file_organizer.py:83) is untested |
| [GST RULE] Report extraction map (GSTR-3B/GSTR-1 differences, Table 6J, Table 8D, IMS Reco, …) and its calculations. **TIME-VARYING** (table numbers) | power_query_extractor/config/cell_mappings.py; data_consolidator.py:188 | **none** | completely untested |

---

## 13. Possible client names

**None found** in tests/*.py, tests/README.md, pytest.ini, or the messages and diffs of the test-touching commits. The test data uses obviously generic placeholders, which are code, not client data:
- client names: "ABC", "XYZ", "Client00"–"Client29", "Valid Name", "A Very Long Client Name That Exceeds The Limit For Sure", "Privateer Unlimited";
- folders and states: "ABC-DL", the states Delhi, Maharashtra, Tamil Nadu and West Bengal, "Unknown Place";
- paths: `C:\gst\…`, `\\server\share\level0..19`.

[RAN] A GSTIN/PAN-shaped pattern search over the tests and the A commit messages since 6c1e1c6 found nothing.

Other sensitive-looking strings that are *not* client names (locations only, not copied): the messages of commits **4f84ba9** and **7347383** quote the author's real office network-share path (server and folder names) and file counts from real runs.

---

## 14. Questions for the author

1. Which Python do you run the tests with on your PC (`python` or `py`), and which Tk version does it report (`python -c "import tkinter; print(tkinter.TkVersion)"`)? This machine has Tk 9.0. If you use `py` (said to be the free-threaded 3.14t without pywin32), `test_openpyxl_is_actually_installed_here` will fail, and it is not certain that tkinter is available there.
2. Did the main-window test ever change your `~/.gst_organizer_cache.json` (for example new "recent_folders" entries, or settings reset after running tests)? It rewrites that file whenever it says dark mode is on (§9).
3. Commit d64d3a9 says "140 tests pass", but that commit contains 132 tests. Did you have extra test files that were never committed? The harnesses described in 0890022, 275d9a4, a603c02, 12d505f, 4219ffb, 56aaad9 and 4f84ba9 are also not in the repository. Do they still exist anywhere?
4. For the merge: should A's pinned behaviours be the reference? Specifically: "Privateer" not abbreviated, ALL-CAPS "PRIVATE LIMITED" abbreviated, run-together names left alone, and IMS files with a trailing "(1)" recognised. B behaves differently on all of these (§11). The IMS recognition is a GST-rule classification for you to decide.
5. Should a trailing "(1)" on GSTR-2B/Sales/Annual/GSTR-3B/Sales-Reco files stay inside the period/month/year value (current behaviour, untested)? Should a leading "(1) " be accepted for types other than GSTR-3B?
6. The state-code table has 40 entries but only 5 are tested, and "Unknown Place" → "UP" collides with Uttar Pradesh. Do you want the full table pinned by tests in the rebuild? Which table is authoritative is your call.

---

## INCOMPLETE

- **Not run on Windows.** All "would pass on Windows" verdicts are reasoned from the code, with two small simulations (path rules, essential-pywin32). They are not executed results. In particular, the geometry ("1000x700") and the scroll-overflow tests depend on Windows fonts and scaling (MEDIUM-HIGH).
- **Not run with the author's real `~/.gst_organizer_cache.json`, or with network paths in it.** The effect of real cached paths on the main-window test's speed and result is inferred (MEDIUM).
- The coverage map is at **function** level ("entered at least once"), not line or branch level, because coverage.py is not installed and installing it would need web access.
- I did not reconstruct aeb83cf's exact "five tests" mutation. I tested the whole pre-aeb83cf file instead (8 of 11 fail).
- Non-test modules were read only where they bear on the tests (§2 lists which line ranges). Their full audit belongs to the other readers.

Scratch artefacts (all under /root/merge-scratch/run/A-T/):
- run outputs: run1_xvfb.txt, run2_nodisplay.txt, collected_ids.txt;
- history exports: hist/<commit>/ and at_6c1e1c6/;
- mutations: mutate.py, mut/, mut_results_all.jsonl;
- coverage: covsite/, covrun/, covout/, covreport.py;
- experiments: exp1/ and exp2/ (cache and temp), exp3/ (warnings filter), leak/ (dark-mode state leak), olddark/ (pre-aeb83cf dark mode), AtestsOnB/ (A's tests against B), headcheck/ (A-copy vs A/main).
