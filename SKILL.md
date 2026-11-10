---
name: smc-dne-reporter
description: >-
  SMC D&E Reporter - turns one state's DIGIT HCM SMC exports (children-treated, HF summary,
  user activity, stock report) into the standard Data & Evidence "SMC Cycle Report"
  (SMC_Cycle_Report_Template_v1.docx + SMC_Report_Standard_Spec_v1.xlsx): 15 sections, 40
  validation checks, anomaly register, scorecard, recommendations, 12 charts, Word + PDF. It
  first makes a de-identified copy of the data (no child/head/CDD names, no GPS points, no
  system IDs) so inputs and report can be shared. Use whenever the user wants a cycle report,
  D&E report, effectiveness/activity report, "fill the SMC template", "report for <state>
  cycle N", wants to automate the SMC report template, or wants SMC data de-identified / PII
  removed before sharing - even if they don't name the template. Triggers: "SMC cycle report",
  "D&E report", "standard report", "report spec", "remove PII and share", "Cycle 4 report for
  Oyo/Bauchi/Kebbi/...", "validation checks V01-V40", "anomaly library".
---

# SMC D&E Reporter

Produces the standard SMC cycle report for one state and one cycle, compared with the
previous cycle, from a de-identified copy of the exports. The scripts do the heavy lifting.
Your job is to set up the config, run the pipeline, read the numbers with a critical eye and
explain the result.

## Files

- `assets/SMC_Cycle_Report_Template_v1.docx` is the Word template. Its styles, header and footer are reused.
- `assets/SMC_Report_Standard_Spec_v1.xlsx` holds the metric IDs, the 40 checks, the anomaly
  triggers and the rating scales. The scripts implement these. Read the spec when a user asks
  "why is X rated Y".
- `scripts/run_all.py <config>` runs the whole pipeline. The stages are below; each can be run alone.
  1. `deidentify.py` derives flags from personal fields, then removes or replaces those fields,
     then scans for leaks.
  2. `verify.py <config> parity` checks that core metrics on raw and de-identified data are
     identical. It must say SAME on every row.
  3. `analysis.py` computes all metrics, tables and figures from the de-identified inputs only.
  4. `build_report.py` builds the Word report. Ratings, anomalies, recommendations and
     stakeholder asks are rule-based.
  5. `verify.py <config> report` scans the finished report for names, usernames, IDs,
     `{{placeholders}}` and the text "nan".
  6. On Windows, Word converts the report to PDF.
- `configs/oyo_c4.json` is a worked example config.
- `references/data_gotchas.md` covers export quirks seen in real states. Read it before trusting a
  surprising number.
- `references/deidentification.md` covers what is removed, what survives, and the trade-offs
  for whoever receives the files.

## Workflow

1. **Agree the scope.** Confirm the state, the cycle and the previous cycle, and give the user a short
   plan before running anything (the user prefers plan, then approval, then run). The scripts
   cover Cycle 2 and later, because a previous cycle is needed. For a Cycle 1 report, say so and
   adapt by hand: drop sections 8 (detail), 12 and the previous-cycle columns.
2. **Locate the inputs.** Per cycle you need a folder holding the children-treated export (all
   parts), the HF administration-status summary, `user_activity_report.xlsx` and the stock
   report. Optional: `healthfacility_activity_report.xlsx` and the user aggregation report. On
   this machine, the Nigeria 2026 dumps are under `Downloads\SMC_Nigeria_2026\Cycle-<n>\<state folder>`.
   Profile the files first. Stubs, cumulative multi-cycle user files and duplicate re-pulls
   exist (see data_gotchas.md).
3. **Write the config.** Copy `configs/oyo_c4.json` to `configs/<state>_c<n>.json`. Set:
   - `state`, `country`, `cycle`, `prev_cycle`, `bbox` (state bounding box: min lat, max lat, min lon, max lon)
   - `cycles` (folder per cycle; use `{"dir": ..., "files": {"ct": [...], "stock": "..."}}` when the auto-match picks the wrong file)
   - `out_root`, and `thresholds` only when the programme asked for different ones

   If discovery reports several matches, ask the user which file is right rather than guessing.
4. **Run** `python -X utf8 scripts/run_all.py configs/<file>.json`. It stops at the first failed
   gate. Fix the problem, then re-run with `--from <stage>`. A full state takes about 3 to 6 minutes.
5. **Sense-check before handing over.** The report is generated, not reviewed. Look at:
   - Coverage above 100%. This usually means a stale denominator (G09). The report already
     caveats it when 90% or more of facilities are unchanged; say so plainly to the user.
   - HF summary vs raw export (V37). Report it as measured, cause not established, unless
     you have checked the source index.
   - The anomaly register and the recommendations. Remove anything that is plainly an export
     artefact, and tell the user what you removed.
6. **Deliver.** Give the paths to the report (docx + pdf), the analysis workbook, and the
   shareable `01_input_deidentified` folder. State clearly that the
   `<out_root>_PRIVATE_do_not_share` folder holds the salt and username map and must never be
   shared. Summarise 3 to 5 headline findings and the main things lost to de-identification.

## Writing rules for anything you add to the report

- Plain words, no emojis or decorative symbols. Label findings as measured or inferred, and
  mark causal claims [ASSOCIATION].
- With partners, avoid blame words. Say "not dosed again", not "lost"; say "missing", not "blank".
- Every number must come from the analysis outputs, never from a dashboard.
- Never put a username, child name or ID in the report or in chat. Use CDD codes.
- If you change the template or the rules, change the scripts, not the generated docx, so the
  next state gets the fix.

## Extending

- A new check that needs a name, GPS or ID value has to be added in `deidentify.py` as a
  flag. After de-identification the raw value is gone. Re-run from the deidentify stage. The
  salt is reused, so keys stay stable across runs and cycles.
- Anomaly triggers, scorecard rules and recommendation text live in `build_report.py` (the
  "anomaly engine" and "scorecard rules" blocks). Thresholds come from the spec defaults and can
  be overridden in the config `thresholds`.
