# smc-dne-reporter

A Claude Code skill that turns one state's DIGIT HCM SMC exports into the standard Data & Evidence
**SMC Cycle Report** (Word + PDF), following `SMC_Report_Standard_Spec_v1.xlsx`: 15 sections, 40 validation
checks, anomaly register, scorecard, recommendations and 12 charts. It works from a de-identified copy of the
data, so the inputs and the report can be shared.

## Pipeline

```
python -X utf8 scripts/run_all.py configs/<state>_c<n>.json [--from analysis|report|scan] [--no-pdf]
```

| Stage | Script | Gate |
|---|---|---|
| De-identify | `deidentify.py` | Leak scan: no names, usernames or IDs in outputs |
| Parity | `verify.py parity` | Core metrics on raw and de-identified data identical |
| Analysis | `analysis.py` | Reads de-identified inputs only |
| Report | `build_report.py` | Rule-based ratings, anomalies, recommendations |
| Report scan | `verify.py report` | No names, IDs, placeholders or "nan" in the docx |
| PDF | Word (Windows) | |

Outputs go to `out_root`: `01_input_deidentified/` (shareable), `02_analysis/`, `03_report/`.
Keys (salt + username map) go to `<out_root>_PRIVATE_do_not_share/` and must never be shared or committed.

## Install as a skill

Copy this folder to `~/.claude/skills/smc-dne-reporter/`. See `SKILL.md` for the workflow,
`references/data_gotchas.md` for export quirks and `references/deidentification.md` for the trade-offs.

Requirements: Python 3.11+, pandas, numpy, matplotlib, python-docx, python-calamine, xlsxwriter.
Supports Cycle 2 and later (needs the previous cycle). No campaign data is stored in this repository.
