# Data gotchas (Nigeria SMC 2026 exports)

Problem -> what we saw -> how the pipeline handles it / what to do.

1. **Stale registered denominator.** Oyo C4: 112 of 115 facilities had the same Registered Count as C3; 66
   facilities dosed more children than registered; state coverage 112.7%. -> analysis flags G09/V39; report adds
   a footnote, drops scenario estimates and does not rank LGA coverage. Tell the user coverage is a volume
   indicator until the count is refreshed.
2. **HF summary vs raw export.** Oyo C4 HF summary showed +13,450 successes over the raw export although both
   files carry the same generation stamp. -> V37 Action, A31. Report as measured, cause not established.
3. **Cumulative user files.** Some HF_USER / HF_User_activity drops (Plateau C4, FCT C4, Zamfara C3, Borno C3)
   are all-cycles cumulative. Check: sum ADMINISTRATION_SUCCESS against the same-cycle HF summary; if it is
   about N times larger, do not use it (set `cycles.<n>.files.user_agg` to a correct file or leave it out).
4. **Daily user_activity_report sums are about 3x HF admin counts** in every state. This is the normal format, not
   cumulative. The pipeline uses it only for account lists and active days, never for dose totals.
5. **Bauchi C1 user activity is capped at 10,000 rows.** Accounts and activation will be understated; get a re-export.
6. **Multi-part children-treated exports** (Bauchi 4 parts, Kebbi RI re-pull). The pipeline concatenates every
   matching file and sheet and records `source_part`. Kebbi C3 main file and the `RI\` re-pull OVERLAP: never
   sum them; point `files.ct` at one set.
7. **ORS-Zinc rows.** Kebbi has a 27-column ORSZ_TREATED format. It must not be mixed into the SPAQ export.
8. **Missing eligibility fields.** The Oyo export has no SP/CTX, sick/fever or first-time-SPAQ columns, so V30, V31
   and V34 show Not assessable. Other states (Taraba, Bauchi) do have them; the pipeline auto-detects them.
9. **Hand-typed facility names.** Facility names usually match across reports. When V38 > 0, list the variants
   and ask the user before mapping them. Never infer a missing LGA for a facility; ask.
10. **Person names that equal place names.** Some household-head entries are a ward or facility name. The
    report scan ignores names that are substrings of LGA, ward or facility names.
11. **Record Type = duplicate.** This is the export's own flag, not the repeat-dose rule. F08/F09 use it; S01/S02
    (repeat dosing) use child-key counts. Do not mix the two in narrative.
12. **Code words typed as names.** Nasarawa C3/C4 have child or head "names" equal to system codes
    ("unique", "blank", "ok" ...). The de-identify leak scan used to flag the Record Type and gps_status
    columns as leaks. -> System code values (status, record type, gender, product, symptom, cycle index,
    gps labels) are treated as non-personal, like place names. A real name in any column is still caught.
13. **Junk names made of ordinary words** ("may have", "at least"). The report scan ignores names whose
    every word also appears in the report's own prose (build_report.py, template, spec); see verify.py.
