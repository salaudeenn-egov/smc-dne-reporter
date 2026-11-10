# SMC cycle report: de-identified input data

These files are the inputs for the standard SMC cycle report (SMC_Report_Standard_Spec_v1).
Personal data has been removed. Every row from the source exports is kept (row counts are in `_deidentification_audit.csv`).

## Files

| File | Spec code | One row per |
|---|---|---|
| IN-1_children_treated_C<n>.csv | IN-1 | child contact record |
| IN-2_hf_summary_C<n>.csv | IN-2 | health facility |
| IN-3_hf_activity_by_day_C<n>.xlsx | IN-3 | health facility (daily counts by column) |
| IN-3b_user_outcome_agg_C<n>.csv | IN-3 | CDD account |
| IN-4_stock_C<n>.csv | IN-4 | health facility |
| IN-5_user_activity_by_day_C<n>.xlsx | IN-5 | CDD or OIC account (daily counts by column) |
| IN-8_facility_points_C<n>.csv | IN-8 | health facility (median dose location, 3 decimals) |

## What replaced the personal fields

| Column | Meaning |
|---|---|
| child_key | Stands in for Individual ID. Same child = same key in both cycles, so retention and repeat dosing still work |
| record_key | Stands in for Client Reference ID |
| worker_code | Stands in for the username (CDD-0001, OIC-0001). Same CDD = same code in every file and cycle |
| id_created_date | Date the original ID was created (from the time-based ID), used for "new vs earlier registry" |
| name_blank, name_short, name_has_digit, name_has_symbol, name_single_word, name_length | Child-name quality flags (1 = yes) |
| hh_head_blank, hh_head_short, name_equals_head | Household-head flags |
| name_dup_key | Same value = same child name + head name + age + sex + ward (probable duplicate registration) |
| household_key | Same value = same head name + CDD + GPS within about 11 m (household proxy) |
| gps_status | ok / blank / zero / invalid / outside_state |
| gps_low_precision | 1 = 3 or fewer decimals |
| gps_km_from_hf_median | Distance in km from the facility's median dose point |
| gps_point_key | Same value = identical GPS point (stationary-entry check) |
| timestamp_local | Timestamp in WAT (UTC+1) |

The keys cannot be reversed without the private key file, which is not shared.

## What you cannot do with these files

- Find or contact a specific child, household or CDD.
- Run a new name or GPS check that is not already a flag (for example fuzzy name matching, GPS distance between cycles, or a household map).
- Match these records against the live system or another dataset by ID.
- Compare keys with another de-identified extract made with a different key.
