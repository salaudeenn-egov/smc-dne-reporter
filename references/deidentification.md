# De-identification: what it does and what it costs

## Removed or replaced

| Source field | Treatment |
|---|---|
| Child Name, Household Head Name | Removed after deriving: blank / short / digits / symbols / single word / length, head blank / short, child = head, `name_dup_key` (name+head+age+sex+LGA+ward), `household_key` (head+CDD+GPS 4 dp) |
| latitude, longitude | Removed after deriving: `gps_status`, `gps_low_precision`, `gps_km_from_hf_median`, `gps_point_key`; facility median point kept at 3 dp (about 110 m) |
| Individual ID | `child_key` = HMAC-SHA256(salt, ID), 12 hex; `id_created_date` from the time-based ID |
| Client Reference ID | `record_key` (same method) |
| Username / userName | `worker_code` CDD-0001, OIC-0001 ...; map kept privately |
| `name` in user reports (real CDD name) | Removed |
| Any column matching name/phone/address/lat/lon/NIN/email | Removed and listed in `_deidentification_audit.csv` |

The salt and username map live in `<out_root>_PRIVATE_do_not_share/`. Re-runs reuse the salt, so keys are stable
across cycles and runs. A new salt breaks linkage with every earlier extract.

## Proof it does not change the numbers

`verify.py parity` recomputes rows, children dosed, repeat children, extra records, workers, worker volumes,
GPS-missing and children dosed in both cycles on raw and de-identified data. All must match (Oyo C3/C4: 17 of 17).

## What the receiver can do

Read the report; recompute every exhibit from the shared CSVs; follow a child across cycles by `child_key`; rank
CDDs by `worker_code`; rerun all 40 checks that were pre-flagged.

## What the receiver cannot do (the trade-offs)

1. No field follow-up. They cannot find a child, household or CDD. Case lists for repeat doses or out-of-range
   ages must come back through the data owner, who holds the key.
2. No new name or GPS analysis. Fuzzy duplicate matching, name match between cycles, GPS recapture distance and
   household maps need the raw fields (E8.4 rows show Not assessable).
3. No lookup in the live system by ID, so report-vs-source mismatches (V37) cannot be investigated from the
   shared files.
4. No linkage with other extracts made with a different salt.
5. Residual risk. Age + sex + ward + date can still single out a child in a very small ward. Share row-level
   files only inside the programme; share the report more widely.
