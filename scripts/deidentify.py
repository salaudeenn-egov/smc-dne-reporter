"""Stage 1 - de-identify one state's SMC exports (current + previous cycle).

Usage:  python -X utf8 deidentify.py <config.json>

Derives every flag the standard report needs from personal fields (names, GPS,
IDs, usernames) FIRST, then writes copies with those fields removed or replaced
by keyed pseudonyms. Anything not derived here is permanently unavailable to
the analysis, so add new flags here, never in analysis.py.

Shareable output : <out_root>/01_input_deidentified/
Private keys     : <out_root>_PRIVATE_do_not_share/  (salt + username map)
"""
import hashlib, hmac, json, re, secrets, sys, uuid
from pathlib import Path
import numpy as np
import pandas as pd

cfg = json.load(open(sys.argv[1], encoding="utf-8"))
ROOT = Path(cfg["out_root"])
OUT = ROOT / "01_input_deidentified"
PRIV = ROOT.parent / (ROOT.name + "_PRIVATE_do_not_share")
OUT.mkdir(parents=True, exist_ok=True); PRIV.mkdir(parents=True, exist_ok=True)
BBOX = cfg["bbox"]                       # [min lat, max lat, min lon, max lon]
CYCLES = [c for c in (cfg.get("prev_cycle"), cfg["cycle"]) if c]

PATTERNS = {   # case-insensitive, matched on file name; override in config "patterns"
    "ct": r"(children_treated|child_treatment)",
    "hf": r"^(?!.*user).*hf_administration_status_agg",
    "stock": r"stock",
    "user_activity": r"^user_activity_report",
    "hf_activity": r"^healthfacility_activity_report",
    "user_agg": r"user_administration_status_agg",
}
PATTERNS.update(cfg.get("patterns", {}))
REQUIRED = {"ct", "hf", "stock", "user_activity"}
CT_REQUIRED_COLS = ["Individual ID", "LGA", "Ward", "Health Facility", "Username", "Child Name", "Age", "Gender",
                    "Household Head Name", "latitude", "longitude", "productName", "Administration Status", "Timestamp", "Client Reference ID"]
USER_COLS = {"username", "user name", "userName".lower()}
# any column whose header matches this is treated as personal data and removed (audit lists it)
PII_COL_RE = re.compile(r"(?i)(child name|head name|^name$|phone|mobile|msisdn|address|^lat|^lon|national id|\bnin\b|email)")

salt_f = PRIV / "salt.hex"
SALT = bytes.fromhex(salt_f.read_text().strip()) if salt_f.exists() else secrets.token_bytes(32)
salt_f.write_text(SALT.hex())


def h(prefix, v, n=12):
    if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "":
        return ""
    return prefix + hmac.new(SALT, str(v).strip().encode(), hashlib.sha256).hexdigest()[:n]


def norm_name(s):
    s = "" if s is None or (isinstance(s, float) and np.isnan(s)) else str(s)
    return re.sub(r"\s+", " ", re.sub(r"[^a-z ]", "", s.lower())).strip()


def uuid1_date(s):
    try:
        u = uuid.UUID(str(s))
        if u.version != 1:
            return ""
        return pd.Timestamp((u.time - 0x01B21DD213814000) / 1e7, unit="s").strftime("%Y-%m-%d")
    except Exception:
        return ""


def cycle_files(c):
    spec = cfg["cycles"][str(c)]
    folder = Path(spec if isinstance(spec, str) else spec["dir"])
    explicit = {} if isinstance(spec, str) else spec.get("files", {})
    allx = [p for p in folder.rglob("*.xls*") if "chunk" not in str(p).lower() and not p.name.startswith("~$")]
    out = {}
    for key, pat in PATTERNS.items():
        if key in explicit:
            v = explicit[key]; out[key] = [Path(x) for x in (v if isinstance(v, list) else [v])]; continue
        hits = sorted(p for p in allx if re.search(pat, p.name, re.I))
        if key == "ct":
            out[key] = hits
        elif len(hits) == 1:
            out[key] = hits
        elif len(hits) > 1:
            raise SystemExit(f"Cycle {c}: {len(hits)} files match '{key}' in {folder}: {[p.name for p in hits]}. "
                             f"Name the right one in config cycles.{c}.files.{key}")
        else:
            out[key] = []
        if key in REQUIRED and not out[key]:
            raise SystemExit(f"Cycle {c}: no file for required input '{key}' (pattern {pat}) in {folder}")
    return out


def read_all(paths, need=None):
    """Read every sheet of every file; keep sheets whose header has the needed columns."""
    frames = []
    for p in paths:
        for sh, df in pd.read_excel(p, sheet_name=None, engine="calamine", dtype=str).items():
            if need and not set(need) <= set(df.columns):
                continue
            df["source_part"] = f"{p.name}[{sh}]"; frames.append(df)
    if not frames:
        raise SystemExit(f"No sheet with columns {need} in {[p.name for p in paths]}")
    return pd.concat(frames, ignore_index=True)


def user_col(df):
    return next((c for c in df.columns if c.lower() in USER_COLS), None)


# ------------------------------------------------------------------ load + worker map
F, all_users, raw_names, admin_vals = {}, set(), set(), set()
for c in CYCLES:
    fl = cycle_files(c)
    ct = read_all(fl["ct"], CT_REQUIRED_COLS)
    ua = pd.read_excel(fl["user_activity"][0], sheet_name=None, engine="calamine", dtype=str)
    ag = pd.read_excel(fl["user_agg"][0], sheet_name=None, engine="calamine", dtype=str) if fl.get("user_agg") else {}
    F[c] = dict(fl=fl, ct=ct, ua=ua, ag=ag)
    all_users |= set(ct["Username"].dropna())
    for d in list(ua.values()) + list(ag.values()):
        u = user_col(d)
        if u:
            all_users |= set(d[u].dropna())
        if "name" in d.columns:
            raw_names |= set(d["name"].dropna())
    raw_names |= set(ct["Child Name"].dropna()) | set(ct["Household Head Name"].dropna())
    for col in ("Country", "State", "LGA", "Ward", "Health Facility"):
        admin_vals |= set(ct[col].dropna()) if col in ct else set()

map_f = PRIV / "worker_code_map.csv"
old = pd.read_csv(map_f, dtype=str) if map_f.exists() else pd.DataFrame(columns=["username", "worker_code"])
wmap = dict(zip(old.username, old.worker_code))
for u in sorted(all_users - set(wmap)):
    role = next((r for r in ("CDD", "OIC", "WM", "SUP") if str(u).upper().startswith(r)), "USR")
    wmap[u] = f"{role}-{sum(1 for v in wmap.values() if v.startswith(role + '-')) + 1:04d}"
pd.DataFrame(sorted(wmap.items()), columns=["username", "worker_code"]).to_csv(map_f, index=False)
W = lambda s: s.map(lambda v: wmap.get(v, "") if isinstance(v, str) else "")


def pseudo_users(df):
    df = df.copy(); u = user_col(df)
    if u:
        df.insert(list(df.columns).index(u), "worker_code", W(df[u])); df = df.drop(columns=[u])
    drop = [c for c in df.columns if PII_COL_RE.search(c) and c != "worker_code"]
    return df.drop(columns=drop), drop


audit = []
for c in CYCLES:
    f = F[c]; ct = f["ct"].copy(); n0 = len(ct)
    nm = ct["Child Name"].fillna(""); hh = ct["Household Head Name"].fillna("")
    ct["name_blank"] = (nm.str.strip() == "").astype(int)
    ct["name_short"] = (nm.str.strip().str.len() < 3).astype(int)
    ct["name_has_digit"] = nm.str.contains(r"\d", regex=True).astype(int)
    ct["name_has_symbol"] = nm.str.contains(r"[^A-Za-z\s.\-']", regex=True).astype(int)
    ct["name_single_word"] = (nm.str.strip().str.split().str.len().fillna(0) < 2).astype(int)
    ct["name_length"] = nm.str.strip().str.len()
    ct["hh_head_blank"] = (hh.str.strip() == "").astype(int)
    ct["hh_head_short"] = (hh.str.strip().str.len() < 3).astype(int)
    ct["name_equals_head"] = ((nm.str.strip().str.lower() == hh.str.strip().str.lower()) & (nm.str.strip() != "")).astype(int)

    lat = pd.to_numeric(ct["latitude"], errors="coerce"); lon = pd.to_numeric(ct["longitude"], errors="coerce")
    status = np.select([lat.isna() | lon.isna(), (lat == 0) | (lon == 0), (lat.abs() > 90) | (lon.abs() > 180),
                        ~lat.between(BBOX[0], BBOX[1]) | ~lon.between(BBOX[2], BBOX[3])],
                       ["blank", "zero", "invalid", "outside_state"], "ok")
    ct["gps_status"] = status; ok = status == "ok"
    dec = ct["latitude"].fillna("").str.split(".").str[1].str.len().fillna(0)
    ct["gps_low_precision"] = (ok & (dec <= 3)).astype(int)
    med = pd.DataFrame({"hf": ct["Health Facility"], "lat": lat.where(ok), "lon": lon.where(ok)}).groupby("hf")[["lat", "lon"]].median()
    mlat = ct["Health Facility"].map(med["lat"]); mlon = ct["Health Facility"].map(med["lon"])
    p1, p2 = np.radians(lat), np.radians(mlat)
    a = np.sin((p2 - p1) / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(np.radians(mlon - lon) / 2) ** 2
    ct["gps_km_from_hf_median"] = np.where(ok, (2 * 6371.0 * np.arcsin(np.sqrt(a))).round(1), np.nan)
    ct["gps_point_key"] = (ct["latitude"].fillna("") + "," + ct["longitude"].fillna("")).where(ok, "").map(lambda v: h("P", v, 10))
    hhk = hh.map(norm_name) + "|" + ct["Username"].fillna("") + "|" + lat.round(4).astype(str) + "," + lon.round(4).astype(str)
    ct["household_key"] = hhk.where(ok & (hh.str.strip() != ""), "").map(lambda v: h("H", v, 10))
    dk = nm.map(norm_name) + "|" + hh.map(norm_name) + "|" + ct["Age"].fillna("") + "|" + ct["Gender"].fillna("") + "|" + ct["LGA"].fillna("") + "|" + ct["Ward"].fillna("")
    ct["name_dup_key"] = dk.map(lambda v: h("N", v, 10))
    ct["id_created_date"] = ct["Individual ID"].map(uuid1_date)
    ct["child_key"] = ct["Individual ID"].map(lambda v: h("C", v))
    ct["record_key"] = ct["Client Reference ID"].map(lambda v: h("R", v))
    ct["worker_code"] = W(ct["Username"])
    ts = pd.to_datetime(ct["Timestamp"], utc=True, errors="coerce")
    ct["timestamp_local"] = (ts + pd.Timedelta(hours=cfg.get("tz_offset_hours", 1))).dt.strftime("%Y-%m-%d %H:%M:%S")

    drop = ["Individual ID", "Client Reference ID", "Username"] + [x for x in ct.columns if PII_COL_RE.search(x)]
    drop = list(dict.fromkeys(drop))
    ct = ct.drop(columns=drop)
    lead = ["child_key", "record_key", "Country", "State", "LGA", "Ward", "Health Facility", "worker_code"]
    ct = ct[[x for x in lead if x in ct.columns] + [x for x in ct.columns if x not in lead]]
    assert len(ct) == n0
    ct.to_csv(OUT / f"IN-1_children_treated_C{c}.csv", index=False, encoding="utf-8")
    parts = f["ct"]["source_part"].value_counts().to_dict()
    audit.append((f"IN-1_children_treated_C{c}.csv", n0, len(ct), f"parts {parts}; removed: {', '.join(drop)}"))
    med.reset_index().rename(columns={"hf": "Health Facility", "lat": "hf_median_lat", "lon": "hf_median_lon"}).round(3).to_csv(OUT / f"IN-8_facility_points_C{c}.csv", index=False)

    for key, name in [("hf", "IN-2_hf_summary"), ("stock", "IN-4_stock")]:
        df = read_all(f["fl"][key]).drop(columns=["source_part"])
        df, dropped = pseudo_users(df)
        df.to_csv(OUT / f"{name}_C{c}.csv", index=False)
        audit.append((f"{name}_C{c}.csv", len(df), len(df), "removed: " + ", ".join(dropped) if dropped else "no personal data; copied"))
    with pd.ExcelWriter(OUT / f"IN-5_user_activity_by_day_C{c}.xlsx", engine="xlsxwriter") as xw:
        for sh, df in f["ua"].items():
            n = len(df); df, dropped = pseudo_users(df); df.to_excel(xw, sheet_name=sh[:31], index=False)
            audit.append((f"IN-5_user_activity_by_day_C{c}.xlsx[{sh}]", n, len(df), "username -> worker_code" + (f"; removed {dropped}" if dropped else "")))
    if f["fl"].get("hf_activity"):
        with pd.ExcelWriter(OUT / f"IN-3_hf_activity_by_day_C{c}.xlsx", engine="xlsxwriter") as xw:
            for sh, df in pd.read_excel(f["fl"]["hf_activity"][0], sheet_name=None, engine="calamine", dtype=str).items():
                pseudo_users(df)[0].to_excel(xw, sheet_name=sh[:31], index=False)
    if f["ag"]:
        sh = next(iter(f["ag"]))
        df, dropped = pseudo_users(f["ag"][sh])
        df.to_csv(OUT / f"IN-3b_user_outcome_agg_C{c}.csv", index=False)
        audit.append((f"IN-3b_user_outcome_agg_C{c}.csv", len(df), len(df), "username -> worker_code; removed: " + ", ".join(dropped)))

# ------------------------------------------------------------------ leak scan (exact cell matches + ID patterns)
leaks = []
uuid_re = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")
bad = (all_users | {n for n in raw_names if len(str(n).strip()) >= 4}) - admin_vals
outs = [(p.name, {"": pd.read_csv(p, dtype=str)}) for p in OUT.glob("*.csv") if not p.name.startswith("_")]
outs += [(p.name, pd.read_excel(p, sheet_name=None, engine="calamine", dtype=str)) for p in OUT.glob("*.xlsx")]
for name, sheets in outs:
    for sh, df in sheets.items():
        obj = df.select_dtypes(include="object")
        for col in obj.columns:
            vals = pd.unique(obj[col].dropna())
            n = sum(1 for v in vals if v in bad or uuid_re.search(v))
            if n:
                leaks.append((f"{name}[{sh}]", col, n))   # column + count only, never the values
pd.DataFrame(audit, columns=["file", "rows_in", "rows_out", "action"]).to_csv(OUT / "_deidentification_audit.csv", index=False)
import shutil; shutil.copy(Path(__file__).resolve().parent.parent / "assets" / "README_deidentified_inputs.md", OUT / "README_how_to_read.md")
with pd.option_context("display.width", 250, "display.max_colwidth", 120):
    print(pd.DataFrame(audit, columns=["file", "rows_in", "rows_out", "action"]).to_string())
print("workers mapped:", len(wmap))
if leaks:
    print("LEAKS FOUND - do not share:", leaks); sys.exit(2)
print("LEAKS: none")
