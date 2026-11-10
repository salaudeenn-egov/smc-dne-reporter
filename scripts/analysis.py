"""Stage 2 - standard SMC cycle report analysis (any state, Cycle 2+).

Usage:  python -X utf8 analysis.py <config.json>
Reads ONLY <out_root>/01_input_deidentified/. Writes 02_analysis/ (workbook, figures, results.pkl).
Metric IDs follow SMC_Report_Standard_Spec_v1 (Metric_Dictionary / Validation_Checks).
"""
import json, pickle, sys
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

cfg = json.load(open(sys.argv[1], encoding="utf-8"))
ROOT = Path(cfg["out_root"])
IN = ROOT / "01_input_deidentified"
OUT = ROOT / "02_analysis"; FIG = OUT / "fig"
FIG.mkdir(parents=True, exist_ok=True)
CUR, PREV = int(cfg["cycle"]), int(cfg["prev_cycle"])
TH = dict(cov=0.90, cov_high=0.95, supply=0.90, workload=400, workload_drv=450, redose_hf=0.12, dup_worker=0.20,
          gps_far=30, stationary=50, recon=0.10, age_min=3, age_max=59, ooh=(20, 6), sex_ratio=105)
TH.update({k: (tuple(v) if isinstance(v, list) else v) for k, v in cfg.get("thresholds", {}).items()})
SUCCESS = "ADMINISTRATION_SUCCESS"; REF = "Children Referred"
R, T = {}, {}


def pct(a, b):
    return np.nan if not b else a / b


def num(s):
    return pd.to_numeric(s, errors="coerce")


# ------------------------------------------------------------------ load
def load(c):
    ct = pd.read_csv(IN / f"IN-1_children_treated_C{c}.csv", dtype=str)
    for k in ["Age", "gps_km_from_hf_median", "name_length"]:
        ct[k] = num(ct[k])
    for k in [x for x in ct.columns if x.startswith(("name_", "hh_head", "gps_low"))]:
        if k not in ("name_dup_key", "name_length"):
            ct[k] = num(ct[k]).fillna(0).astype(int)
    ct["ok"] = ct["Administration Status"] == SUCCESS
    ct["ldt"] = pd.to_datetime(ct["timestamp_local"])
    ct["ldate"] = ct["ldt"].dt.normalize()
    ct["hour"] = ct["ldt"].dt.hour
    hfs = pd.read_csv(IN / f"IN-2_hf_summary_C{c}.csv", dtype=str)
    for k in hfs.columns[4:]:
        hfs[k] = num(hfs[k]).fillna(0)
    st = pd.read_csv(IN / f"IN-4_stock_C{c}.csv", dtype=str)
    for k in st.columns[1:]:
        st[k] = num(st[k]).fillna(0)
    ua = pd.read_excel(IN / f"IN-5_user_activity_by_day_C{c}.xlsx", sheet_name=None, engine="calamine", dtype=str)
    agp = IN / f"IN-3b_user_outcome_agg_C{c}.csv"
    ag = pd.read_csv(agp, dtype=str) if agp.exists() else None
    if ag is not None:
        for k in [x for x in ag.columns if x not in ("healthFacility", "worker_code")]:
            ag[k] = num(ag[k]).fillna(0)
    # user-activity sheet: 'user_activity' if present, else the first sheet
    ua["user_activity"] = ua.get("user_activity", next(iter(ua.values())))
    return dict(ct=ct, hfs=hfs, st=st, ua=ua, ag=ag)


D = {c: load(c) for c in (PREV, CUR)}


# ------------------------------------------------------------------ optional eligibility fields (present in some states' exports)
def _col(df, pat):
    import re as _re
    return next((c for c in df.columns if _re.search(pat, c, _re.I)), None)


def _yes(x):
    return x.fillna("").astype(str).str.strip().str.upper().isin(["YES", "Y", "TRUE", "1"])


_c = D[CUR]["ct"]
_sp, _sick, _fev = _col(_c, r"sp.?/?ctx|28 days"), _col(_c, r"child sick"), _col(_c, r"fever")
_ft, _se = _col(_c, r"first.?time"), _col(_c, r"side.?effect")
ELIG = {
    "spctx": (lambda s: int(_yes(s[_sp]).sum())) if _sp else None,
    "sick": (lambda s: int((_yes(s[_sick]) | (_yes(s[_fev]) if _fev else False)).sum())) if _sick else None,
    "skip": (lambda s: int((_yes(s[_ft]) & s[_se].fillna("").astype(str).str.strip().ne("")).sum())) if (_ft and _se) else None,
}
for _k, _cn in (("spctx", _sp), ("sick", _sick), ("skip", _ft)):
    if ELIG[_k] and _cn not in D[PREV]["ct"].columns:
        _f = ELIG[_k]; ELIG[_k] = (lambda f, cn: (lambda s: f(s) if cn in s.columns else "Not in export"))(_f, _cn)


# ------------------------------------------------------------------ per-cycle core metrics
def core(c):
    d = D[c]; ct, hfs, st = d["ct"], d["hfs"], d["st"]
    s = ct[ct.ok]
    m = {}
    m["R01"] = len(ct); m["R02"] = ct.child_key.nunique(); m["R03"] = s.child_key.nunique(); m["R04"] = len(s)
    m["R05"] = hfs["Registered Count"].sum()
    m["R06"] = pct(m["R04"], m["R05"]); m["R07"] = pct(m["R03"], m["R05"])
    nonref = ct[ct["Administration Status"] != REF]
    m["R08"] = pct(len(s), len(nonref))
    first = s.sort_values("ldt").drop_duplicates("child_key")
    m["spaq1"] = (first.productName == "SPAQ 1").sum(); m["spaq2"] = (first.productName == "SPAQ 2").sum()
    vc = ct["Administration Status"].value_counts()
    for k in ["VISITED", "INELIGIBLE", "BENEFICIARY_ABSENT", "BENEFICIARY_REFUSED", "BENEFICIARY_DIED", "BENEFICIARY_MIGRATED", REF]:
        m[k] = int(vc.get(k, 0))
    m["contacts"] = len(nonref)
    m["R10"] = (hfs["Registered Count"] - hfs["TOTAL"]).clip(lower=0).sum()
    # tempo
    day = s.groupby("ldate").size()
    peak = day.max(); win = day[day >= 0.1 * peak].index
    m["win_start"], m["win_end"] = win.min(), win.max()
    m["window_days"] = (m["win_end"] - m["win_start"]).days + 1
    m["T02"] = pct(day[(day.index >= m["win_start"]) & (day.index <= m["win_end"])].sum(), len(s))
    m["peak_date"], m["peak_doses"] = day.idxmax(), int(peak)
    m["T04_n"] = int(day.get(m["win_start"], 0)); m["T04"] = pct(m["T04_n"], len(s))
    post = s[s.ldate > m["win_end"]]; pre = ct[ct.ldate < m["win_start"]]
    m["T05_n"] = len(post); m["T05"] = pct(len(post), len(s))
    m["tail_days"] = (post.ldate.max() - m["win_end"]).days if len(post) else 0
    m["T06_n"] = len(pre)
    m["first_date"], m["last_date"] = ct.ldate.min(), ct.ldate.max()
    lo, hi = TH["ooh"]
    m["T10_n"] = int(((s.hour >= lo) | (s.hour < hi)).sum()); m["T10"] = pct(m["T10_n"], len(s))
    m["daily"] = day
    # repeat dosing
    rc = s.child_key.value_counts()
    m["S01"] = int((rc > 1).sum()); m["S02"] = int((rc - 1).clip(lower=0).sum())
    # workers
    m["workers_active"] = s.worker_code.nunique()
    m["accounts"] = d["ua"]["user_activity"].worker_code.nunique()
    m["F01"] = pct(m["workers_active"], m["accounts"])
    per = s.groupby("worker_code").child_key.nunique()
    m["F03_med"], m["F03_p10"], m["F03_p90"] = per.median(), per.quantile(.1), per.quantile(.9)
    # stock
    m["K01"] = st["Stock Sent by State to HF"].sum()
    m["K02_recv"] = st["Stock Received by HF from State"].sum(); m["K02_rej"] = st["Stock Rejected by HF from State"].sum()
    m["K02_rej_pct"] = pct(m["K02_rej"], m["K01"])
    m["K12"] = pct(m["K01"], m["R05"])
    m["hfs_n"] = hfs["Health Facility"].nunique(); m["lga_n"] = hfs.LGA.nunique()
    m["ward_n"] = hfs[["LGA", "Ward"]].drop_duplicates().shape[0]
    m["male_pct"] = pct((first.Gender == "MALE").sum(), len(first))
    m["oor"] = int((s.Age > TH["age_max"]).sum())
    m["exact60"] = int((s.Age == 60).sum())
    return m


M = {c: core(c) for c in (PREV, CUR)}
c4, c3 = M[CUR], M[PREV]
ct4, ct3 = D[CUR]["ct"], D[PREV]["ct"]
s4, s3 = ct4[ct4.ok], ct3[ct3.ok]


# ------------------------------------------------------------------ facility table
def hf_table(c):
    d = D[c]; ct = d["ct"]; s = ct[ct.ok]
    hf = d["hfs"].rename(columns={"Registered Count": "registered"}).copy()
    g = s.groupby("Health Facility")
    hf = hf.merge(g.child_key.nunique().rename("dosed"), left_on="Health Facility", right_index=True, how="left")
    hf = hf.merge(g.size().rename("dose_records"), left_on="Health Facility", right_index=True, how="left")
    hf = hf.merge(ct[ct["Administration Status"] == "VISITED"].groupby("Health Facility").size().rename("redose"),
                  left_on="Health Facility", right_index=True, how="left")
    accts = d["ua"]["user_activity"].groupby("healthFacility").worker_code.nunique().rename("accounts")
    active = g.worker_code.nunique().rename("users_active")
    hf = hf.merge(accts, left_on="Health Facility", right_index=True, how="left").merge(active, left_on="Health Facility", right_index=True, how="left")
    st = d["st"].set_index("Health Facility")
    hf = hf.merge(st, left_on="Health Facility", right_index=True, how="left")
    hf = hf.fillna({"dosed": 0, "dose_records": 0, "redose": 0, "accounts": 0, "users_active": 0})
    hf["coverage"] = hf.dosed / hf.registered.replace(0, np.nan)
    hf["supply_ratio"] = hf["Stock Sent by State to HF"] / hf.registered.replace(0, np.nan)
    hf["gap"] = (hf.registered - hf.dosed).clip(lower=0)
    hf["reg_per_user"] = hf.registered / hf.users_active.replace(0, np.nan)
    hf["redose_rate"] = hf.redose / hf.dose_records.replace(0, np.nan)
    return hf


HF = {c: hf_table(c) for c in (PREV, CUR)}
hf4, hf3 = HF[CUR], HF[PREV]
R["unmatched_hf_raw_vs_summary"] = sorted(set(s4["Health Facility"]) ^ set(hf4["Health Facility"]))
R["unmatched_hf_stock_vs_summary"] = sorted(set(D[CUR]["st"]["Health Facility"]) ^ set(hf4["Health Facility"]))
R["unmatched_hf_c3_vs_c4"] = sorted(set(hf3["Health Facility"]) ^ set(hf4["Health Facility"]))


def band(x, edges=(0.8, 0.9, 0.95, 1.0), labels=("Below 80%", "80-89%", "90-94%", "95-99%", "100% or more")):
    if pd.isna(x):
        return "No data"
    for e, l in zip(edges, labels):
        if x < e:
            return l
    return labels[-1]


BANDS = ["Below 80%", "80-89%", "90-94%", "95-99%", "100% or more"]


def band_table(hf, col):
    t = hf.assign(b=hf[col].map(band)).groupby("b").agg(Facilities=("Health Facility", "count"), Registered=("registered", "sum"), Dosed=("dosed", "sum"))
    t = t.reindex(BANDS).fillna(0)
    t["Coverage"] = t.Dosed / t.Registered.replace(0, np.nan)
    return t.reset_index().rename(columns={"b": "Band"})


# ------------------------------------------------------------------ Section 3: reach
lga = s4.groupby("LGA")
first4 = s4.sort_values("ldt").drop_duplicates("child_key")
reg_lga = hf4.groupby("LGA").registered.sum()
t = pd.DataFrame({"Registered": reg_lga, "Dosed": lga.child_key.nunique()})
t["Coverage"] = t.Dosed / t.Registered
t["Male %"] = first4.groupby("LGA").Gender.apply(lambda g: (g == "MALE").mean())
for lab, stt in [("Redose /1k", "VISITED"), ("Inelig. /1k", "INELIGIBLE"), ("Referred /1k", REF)]:
    t[lab] = ct4[ct4["Administration Status"] == stt].groupby("LGA").size().reindex(t.index).fillna(0) / t.Dosed * 1000
rcl = s4.groupby(["LGA", "child_key"]).size()
t["Repeat /100"] = (rcl > 1).groupby("LGA").sum() / t.Dosed * 100
t["GPS invalid or blank"] = s4[s4.gps_status != "ok"].groupby("LGA").size().reindex(t.index).fillna(0).astype(int)
t["After window %"] = s4[s4.ldate > c4["win_end"]].groupby("LGA").size().reindex(t.index).fillna(0) / lga.size()
T["E3.3"] = t.sort_values("Registered", ascending=False).reset_index()
T["E3.5"] = band_table(hf4, "coverage")
w = hf4.groupby(["LGA", "Ward"]).agg(Registered=("registered", "sum"), Dosed=("dosed", "sum")).reset_index()
w["Coverage"] = w.Dosed / w.Registered
T["E3.6"] = w.sort_values("Coverage").head(10)[["Ward", "LGA", "Registered", "Dosed", "Coverage"]]
T["E3.7"] = hf4.sort_values("coverage").head(10)[["Health Facility", "LGA", "registered", "dosed", "coverage", "supply_ratio", "users_active"]]
gap = hf4.gap.sort_values(ascending=False); tot_gap = gap.sum()
R["R09"] = tot_gap
R["R14_top10"] = pct(gap.head(int(np.ceil(len(gap) * .1))).sum(), tot_gap)
R["R14_top20"] = pct(gap.head(int(np.ceil(len(gap) * .2))).sum(), tot_gap)
R["R12"] = int((hf4.coverage < TH["cov"]).sum()); R["R13"] = int((hf4.coverage >= TH["cov_high"]).sum())
R["R12_prev"] = int((hf3.coverage < TH["cov"]).sum())
R["gap_share_below"] = pct(hf4[hf4.coverage < TH["cov"]].gap.sum(), tot_gap)
R["D02_girls_short"] = max(0, round(len(first4) * 100 / (100 + TH["sex_ratio"])) - (first4.Gender == "FEMALE").sum())
a12 = first4[first4.Age.between(12, 59)]
R["D03_heap"] = pct((a12.Age % 12 == 0).sum(), len(a12))

# ------------------------------------------------------------------ Section 4: tempo
wd = pd.date_range(c4["win_start"], c4["win_end"])
wd3 = pd.date_range(c3["win_start"], c3["win_end"])


def cum_by_day(s, days, reg):
    fd = s.sort_values("ldt").drop_duplicates("child_key").groupby("ldate").size()
    out = [pct(fd[fd.index <= d].sum(), reg) for d in days]
    return out, pct(fd.sum(), reg)


cur_cum, cur_fin = cum_by_day(s4, wd, c4["R05"]); prv_cum, prv_fin = cum_by_day(s3, wd3, c3["R05"])
n = max(len(wd), len(wd3))
rows = []
for i in range(n):
    a = prv_cum[i] if i < len(prv_cum) else np.nan; b = cur_cum[i] if i < len(cur_cum) else np.nan
    rows.append((f"Day {i+1}", a, b, (b - a) * 100))
rows.append(("Final", prv_fin, cur_fin, (cur_fin - prv_fin) * 100))
T["E4.2"] = pd.DataFrame(rows, columns=["Day", f"Cycle {PREV} cumulative", f"Cycle {CUR} cumulative", "Difference (pts)"])
rows = []
for L, g in s4.groupby("LGA"):
    cc, fin = cum_by_day(g, wd, reg_lga[L])
    rows.append([L] + cc + [fin, (fin - cc[-1]) * 100])
T["E4.3"] = pd.DataFrame(rows, columns=["LGA"] + [f"Day {i+1}" for i in range(len(wd))] + ["Final", "Catch-up gain (pts)"])
aw = s4.groupby("ldate").worker_code.nunique(); dp = s4.groupby("ldate").size()
aw3 = s3.groupby("ldate").worker_code.nunique(); dp3 = s3.groupby("ldate").size()
rows = []
for i in range(n):
    a = (aw.get(wd[i], np.nan), dp.get(wd[i], np.nan) / aw.get(wd[i], np.nan)) if i < len(wd) else (np.nan, np.nan)
    b = (aw3.get(wd3[i], np.nan), dp3.get(wd3[i], np.nan) / aw3.get(wd3[i], np.nan)) if i < len(wd3) else (np.nan, np.nan)
    rows.append((f"Day {i+1}", a[0], a[1], b[0], b[1]))
T["E4.4"] = pd.DataFrame(rows, columns=["Day", "Active workers", "Doses per worker", "Active workers (prev.)", "Doses per worker (prev.)"])
T["E4.6"] = pd.DataFrame([
    ("Day-1 records and share of cycle", f"{c3['T04_n']:,} ({c3['T04']:.1%})", f"{c4['T04_n']:,} ({c4['T04']:.1%})"),
    ("Records before the window", f"{c3['T06_n']:,}", f"{c4['T06_n']:,}"),
    ("Post-window records (count, share, days)", f"{c3['T05_n']:,} ({c3['T05']:.1%}), {c3['tail_days']} days", f"{c4['T05_n']:,} ({c4['T05']:.1%}), {c4['tail_days']} days"),
], columns=["Metric", f"Cycle {PREV}", f"Cycle {CUR}"])

# ------------------------------------------------------------------ Section 5: outcomes
out_rows = [("Successful administration", c4["R04"]), ("Redose visit (VISITED)", c4["VISITED"]), ("Ineligible", c4["INELIGIBLE"]),
            ("Beneficiary absent", c4["BENEFICIARY_ABSENT"]), ("Beneficiary refused", c4["BENEFICIARY_REFUSED"]),
            ("Beneficiary migrated", c4["BENEFICIARY_MIGRATED"]), ("Beneficiary died", c4["BENEFICIARY_DIED"])]
t = pd.DataFrame(out_rows, columns=["Outcome", "Records"]); t["Share of contacts"] = t.Records / c4["contacts"]
t.loc[len(t)] = ["Referred (separate record)", c4[REF], np.nan]
T["E5.1"] = t
red = hf4[hf4.redose_rate > TH["redose_hf"]]
red3 = hf3.set_index("Health Facility").redose_rate
T["E5.2"] = pd.DataFrame([(r.LGA, r["Health Facility"], red3.get(r["Health Facility"], np.nan), r.redose_rate,
                           "Improved" if red3.get(r["Health Facility"], 0) > r.redose_rate else "Not improved") for _, r in red.iterrows()],
                         columns=["LGA", "Facility", f"Cycle {PREV} rate", f"Cycle {CUR} rate", "Assessment"])
R["max_hf_redose"] = hf4.redose_rate.max(); R["max_hf_redose_name"] = hf4.loc[hf4.redose_rate.idxmax(), "Health Facility"]
sym = lambda ct: ct[ct["Administration Status"] == REF].Symptom.fillna("(blank)").value_counts()
T["E5.3"] = pd.DataFrame({f"Cycle {PREV}": sym(ct3), f"Cycle {CUR}": sym(ct4)}).fillna(0).astype(int).reset_index().rename(columns={"index": "Symptom", "Symptom": "Symptom"})
hhc = s4[s4.household_key.notna() & (s4.household_key != "")].groupby("household_key").child_key.nunique()
R["H06_clusters"] = len(hhc); R["H06_mean"] = hhc.mean(); R["H06_p90"] = hhc.quantile(.9); R["H06_over10"] = int((hhc > 10).sum())

# ------------------------------------------------------------------ Section 6: drivers
def spearmanr(a, b):
    return pd.Series(np.asarray(a)).rank().corr(pd.Series(np.asarray(b)).rank()), np.nan


v = hf4.dropna(subset=["coverage", "supply_ratio"])
R["rank_corr"], R["rank_p"] = spearmanr(v.supply_ratio, v.coverage); R["rank_n"] = len(v)
T["E6.2"] = band_table(hf4, "supply_ratio")
R["cov_low_supply"] = pct(hf4[hf4.supply_ratio < TH["supply"]].dosed.sum(), hf4[hf4.supply_ratio < TH["supply"]].registered.sum())
R["cov_high_supply"] = pct(hf4[hf4.supply_ratio >= .95].dosed.sum(), hf4[hf4.supply_ratio >= .95].registered.sum())
R["n_low_supply"] = int((hf4.supply_ratio < TH["supply"]).sum())
f = hf4.assign(short=(hf4.registered - hf4["Stock Sent by State to HF"]).clip(lower=0))
t = f.groupby("LGA").agg(Registered=("registered", "sum"), Outcome=("TOTAL", "sum"), Dosed=("dosed", "sum"), NotReached=("gap", "sum"), Short=("short", "sum"))
t["No outcome, % of gap"] = (t.Registered - t.Outcome).clip(lower=0) / t.NotReached.replace(0, np.nan)
t["Shortfall, % of gap"] = t.Short / t.NotReached.replace(0, np.nan)
T["E6.3"] = t.reset_index()[["LGA", "Registered", "Outcome", "Dosed", "NotReached", "No outcome, % of gap", "Short", "Shortfall, % of gap"]]


def bandf(x, edges, labels):
    if pd.isna(x):
        return "No active user"
    for e, l in zip(edges, labels):
        if x <= e:
            return l
    return labels[-1]


ul = ["Up to 400", "401-500", "501-600", "Above 600"]
t = hf4.assign(b=hf4.reg_per_user.map(lambda x: bandf(x, (400, 500, 600), ul))).groupby("b").agg(Facilities=("Health Facility", "count"), Registered=("registered", "sum"), Dosed=("dosed", "sum")).reindex(ul).fillna(0)
t["Coverage"] = t.Dosed / t.Registered.replace(0, np.nan); T["E6.4"] = t.reset_index().rename(columns={"b": "Registered per user"})
R["cov_low_load"] = t.loc["Up to 400", "Coverage"]; R["cov_high_load"] = t.loc["Above 600", "Coverage"]
sl = ["Up to 2,000", "2,001-4,000", "4,001-6,000", "Above 6,000"]
t = hf4.assign(b=hf4.registered.map(lambda x: bandf(x, (2000, 4000, 6000), sl))).groupby("b").agg(
    Facilities=("Health Facility", "count"), Registered=("registered", "sum"), Dosed=("dosed", "sum"),
    Sent=("Stock Sent by State to HF", "sum"), Below=("coverage", lambda x: int((x < TH["cov"]).sum()))).reindex(sl).fillna(0)
t["Coverage"] = t.Dosed / t.Registered.replace(0, np.nan); t["Supply ratio"] = t.Sent / t.Registered.replace(0, np.nan)
T["E6.5"] = t.reset_index().rename(columns={"b": "Registered per facility"})[["Registered per facility", "Facilities", "Registered", "Dosed", "Coverage", "Supply ratio", "Below"]]
below = hf4[hf4.coverage < TH["cov"]].copy()
below["driver"] = np.select([(below.supply_ratio < TH["supply"]) & (below.reg_per_user > TH["workload_drv"]), below.supply_ratio < TH["supply"], below.reg_per_user > TH["workload_drv"]],
                            ["Supply and workload", "Supply shortfall", "Workload"], "Other (check local causes)")
dl = ["Supply and workload", "Supply shortfall", "Workload", "Other (check local causes)"]
T["E6.6"] = below.groupby("driver").agg(Facilities=("Health Facility", "count"), Registered=("registered", "sum"), Dosed=("dosed", "sum"), NotReached=("gap", "sum")).reindex(dl).fillna(0).reset_index()
R["A36_share"] = pct(f.short.clip(upper=f.gap).sum(), tot_gap)

# ------------------------------------------------------------------ Section 7: field force
ua4 = D[CUR]["ua"]["user_activity"].copy()
accounts = ua4.groupby("lga").worker_code.nunique()
act = s4.groupby("LGA").worker_code.nunique()
t = pd.DataFrame({"Accounts created": accounts, "Active": act}).fillna(0)
t["Active %"] = t.Active / t["Accounts created"]
t["Doses per active worker"] = s4.groupby("LGA").size() / t.Active
t["Registered per user"] = reg_lga / t.Active
t["Facilities <100% active"] = hf4[hf4.users_active < hf4.accounts].groupby("LGA").size().reindex(t.index).fillna(0).astype(int)
T["E7.1"] = t.reset_index().rename(columns={"index": "LGA"})
T["E7.2"] = pd.DataFrame([
    ("Unique worker accounts with doses", c3["workers_active"], c4["workers_active"], pct(c4["workers_active"], c3["workers_active"]) - 1),
    ("Facilities", c3["hfs_n"], c4["hfs_n"], pct(c4["hfs_n"], c3["hfs_n"]) - 1),
    ("Active days (median)", s3.groupby("worker_code").ldate.nunique().median(), s4.groupby("worker_code").ldate.nunique().median(), np.nan)],
    columns=["Metric", f"Cycle {PREV}", f"Cycle {CUR}", "Growth"])
wk = s4.groupby("worker_code").agg(days=("ldate", "nunique"), recs=("record_key", "size"), dup=("Record Type", lambda x: (x == "duplicate").mean()),
                                   kids=("child_key", "nunique"), hfs=("Health Facility", "nunique"), lgas=("LGA", "nunique"))
wk["dband"] = wk.days.map(lambda d: "5+ days" if d >= 5 else f"{d} days" if d >= 2 else "1 day")
T["E7.4"] = wk.groupby("dband").agg(rate=("dup", "mean"), workers=("dup", "size")).reindex(["1 day", "2 days", "3 days", "4 days", "5+ days"]).dropna().reset_index()
rep = s4[s4.child_key.isin(s4.child_key.value_counts().loc[lambda x: x > 1].index)]
extra = rep.sort_values("ldt").groupby("child_key", group_keys=False).apply(lambda g: g.iloc[1:])
T["E7.5"] = extra.groupby("worker_code").size().rename("Repeat SPAQ records").to_frame().join(wk.recs.rename("Total doses")).sort_values("Repeat SPAQ records", ascending=False).head(10).reset_index()
R["F09"] = int((wk.dup > TH["dup_worker"]).sum())
R["F04_50"] = int((wk.kids < 50).sum()); R["F04_100"] = int((wk.kids < 100).sum())
wd_ = s4.groupby(["worker_code", "ldate"]).size()
R["F05_days200"] = int((wd_ > 200).sum()); R["F05_days300"] = int((wd_ > 300).sum()); R["F05_tot600"] = int((wk.recs > 600).sum())
R["F05_max_day"] = int(wd_.max())
R["F06_multi_hf"] = int((wk.hfs > 1).sum()); R["F06_multi_lga"] = int((wk.lgas > 1).sum())
pts = s4[s4.gps_point_key.notna() & (s4.gps_point_key != "")].groupby(["worker_code", "gps_point_key"]).size()
topshare = pts.groupby(level=0).max() / wk.recs
R["F07"] = int(((pts.groupby(level=0).max() >= TH["stationary"]) & (topshare > .5)).sum())
R["F12"] = int((hf4.users_active == 1).sum())
R["F10_nonstandard"] = int(wk.index.to_series().str.startswith("USR-").sum())
R["F11_over"] = int((hf4.reg_per_user > TH["workload"]).sum())
R["F03_perday_med"] = wd_.median()
span = s4.groupby(["worker_code", "ldate"]).ldt.agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
R["F13_span_med"] = span.median()
R["F15_any_repeat_share"] = pct(extra.worker_code.nunique(), len(wk))
top50 = extra.worker_code.value_counts().head(50).sum(); R["F15_top50"] = pct(top50, len(extra))
wm = D[CUR]["ua"].get("warehousemanager_activity", pd.DataFrame(columns=["healthFacility", "worker_code"])).copy()
dcols = [x for x in wm.columns if x[:4] == "2026"]
wm[dcols] = wm[dcols].apply(num).fillna(0)
wm["tot"] = wm[dcols].sum(axis=1)
d1 = c4["win_start"].strftime("%Y-%m-%d"); dm1 = (c4["win_start"] - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
d4 = (c4["win_start"] + pd.Timedelta(days=3)).strftime("%Y-%m-%d")
T["E7.7"] = pd.DataFrame([
    ("OIC accounts / transacting", f"{wm.worker_code.nunique():,} / {(wm.tot > 0).sum():,}"),
    ("Transactions per OIC (median)", f"{wm.loc[wm.tot > 0, 'tot'].median():,.0f}"),
    ("Transactions on day before Day 1 / Day 1 / Day 4", " / ".join(f"{wm[c].sum():,.0f}" if c in wm else "not in report" for c in (dm1, d1, d4))),
    ("Facilities with warehouse activity logged", f"{wm[wm.tot > 0].healthFacility.nunique():,} of {c4['hfs_n']:,}")], columns=["Metric", "Value"])
R["K15_adoption"] = pct((wm.tot > 0).sum(), wm.worker_code.nunique()); R["has_wm"] = len(wm) > 0
R["K15_before_d1"] = wm[dm1].sum() if dm1 in wm else 0

# ------------------------------------------------------------------ Section 8: registry
ids3 = set(s3.child_key)
win_start_c4 = c4["win_start"]
src = first4.assign(src=np.where(first4.child_key.isin(ids3), "Reused from previous cycle",
                                 np.where(pd.to_datetime(first4.id_created_date, errors="coerce") < win_start_c4 - pd.Timedelta(days=3),
                                          "Earlier-season registry (ID created before cycle)", "Newly registered this cycle")))
first3 = s3.sort_values("ldt").drop_duplicates("child_key")
ids_before3 = set(ct3.child_key)  # (C2 not loaded: C3 split uses creation date only)
src3 = first3.assign(src=np.where(pd.to_datetime(first3.id_created_date, errors="coerce") < c3["win_start"] - pd.Timedelta(days=3),
                                  "Earlier-season registry (ID created before cycle)", "Newly registered this cycle"))
order = ["Reused from previous cycle", "Earlier-season registry (ID created before cycle)", "Newly registered this cycle"]
a = src3.src.value_counts().reindex(order).fillna(0).astype(int); b = src.src.value_counts().reindex(order).fillna(0).astype(int)
T["E8.1"] = pd.DataFrame({"Source": order, f"Cycle {PREV}": [f"n/a (Cycle {PREV-1} not in input)"] + list(a.values[1:]), f"Cycle {CUR}": b.values, "Share": (b / b.sum()).values})
R["G01_reused"] = int(b.iloc[0]); R["G01_new"] = int(b.iloc[2]); R["G01_earlier"] = int(b.iloc[1])
R["G01_note_prev"] = f"Cycle {PREV} split uses ID creation date only (Cycle {PREV-1} file not in this input set), so its 'reused' row is folded into 'earlier-season'."
t = src.groupby("LGA").src.value_counts().unstack().reindex(columns=order).fillna(0)
t["Reuse rate"] = t[order[0]] / t.sum(axis=1)
T["E8.3"] = t.reset_index()[["LGA", order[2], order[0], "Reuse rate"]].rename(columns={order[2]: "New registrations", order[0]: "Reused records"})
j = first3.set_index("child_key")[["Gender", "LGA", "Ward", "Health Facility", "worker_code", "Age"]].join(
    first4.set_index("child_key")[["Gender", "LGA", "Ward", "Health Facility", "worker_code", "Age"]], rsuffix="_4", how="inner")
dage = j.Age_4 - j.Age
T["E8.4"] = pd.DataFrame([
    ("Individual ID", f"{len(j):,} matched (hashed key)", "Registry carried over"),
    ("Sex / ward / facility / worker", " / ".join(f"{(j[k] == j[k + '_4']).mean():.1%}" for k in ["Gender", "Ward", "Health Facility", "worker_code"]), "Stable"),
    ("Name and household head", "Not assessable", "Names removed in de-identified input (see Annex E)"),
    ("Age", f"{(dage == 0).mean():.1%} unchanged / {(dage == 1).mean():.1%} +1 month / {(dage == 2).mean():.1%} +2", "Age refreshed" if (dage > 0).mean() > .5 else "Carried forward"),
    ("GPS", "Not assessable", "Coordinates removed in de-identified input (see Annex E)")], columns=["Field", "Match between cycles", "Reading"])
R["G05_unchanged"] = (dage == 0).mean(); R["G05_gap_days"] = (c4["win_start"] - c3["win_start"]).days
ret = first3[["child_key", "LGA"]].assign(again=first3.child_key.isin(set(s4.child_key)))
t = ret.groupby("LGA").agg(prev=("child_key", "size"), again=("again", "sum"))
t["not_again"] = t.prev - t.again; t["Retention"] = t.again / t.prev
t["New in this cycle"] = src[src.src != order[0]].groupby("LGA").size().reindex(t.index).fillna(0).astype(int)
T["E8.5"] = t.reset_index()
R["G06"] = pct(ret.again.sum(), len(ret)); R["G06_n"] = int(ret.again.sum()); R["G06_prev"] = len(ret)
R["G06_lost_top"] = t.not_again.sort_values(ascending=False).head(2).index.tolist()
grp = []
for lab, stt in [("Ineligible only", "INELIGIBLE"), ("Referred", REF), ("Redose visit", "VISITED")]:
    ids = set(ct3[ct3["Administration Status"] == stt].child_key)
    if stt == "INELIGIBLE":
        ids -= ids3
    k = len(ids & set(s4.child_key)); grp.append((lab, len(ids), k, pct(k, len(ids))))
T["E8.6"] = pd.DataFrame(grp, columns=["Previous-cycle group", "Children", "Dosed this cycle", "Share"])
regd = pd.to_datetime(s4.dateOfRegistration, errors="coerce")
R["G04_regdate_diff"] = int((regd.dt.normalize() != s4.ldate).sum())
cd = pd.to_datetime(first4.id_created_date, errors="coerce")
lag = (first4.ldate - cd).dt.days
R["G03_med_days"] = lag.median(); R["G03_p90_days"] = lag.quantile(.9)
newk = src[src.src == order[2]]
R["G07_infant_share"] = pct((newk.Age < 12).sum(), len(newk))
R["G07_first_day"] = newk.ldate.min()
R["G09_same"] = bool((hf3.set_index("Health Facility").registered.reindex(hf4["Health Facility"]).values == hf4.registered.values).all())
R["G09_n_same"] = int((hf3.set_index("Health Facility").registered.reindex(hf4["Health Facility"]).values == hf4.registered.values).sum())

# ------------------------------------------------------------------ Section 9: safety
def rep_metrics(s):
    vc = s.child_key.value_counts(); kids = vc[vc > 1].index
    r = s[s.child_key.isin(kids)].sort_values(["child_key", "ldt"])
    r["prev_w"] = r.groupby("child_key").worker_code.shift(); r["prev_t"] = r.groupby("child_key").ldt.shift()
    e = r[r.prev_w.notna()]
    gaph = (e.ldt - e.prev_t).dt.total_seconds() / 3600
    return dict(children=len(kids), extra=len(e), diff_worker=pct((e.worker_code != e.prev_w).sum(), len(e)),
                same_day=pct((e.ldt.dt.normalize() == e.prev_t.dt.normalize()).sum(), len(e)), within24=pct((gaph <= 24).sum(), len(e)),
                workers=e.worker_code.nunique(), top50=pct(e.worker_code.value_counts().head(50).sum(), len(e)))


RM = rep_metrics(s4); RM3 = rep_metrics(s3)
reused_ids = set(src[src.src == order[0]].child_key)
rr = s4.child_key.value_counts()
R["S04_reused"] = pct((rr[rr.index.isin(reused_ids)] > 1).sum(), len(reused_ids))
R["S04_new"] = pct((rr[~rr.index.isin(reused_ids)] > 1).sum(), (~first4.child_key.isin(reused_ids)).sum())
T["E9.1"] = pd.DataFrame([
    ("Children with 2+ dose records", f"{RM['children']:,} ({pct(RM['children'], c4['R03']):.2%})"),
    ("Extra dose records", f"{RM['extra']:,}"),
    ("Entered by a different worker", f"{RM['diff_worker']:.1%}"),
    ("Same day / within 24 hours", f"{RM['same_day']:.1%} / {RM['within24']:.1%}"),
    ("Workers involved / top-50 share", f"{RM['workers']:,} / {RM['top50']:.1%}"),
    ("Repeat rate: reused vs new children", f"{R['S04_reused']:.2%} / {R['S04_new']:.2%}")], columns=["Metric", "Value"])
el = extra.groupby("LGA").size(); T["E9.2"] = pd.DataFrame({"Extra dose records": el, "Share of LGA doses": el / s4.groupby("LGA").size()}).reset_index()
oor = lambda s: s.groupby("LGA").apply(lambda g: (g.Age > TH["age_max"]).mean())
ex60 = s4[s4.Age > TH["age_max"]].groupby("LGA").apply(lambda g: (g.Age == 60).mean())
T["E9.3"] = pd.DataFrame({f"Cycle {PREV} rate": oor(s3), f"Cycle {CUR} rate": oor(s4), "Exact 60 months share": ex60}).reset_index()
ages = lambda s: s.Age.value_counts().reindex([58, 59, 60, 61, 62]).fillna(0).astype(int)
T["E9.4"] = pd.DataFrame({"Age (months)": [58, 59, 60, 61, 62], f"Cycle {PREV}": ages(s3).values, f"Cycle {CUR}": ages(s4).values})
dk = first4[first4.name_dup_key.notna() & (first4.name_dup_key != "")]
dupg = dk.groupby("name_dup_key").child_key.nunique()
dupkids = set(dk[dk.name_dup_key.isin(dupg[dupg > 1].index)].child_key)
first4["probdup"] = first4.child_key.isin(dupkids)
R["V24_n"] = len(dupkids); R["V24_rate"] = pct(len(dupkids), len(first4))
t = first4.groupby(["Health Facility", "Ward"]).agg(total=("child_key", "size"), dup=("probdup", "mean"), oor=("Age", lambda a: (a > TH["age_max"]).mean()))
T["E9.5"] = t.sort_values("dup", ascending=False).head(10).reset_index()
R["V24_hf_over20"] = int((t.dup > .2).sum())
mism = lambda s: int((((s.productName == "SPAQ 1") & (s.Age >= 12)) | ((s.productName == "SPAQ 2") & (s.Age < 12))).sum())
ib = lambda ct: len(set(ct[ct["Administration Status"] == "INELIGIBLE"].sort_values("ldt").child_key) & set(ct[ct.ok].child_key))
sedrug = lambda ct: int(((ct["Administration Status"] == REF) & (ct.Symptom == "DRUG_SE_PC")).sum())
el_ = lambda k, s: ELIG[k](s) if ELIG[k] else "Not in export"
T["E9.6"] = pd.DataFrame([
    ("Dosed after SP/CTX use in 28 days", el_("spctx", s3), el_("spctx", s4)),
    ("Dosed while sick with fever", el_("sick", s3), el_("sick", s4)),
    ("Ineligible then dosed", ib(ct3), ib(ct4)),
    ("SPAQ product vs age mismatches", mism(s3), mism(s4)),
    ("Drug side-effect referrals", sedrug(ct3), sedrug(ct4))], columns=["Check", f"Cycle {PREV}", f"Cycle {CUR}"])
R["V32"] = ib(ct4); R["V11"] = mism(s4); R["S07"] = sedrug(ct4)

# ------------------------------------------------------------------ Section 10: stock
def stock_pipe(c):
    st = D[c]["st"]; g = lambda k: st[k].sum()
    sent = g("Stock Sent by State to HF")
    return [("Sent by State to HF", f"{sent:,.0f}"),
            ("Received by HF from State", f"{g('Stock Received by HF from State'):,.0f} ({pct(g('Stock Received by HF from State'), sent):.1%})"),
            ("Rejected by HF", f"{g('Stock Rejected by HF from State'):,.0f} ({pct(g('Stock Rejected by HF from State'), sent):.2%})"),
            ("Sent by HF to staff", f"{g('Stock Sent by HF to Staff'):,.0f}"),
            ("Received by staff", f"{g('Stock Received by Staff from HF'):,.0f}"),
            ("Returned by staff / by HF to State", f"{g('Stock Returned to HF by Staff'):,.0f} / {g('Stock Returned by HF to State'):,.0f}"),
            ("Rejected on return (HF / State)", f"{pct(g('Returned Stock Rejected by HF'), g('Stock Returned to HF by Staff')):.1%} / {pct(g('Returned Stock Rejected by State from HF'), g('Stock Returned by HF to State')):.1%}")]


p3, p4 = stock_pipe(PREV), stock_pipe(CUR)
T["E10.1"] = pd.DataFrame([(a[0], a[1], b[1]) for a, b in zip(p3, p4)], columns=["Pipeline stage", f"Cycle {PREV}", f"Cycle {CUR}"])
st4 = hf4.copy()
st4["rej_rate"] = st4["Stock Rejected by HF from State"] / st4["Stock Sent by State to HF"].replace(0, np.nan)
e50 = st4[(st4.rej_rate - .5).abs() <= .005]
T["E10.3"] = e50[["LGA", "Health Facility", "Stock Sent by State to HF", "Stock Received by HF from State", "Stock Rejected by HF from State", "rej_rate"]]
zr = st4[(st4["Stock Received by HF from State"] == 0) & (st4.dose_records > 0)]
T["E10.4"] = zr[["LGA", "Health Facility", "Stock Sent by HF to Staff", "dose_records"]]
R["K04_n"] = len(zr); R["K04_doses"] = zr.dose_records.sum(); R["K03_n"] = len(e50)
st4["util"] = st4.dose_records / st4["Stock Received by HF from State"].replace(0, np.nan)
T["E10.5"] = st4.sort_values("util", ascending=False).head(10)[["LGA", "Health Facility", "Stock Received by HF from State", "dose_records", "util"]]
R["K05_n"] = int((st4.util > 1).sum()); R["K05_over200"] = int((st4.util > 2).sum())
st4["k06"] = st4["Stock Sent by HF to Staff"] - st4["Stock Received by Staff from HF"] - st4["Stock Rejected by Staff from HF"] - st4["Stock In-Transit from HF to Staff"]
held = st4["Stock Received by HF from State"] + st4["Returned Stock Received by HF"]
st4["over"] = (st4["Stock Sent by HF to Staff"] - held).clip(lower=0)
t = st4.groupby("LGA").agg(mis=("k06", lambda x: int((x != 0).sum())), units=("k06", lambda x: x.abs().sum()), over=("over", "sum"), held=("Stock Received by HF from State", "sum"))
t["Over-issuance %"] = t.over / t.held
T["E10.6"] = t.reset_index()[["LGA", "mis", "units", "Over-issuance %"]]
st4["net"] = st4["Stock Received by Staff from HF"] + st4["Stock In-Transit from HF to Staff"] - st4["Stock Returned to HF by Staff"]
st4["diff"] = st4.net - st4.dose_records
R["K11_net"] = st4.net.sum(); R["K11_doses"] = st4.dose_records.sum(); R["K11_diff"] = R["K11_net"] - R["K11_doses"]
R["K11_within"] = int((st4["diff"].abs() <= TH["recon"] * st4.dose_records).sum())
T["E10.7"] = pd.concat([st4.sort_values("diff", ascending=False).head(5), st4.sort_values("diff").head(5)])[["Health Facility", "LGA", "net", "dose_records", "diff"]]
led = 0
for a, b, c_, d_ in [("Stock Sent by State to HF", "Stock Received by HF from State", "Stock Rejected by HF from State", "Stock In-Transit from State to HF"),
                     ("Stock Sent by HF to Staff", "Stock Received by Staff from HF", "Stock Rejected by Staff from HF", "Stock In-Transit from HF to Staff"),
                     ("Stock Returned to HF by Staff", "Returned Stock Received by HF", "Returned Stock Rejected by HF", "Returned Stock In-Transit to HF"),
                     ("Stock Returned by HF to State", "Returned Stock Accepted by State from HF", "Returned Stock Rejected by State from HF", "Returned Stock In-Transit from HF to State")]:
    led += int((st4[a] != st4[b] + st4[c_] + st4[d_]).sum())
R["K14"] = led
R["K13"] = pct(c4["R04"], c4["K01"])
R["K10_rej"] = pct(st4["Returned Stock Rejected by State from HF"].sum(), st4["Stock Returned by HF to State"].sum())
R["K10_transit"] = st4["Returned Stock In-Transit from HF to State"].sum()
R["K09_ret_share"] = pct(st4["Stock Returned to HF by Staff"].sum(), st4["Stock Sent by HF to Staff"].sum())

# ------------------------------------------------------------------ Section 11: validation checks
A, S_ = ct4, s4
withgps = S_[S_.gps_status != "blank"]; okgps = S_[S_.gps_status == "ok"]
hfsum_succ = D[CUR]["hfs"]["ADMINISTRATION_SUCCESS"].sum()
ua_succ = D[CUR]["ag"]["ADMINISTRATION_SUCCESS"].sum() if D[CUR]["ag"] is not None and "ADMINISTRATION_SUCCESS" in D[CUR]["ag"] else None
redo = A[A["Administration Status"] == "VISITED"]; refr = A[A["Administration Status"] == REF]
V = [
    ("V01", "Child name blank", A.name_blank.sum(), len(A), False),
    ("V02", "Child name shorter than 3 characters", A.name_short.sum(), len(A), False),
    ("V03", "Child name contains digits", A.name_has_digit.sum(), len(A), False),
    ("V04", "Child name contains symbols", A.name_has_symbol.sum(), len(A), False),
    ("V05", "Child name identical to household head", A.name_equals_head.sum(), len(A), False),
    ("V06", "Household head blank", A.hh_head_blank.sum(), len(A), False),
    ("V07", "Household head shorter than 3 characters", A.hh_head_short.sum(), len(A), False),
    ("V08", "Age blank or non-numeric", A.Age.isna().sum(), len(A), False),
    ("V09", "Age below minimum (3 months)", (A.Age < TH["age_min"]).sum(), len(A), False),
    ("V10", "Age above maximum (59 months)", (A.Age > TH["age_max"]).sum(), len(A), True),
    ("V11", "SPAQ product inconsistent with age", R["V11"], len(S_), True),
    ("V12", "Gender missing or invalid", (~A.Gender.isin(["MALE", "FEMALE"])).sum(), len(A), False),
    ("V13", "Dose record without GPS", (S_.gps_status == "blank").sum(), len(S_), False),
    ("V14", "GPS outside bounding box", (withgps.gps_status == "outside_state").sum(), len(withgps), False),
    ("V15", "GPS exactly 0 for latitude or longitude", (A.gps_status == "zero").sum(), len(A), False),
    ("V16", "GPS precision of 3 decimals or fewer", withgps.gps_low_precision.sum(), len(withgps), False),
    ("V17", "GPS farther than 30 km from facility median point", (okgps.gps_km_from_hf_median > TH["gps_far"]).sum(), len(okgps), False),
    ("V18", "Product blank on successful administration", S_.productName.isna().sum(), len(S_), False),
    ("V19", "Quantity administered blank or not 1 on success", (num(S_["Quantity Administered"]) != 1).sum(), len(S_), False),
    ("V20", "Redose visit without redose quantity", redo["Redose Quantity Administered"].isna().sum(), len(redo), False),
    ("V21", "Referral without symptom", refr.Symptom.isna().sum(), len(refr), False),
    ("V22", "Duplicate Client Reference ID", A.record_key.duplicated().sum(), len(A), False),
    ("V23", "Child with 2+ SPAQ dose records in the cycle", c4["S01"], c4["R03"], True),
    ("V24", "Probable duplicate registration (same name, head, age, sex, ward)", R["V24_n"], len(first4), False),
    ("V25", "Country or state not the expected value", ((A.Country.str.strip().str.lower() != cfg["country"].lower()) | (A.State.str.strip().str.lower() != cfg["state"].lower())).sum(), len(A), False),
    ("V26", "Cycle Index not the expected value", (num(A["Cycle Index"]) != CUR).sum() if "Cycle Index" in A else None, len(A), False),
    ("V27", "Dose records in the out-of-hours window", c4["T10_n"], len(S_), False),
    ("V28", "Dose records dated before the main window", c4["T06_n"], len(A), False),
    ("V29", "Dose records dated after the main window", c4["T05_n"], len(S_), False),
    ("V30", "Doses despite SP/CTX use in past 28 days", ELIG["spctx"](S_) if ELIG["spctx"] else None, len(S_), True),
    ("V31", "Doses while recorded sick with fever", ELIG["sick"](S_) if ELIG["sick"] else None, len(S_), True),
    ("V32", "Children recorded ineligible and later dosed", R["V32"], c4["R02"], True),
    ("V33", "12-59 month children at exact multiples of 12 months", (a12.Age % 12 == 0).sum(), len(a12), False),
    ("V34", "First-time flag answered yes with side-effect question answered", ELIG["skip"](S_) if ELIG["skip"] else None, len(S_), False),
    ("V35", "Registered children with no recorded outcome", c4["R10"], c4["R05"], False),
    ("V36", "User-activity report vs HF summary doses", abs(ua_succ - hfsum_succ) if ua_succ is not None else None, hfsum_succ, False),
    ("V37", "HF summary vs raw export doses", abs(hfsum_succ - c4["R04"]), c4["R04"], False),
    ("V38", "Facility names that differ across reports or cycles", len(set(R["unmatched_hf_raw_vs_summary"]) | set(R["unmatched_hf_stock_vs_summary"]) | set(R["unmatched_hf_c3_vs_c4"])), c4["hfs_n"], False),
    ("V39", "Registered Count identical to previous cycle (facilities unchanged)", R["G09_n_same"], c4["hfs_n"], False),
    ("V40", "Stock ledger identity violations", R["K14"], c4["hfs_n"] * 4, False),
]


def status(n, base, safety):
    if n is None:
        return "Not assessable"
    r = n / base if base else 0
    if safety and n > 0:
        return "Action"
    return "Pass" if r <= .0005 else ("Watch" if r <= .02 else "Action")


vt = pd.DataFrame([(i, nme, n, b, (n / b if (n is not None and b) else np.nan), status(n, b, sf)) for i, nme, n, b, sf in V],
                  columns=["ID", "Validation check", "Failing records", "Base", "% of base", "Status"])
vt.loc[vt["Failing records"].isna(), "Validation check"] += " (fields not in export)"
T["E11.1"] = vt
R["V_counts"] = vt.Status.value_counts().to_dict()
R["ua_succ"], R["hfsum_succ"] = ua_succ, hfsum_succ
T["E11.4"] = pd.DataFrame([
    ("HF summary", "Raw export", f"{hfsum_succ - c4['R04']:+,.0f} doses", "Extraction time or status mapping" if hfsum_succ != c4["R04"] else "Matches"),
    ("User aggregation report", "HF summary", f"{ua_succ - hfsum_succ:+,.0f} doses" if ua_succ is not None else "report not supplied", ("Definition or extraction difference" if abs(ua_succ - hfsum_succ) > 0 else "Matches") if ua_succ is not None else "-"),
    ("Stock report facilities", "HF summary facilities", f"{len(R['unmatched_hf_stock_vs_summary'])} unmatched", "Name variants" if R["unmatched_hf_stock_vs_summary"] else "All matched")],
    columns=["Source A", "Source B", "Difference", "Likely cause"])
R["V16_rate"] = pct(withgps.gps_low_precision.sum(), len(withgps))
R["V17_n"] = int((okgps.gps_km_from_hf_median > TH["gps_far"]).sum())
R["gps_med_km"] = okgps.gps_km_from_hf_median.median()

# ------------------------------------------------------------------ Section 12: trend
def bandc(x):
    return "Below 80%" if x < .8 else ("80-89%" if x < .9 else "90%+")


tr = hf3[["Health Facility", "coverage"]].merge(hf4[["Health Facility", "LGA", "registered", "coverage", "supply_ratio"]], on="Health Facility", suffixes=("_3", "_4"))
tr["b3"] = tr.coverage_3.map(bandc); tr["b4"] = tr.coverage_4.map(bandc)
x = pd.crosstab(tr.b3, tr.b4).reindex(index=["Below 80%", "80-89%", "90%+"], columns=["Below 80%", "80-89%", "90%+"]).fillna(0).astype(int)
x["All"] = x.sum(axis=1)
T["E12.2"] = x.reset_index().rename(columns={"b3": "Previous band", "Below 80%": "Now below 80%", "80-89%": "Now 80-89%", "90%+": "Now 90%+"})
tr["chg"] = (tr.coverage_4 - tr.coverage_3) * 100
T["E12.3"] = tr.sort_values("chg").head(10)[["Health Facility", "LGA", "registered", "coverage_3", "coverage_4", "chg", "supply_ratio"]]
R["prio_prev"] = int((tr.coverage_3 < TH["cov"]).sum()); R["prio_recovered"] = int(((tr.coverage_3 < TH["cov"]) & (tr.coverage_4 >= TH["cov"])).sum())
bm = pd.DataFrame({
    "OOR age rate": s4.groupby("LGA").Age.apply(lambda a: (a > TH["age_max"]).mean()),
    "Duplicate rate": first4.groupby("LGA").probdup.mean(),
    "Repeat-dose rate": (rcl > 1).groupby("LGA").mean(),
    "Coverage": reg_lga.rdiv(s4.groupby("LGA").child_key.nunique()),
    "Stock over-issuance": T["E10.6"].set_index("LGA")["Over-issuance %"],
    "Retention": T["E8.5"].set_index("LGA").Retention})
T["E12.5"] = bm.T.reset_index().rename(columns={"index": "Metric"})


def X(metric, a, b, better):
    ch = b - a
    if abs(ch) <= 1e-4 * max(abs(a), 1):
        d = "Same"
    elif better == "mixed":
        d = "Mixed"
    else:
        d = "Better" if (ch > 0) == (better == "up") else "Worse"
    return (metric, a, b, ch, d)


R["X01"] = [X("Children registered", c3["R05"], c4["R05"], "up"), X("Children dosed (unique)", c3["R03"], c4["R03"], "up"),
            X("Coverage", c3["R07"], c4["R07"], "up"), X("Facilities below threshold", R["R12_prev"], R["R12"], "down"),
            X("Repeat-dosed children", c3["S01"], c4["S01"], "down"), X("Stock rejected State to HF", c3["K02_rej"], c4["K02_rej"], "down")]

# ------------------------------------------------------------------ figures
plt.rcParams.update({"font.family": "Arial", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
BLUE, GREY, RED, GREEN, ORANGE = "#2F5597", "#A6A6A6", "#C00000", "#548235", "#C55A11"


def save(fig, name):
    fig.tight_layout(); fig.savefig(FIG / name, dpi=200); plt.close(fig)


# Fig 1 coverage by LGA (dumbbell, C2+)
cl3 = hf3.groupby("LGA").dosed.sum() / hf3.groupby("LGA").registered.sum()
cl4 = hf4.groupby("LGA").dosed.sum() / reg_lga
o = cl4.sort_values().index
fig, ax = plt.subplots(figsize=(6.5, 3.2))
for i, L in enumerate(o):
    ax.plot([cl3[L] * 100, cl4[L] * 100], [i, i], color=GREY, zorder=1)
ax.scatter(cl3[o] * 100, range(len(o)), color=GREY, label=f"Cycle {PREV}", zorder=2)
ax.scatter(cl4[o] * 100, range(len(o)), color=BLUE, label=f"Cycle {CUR}", zorder=3)
for i, L in enumerate(o):
    ax.text(cl4[L] * 100 + .3, i + .18, f"{cl4[L]:.1%}  ({int(hf4[hf4.LGA == L].gap.sum()):,} not reached)", fontsize=7.5)
ax.axvline(90, color=RED, ls="--", lw=1); ax.set_yticks(range(len(o))); ax.set_yticklabels(o); ax.set_xlabel("Coverage of registered children (%)")
ax.legend(loc="lower right", frameon=False); save(fig, "fig01_coverage_lga.png")
# Fig 2 concentration
g = hf4.gap.sort_values(ascending=False).values; cs = np.cumsum(g) / g.sum() * 100
fig, ax = plt.subplots(figsize=(6.5, 3)); xs = np.arange(1, len(g) + 1) / len(g) * 100
ax.plot(xs, cs, color=BLUE)
for p in (10, 20):
    k = int(np.ceil(len(g) * p / 100)); ax.axvline(p, color=GREY, ls=":"); ax.text(p + 1, cs[k - 1] - 8, f"top {p}%: {cs[k-1]:.0f}%", fontsize=8)
ax.set_xlabel("Facilities ranked by children not reached (%)"); ax.set_ylabel("Cumulative share of gap (%)"); save(fig, "fig02_concentration.png")
# Fig 3 age band x sex
bins = [3, 6, 12, 24, 36, 48, 60]; lab = ["3-5", "6-11", "12-23", "24-35", "36-47", "48-59"]
f4b = first4[first4.Age.between(3, 59)].assign(b=lambda d: pd.cut(d.Age, bins, right=False, labels=lab))
ab = f4b.groupby(["b", "Gender"], observed=False).size().unstack().reindex(lab)
fig, ax = plt.subplots(figsize=(6.5, 3)); xi = np.arange(len(lab))
ax.bar(xi - .2, ab.get("MALE", 0) / 1000, .4, color=BLUE, label="Male"); ax.bar(xi + .2, ab.get("FEMALE", 0) / 1000, .4, color=ORANGE, label="Female")
ax.set_xticks(xi); ax.set_xticklabels(lab); ax.set_xlabel("Age band (months)"); ax.set_ylabel("Children dosed (thousands)"); ax.legend(frameon=False); save(fig, "fig03_age_sex.png")
# Fig 4 age distribution
ad = first4.Age.value_counts().sort_index()
fig, ax = plt.subplots(figsize=(6.5, 2.8))
ax.bar(ad.index, ad.values / 1000, color=[RED if (a >= 12 and a % 12 == 0) or a > 59 else BLUE for a in ad.index], width=.8)
ax.set_xlabel("Age (months); red = exact years or above 59"); ax.set_ylabel("Children (thousands)"); save(fig, "fig04_age_dist.png")
# Fig 5 daily doses
dd = c4["daily"]; cum = dd.cumsum() / dd.sum() * 100
fig, ax = plt.subplots(figsize=(6.5, 3))
ax.bar(dd.index.strftime("%d %b"), dd.values / 1000, color=[BLUE if c4["win_start"] <= d <= c4["win_end"] else GREY for d in dd.index])
ax2 = ax.twinx(); ax2.plot(dd.index.strftime("%d %b"), cum.values, color=ORANGE, marker="o", ms=3, label=f"Cycle {CUR} cumulative %")
d3 = c3["daily"]; d3w = d3[d3.index >= c3["win_start"]]
c3c = (d3w.cumsum() / d3.sum() * 100).values[:len(dd)]
off = list(dd.index).index(c4["win_start"]) if c4["win_start"] in dd.index else 0
ax2.plot([dd.index.strftime("%d %b")[off + i] for i in range(min(len(c3c), len(dd) - off))], c3c[:len(dd) - off], color=GREY, ls="--", label=f"Cycle {PREV} cumulative % (aligned to Day 1)")
ax2.set_ylim(0, 105); ax2.set_ylabel("Cumulative %"); ax.set_ylabel("Dose records (thousands)"); ax.tick_params(axis="x", rotation=45, labelsize=7)
ax2.legend(loc="center right", frameon=False, fontsize=7); ax2.spines["top"].set_visible(False); save(fig, "fig05_daily.png")
# Fig 6 hour
hh = s4.hour.value_counts().reindex(range(24)).fillna(0)
fig, ax = plt.subplots(figsize=(6.5, 2.8))
ax.bar(hh.index, hh.values / 1000, color=[RED if (h >= 20 or h < 6) else BLUE for h in hh.index]); ax.set_xticks(range(24))
ax.set_xlabel("Local hour (WAT)"); ax.set_ylabel("Dose records (thousands)"); save(fig, "fig06_hour.png")
# Fig 7 coverage vs supply
fig, ax = plt.subplots(figsize=(6.5, 3.4))
ax.scatter(hf4.supply_ratio * 100, hf4.coverage * 100, s=hf4.registered / 60, alpha=.55, color=BLUE, edgecolor="white")
ax.axhline(90, color=RED, ls="--", lw=1); ax.axvline(100, color=GREY, ls=":", lw=1)
ax.set_xlabel("Supply ratio: units sent by State / registered (%)"); ax.set_ylabel("Coverage (%)"); save(fig, "fig07_cov_supply.png")
# Fig 8 per worker histogram
per = s4.groupby("worker_code").child_key.nunique()
fig, ax = plt.subplots(figsize=(6.5, 2.8)); ax.hist(per, bins=40, color=BLUE); ax.axvline(per.median(), color=RED, ls="--")
ax.text(per.median() + 3, ax.get_ylim()[1] * .9, f"median {per.median():.0f}", color=RED, fontsize=8)
ax.set_xlabel("Children dosed per active CDD"); ax.set_ylabel("CDDs"); save(fig, "fig08_per_worker.png")
# Fig 9 registration source by LGA
e83 = src.groupby("LGA").src.value_counts(normalize=True).unstack().reindex(columns=order).fillna(0) * 100
fig, ax = plt.subplots(figsize=(6.5, 3)); left = np.zeros(len(e83))
for col, cc in zip(order, [BLUE, GREY, ORANGE]):
    ax.barh(e83.index, e83[col], left=left, color=cc, label=col); left += e83[col].values
ax.set_xlabel(f"Share of children dosed in Cycle {CUR} (%)"); ax.legend(frameon=False, fontsize=7, loc="lower center", bbox_to_anchor=(.5, 1), ncol=2); save(fig, "fig09_reg_source.png")
# Fig 10 stock tiers
st_ = D[CUR]["st"]; vals = [st_["Stock Sent by State to HF"].sum(), st_["Stock Received by HF from State"].sum(), st_["Stock Sent by HF to Staff"].sum(),
                            st_["Stock Received by Staff from HF"].sum(), c4["R04"], st_["Stock Returned to HF by Staff"].sum(), st_["Stock Returned by HF to State"].sum()]
labs = ["State sent", "HF received", "HF issued", "Staff received", "Administered", "Returned by staff", "Returned to State"]
fig, ax = plt.subplots(figsize=(6.5, 3)); ax.bar(labs, np.array(vals) / 1000, color=[BLUE] * 4 + [GREEN] + [GREY] * 2)
for i, v_ in enumerate(vals):
    ax.text(i, v_ / 1000 + 2, f"{v_/1000:,.1f}k", ha="center", fontsize=7.5)
ax.set_ylabel("Units (thousands)"); ax.tick_params(axis="x", rotation=20, labelsize=8); save(fig, "fig10_stock_tiers.png")
# Fig 11 doses vs net stock
fig, ax = plt.subplots(figsize=(6.5, 3.4)); v_ = st4[(st4.net > 0) & (st4.dose_records > 0)]
ax.scatter(v_.net, v_.dose_records, s=14, color=BLUE, alpha=.7); lim = [min(v_.net.min(), v_.dose_records.min()) * .8, max(v_.net.max(), v_.dose_records.max()) * 1.2]
ax.plot(lim, lim, color=RED, ls="--", lw=1); ax.set_xscale("log"); ax.set_yscale("log")
ax.set_xlabel("Net stock issued to CDDs (units)"); ax.set_ylabel("Dose records"); save(fig, "fig11_net_vs_doses.png")
# Fig 12 dumbbell same as fig1 (C2+ trend) reuse
import shutil; shutil.copy(FIG / "fig01_coverage_lga.png", FIG / "fig12_trend_lga.png")

# ------------------------------------------------------------------ save
R["cfg"] = cfg; R["elig_fields"] = {k: bool(v) for k, v in ELIG.items()}
R["M"] = M; R["RM"] = RM; R["RM3"] = RM3; R["hf4"] = hf4; R["hf3"] = hf3; R["below"] = below; R["src_order"] = order
R["wk_per_med"] = per.median()
with open(OUT / "results.pkl", "wb") as fh:
    pickle.dump((R, T), fh)
with pd.ExcelWriter(OUT / f"{cfg['state']}_SMC_C{CUR}_analysis.xlsx", engine="xlsxwriter") as xw:
    pd.DataFrame([(k, str(v)) for k, v in R.items() if not isinstance(v, (pd.DataFrame, dict, list))], columns=["Metric", "Value"]).to_excel(xw, sheet_name="Scalars", index=False)
    pd.DataFrame({k: {kk: (vv if not isinstance(vv, pd.Series) else "series") for kk, vv in M[k].items()} for k in M}).to_excel(xw, sheet_name="Cycle_metrics")
    for k, v in T.items():
        v.to_excel(xw, sheet_name=k.replace(".", "_"), index=False)
    hf4.to_excel(xw, sheet_name="Facility_C4", index=False)
    pd.DataFrame(R["X01"], columns=["Metric", f"C{PREV}", f"C{CUR}", "Change", "Direction"]).to_excel(xw, sheet_name="X01", index=False)
print("C4:", {k: v for k, v in c4.items() if not isinstance(v, pd.Series)})
print("C3:", {k: v for k, v in c3.items() if not isinstance(v, pd.Series)})
print("V:", R["V_counts"])
print({k: v for k, v in R.items() if not isinstance(v, (pd.DataFrame, dict, list))})
