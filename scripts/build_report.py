"""Stage 3 - build the Word report from the standard template (any state, Cycle 2+).

Usage:  python -X utf8 build_report.py <config.json>

Template styles, page set-up, header/footer come from assets/SMC_Cycle_Report_Template_v1.docx.
All narrative is generated from the analysis results; scorecard ratings, anomalies,
recommendations and stakeholder asks follow the rules in the spec workbook.
"""
import json, pickle, re, sys
from datetime import date
from pathlib import Path
import numpy as np
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

cfg = json.load(open(sys.argv[1], encoding="utf-8"))
SKILL = Path(__file__).resolve().parent.parent
TEMPLATE = Path(cfg.get("template", SKILL / "assets" / "SMC_Cycle_Report_Template_v1.docx"))
ROOT = Path(cfg["out_root"])
R, T = pickle.load(open(ROOT / "02_analysis" / "results.pkl", "rb"))
FIG = ROOT / "02_analysis" / "fig"
OUTDIR = ROOT / "03_report"; OUTDIR.mkdir(exist_ok=True)
CUR, PREV = int(cfg["cycle"]), int(cfg["prev_cycle"])
STATE, COUNTRY, YEAR = cfg["state"], cfg["country"], cfg.get("year", date.today().year)
NAT = cfg.get("national_programme", "National programme")
AGE_MAX = cfg.get("thresholds", {}).get("age_max", 59)
M = R["M"]; c4, c3 = M[CUR], M[PREV]; RM, RM3 = R["RM"], R["RM3"]
NAVY, BLUE, GREY, BODY = RGBColor(0x1F, 0x38, 0x64), RGBColor(0x2E, 0x75, 0xB6), RGBColor(0x59, 0x59, 0x59), RGBColor(0x26, 0x26, 0x26)
SRC_RAW = f"Cycle {CUR} children-treated export (de-identified)"
SRC_HF = "HF summary (IN-2) and children-treated export"
SRC_ST = "Stock report (IN-4)"
VT = T["E11.1"].set_index("ID")


# ------------------------------------------------------------------ formatting
def isnan(x):
    return x is None or (isinstance(x, float) and np.isnan(x))


def N(x):
    return "-" if isnan(x) else (x if isinstance(x, str) else f"{x:,.0f}")


def P(x, d=1):
    return "-" if isnan(x) or isinstance(x, str) else f"{x*100:.{d}f}%"


def F(x, d=1):
    return "-" if isnan(x) else f"{x:,.{d}f}"


def D(ts):
    return pd.Timestamp(ts).strftime("%d %b %Y").lstrip("0")


def Ds(ts):
    return pd.Timestamp(ts).strftime("%d %b").lstrip("0")


def short(hf):
    return re.sub(r"^[a-z]{2} ", "", str(hf))       # strip tenant prefix like "oy "


def names(lst, k=None):
    lst = [str(x) for x in lst][:k] if k else [str(x) for x in lst]
    return lst[0] if len(lst) == 1 else ", ".join(lst[:-1]) + " and " + lst[-1] if lst else "none"


def vstat(i):
    return VT.loc[i, "Status"]


def vfail(i):
    v = VT.loc[i, "Failing records"]; return 0 if isnan(v) else v


# ------------------------------------------------------------------ document helpers
doc = docx.Document(TEMPLATE)
body = doc.element.body
for el in list(body):
    if el.tag != qn("w:sectPr"):
        body.remove(el)
TN = [0]; FN = [0]


def run(p, text, bold=False, size=10, color=BODY):
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size); r.font.name = "Arial"; r.font.color.rgb = color
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    return r


def para(text="", bold_lead=None, size=10, space_after=6, color=BODY):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space_after)
    if bold_lead:
        run(p, bold_lead + " ", bold=True, size=size, color=color)
    if text:
        run(p, text, size=size, color=color)
    return p


def note(text):
    return para(text, size=9, color=GREY)


def bullets(items):
    for lead, text in items:
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Cm(0.6); p.paragraph_format.first_line_indent = Cm(-0.4)
        p.paragraph_format.space_after = Pt(4); run(p, "- ")
        if lead:
            run(p, lead + " ", bold=True)
        run(p, text)


def heading(text, style, size):
    p = doc.add_paragraph(text, style=style)
    for r in p.runs:
        r.font.name = "Arial"; r.font.color.rgb = NAVY; r.font.size = Pt(size); r.bold = True


h1 = lambda t: heading(t, "Heading 1", 16)
h3 = lambda t: heading(t, "Heading 3", 10.5)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), fill); tcPr.append(s)


def borders(cell, color="D0D7E2"):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:color"), color); b.append(e)
    tcPr.append(b)


def caption(kind, title, source=None, base=None):
    ctr = TN if kind == "Table" else FN; ctr[0] += 1
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(10)
    run(p, f"{kind} {ctr[0]}. ", bold=True, size=8.5, color=BLUE)
    tail = title.rstrip(".") + "." + (f" Source: {source}." if source else "") + (f" Base: {base}." if base else "")
    run(p, tail, size=8.5, color=GREY)


FILL = {"Strong": "E2EFDA", "Mixed": "FFF2CC", "Weak": "F8CBAD", "Pass": "E2EFDA", "Watch": "FFF2CC", "Action": "F8CBAD",
        "Better": "E2EFDA", "Worse": "F8CBAD", "Same": "EDEDED", "Not assessable": "EDEDED",
        "High (safety)": "F8CBAD", "High": "F8CBAD", "Medium": "FFF2CC", "Low": "EDEDED"}


def fills_for(rows, col):
    return {(i, col): FILL[r[col]] for i, r in enumerate(rows) if r[col] in FILL}


def table(header, rows, widths, title=None, source=None, base=None, first_bold=False, fills=None):
    t = doc.add_table(rows=1, cols=len(header)); t.alignment = WD_TABLE_ALIGNMENT.LEFT; t.autofit = False
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; shade(c, "1F3864"); borders(c)
        run(c.paragraphs[0], str(h), bold=True, size=8.5, color=RGBColor(255, 255, 255))
    trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            c = cells[i]; c.text = ""; borders(c)
            if ri % 2 == 1:
                shade(c, "F7F9FC")
            if fills and (ri, i) in fills:
                shade(c, fills[(ri, i)])
            run(c.paragraphs[0], str(v), bold=(first_bold and i == 0), size=8.5)
            if i > 0 and isinstance(v, str) and (v[:1].isdigit() or v[:1] in "-+") and len(v) < 22:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    if title:
        caption("Table", title, source, base)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(2)


def figure(fname, title, source):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(FIG / fname), width=Cm(16)); caption("Figure", title, source)


# ------------------------------------------------------------------ derived facts used across sections
hf4, hf3, below = R["hf4"], R["hf3"], R["below"]
win = f"{Ds(c4['win_start'])} - {D(c4['win_end'])}"
stale_n = R["G09_n_same"]; STALE = stale_n >= 0.9 * c4["hfs_n"]
over100 = int((hf4.coverage > 1).sum())
e33 = T["E3.3"]
lga_over = e33[e33.Coverage > 1.2].LGA.tolist()
low4 = [short(x) for x in T["E3.7"]["Health Facility"].head(4)]
rep_lgas = T["E9.2"].sort_values("Share of LGA doses", ascending=False)
other_n = int(T["E6.6"].set_index("driver").loc["Other (check local causes)", "Facilities"])
e44 = T["E4.4"]; dpw_min, dpw_max = e44["Doses per worker"].min(), e44["Doses per worker"].max()
fu = T["E8.6"].set_index("Previous-cycle group")
e125 = T["E12.5"].set_index("Metric"); lgas = list(e125.columns)
rank = (e125.loc["OOR age rate"].rank() + e125.loc["Duplicate rate"].rank() + e125.loc["Repeat-dose rate"].rank() + e125.loc["Retention"].rank(ascending=False))
bench = rank.idxmin()
X01 = R["X01"]; n_better = sum(d == "Better" for *_, d in X01); n_worse = sum(d == "Worse" for *_, d in X01); n_same = len(X01) - n_better - n_worse
oor_share60 = c4["exact60"] / c4["oor"] if c4["oor"] else 0
v13 = VT.loc["V13", "% of base"]
lowsup = hf4[hf4.supply_ratio < 0.9]

# ------------------------------------------------------------------ anomaly engine (Anomaly_Library triggers)
A = []


def add(i, name, ev, sev, action, owner, rec=None, timing="Before next cycle", prio=None):
    A.append(dict(id=i, name=name, ev=ev, sev=sev, action=action, owner=owner, rec=rec, timing=timing,
                  prio=prio or ("High" if sev.startswith("High") else sev)))


if c4["S01"] > 0:
    add("A01", "Same child dosed 2+ times in a cycle", f"{N(c4['S01'])} children, {N(c4['S02'])} extra records; {P(RM['diff_worker'])} by a different CDD", "High (safety)",
        "Block a second administration with a cross-device 'dosed this cycle' check; review cases with pharmacovigilance", "platform",
        ("Add a cross-device 'already dosed this cycle' check at search and before administration.", f"{N(c4['S01'])} repeat-dosed children; {P(RM['diff_worker'])} by a different CDD."))
if c4["oor"] > 0:
    add("A03", f"Out-of-range age (above {AGE_MAX} months)", f"{N(c4['oor'])} dose records (Cycle {PREV}: {N(c3['oor'])})", "High (safety)",
        "Escalate platform fix for the age ceiling; confirm whether children were dosed; refresh age on reuse", "platform",
        (f"Block dosing above {AGE_MAX} months and derive age from date of birth.", f"{N(c4['oor'])} out-of-range dose records, {P(oor_share60,0)} at exactly 60 months."))
if c4["oor"] > 0 and oor_share60 >= 0.3:
    add("A04", "Exact-60-month age cliff", f"{P(oor_share60,0)} of out-of-range records at exactly 60 months", "High", "Platform defect: derive age from date of birth; report by LGA", "platform")
if R["V32"] > 0:
    add("A26", "Children recorded ineligible and later dosed", f"{N(R['V32'])} children", "High (safety)", "Review each case; make ineligibility sticky within the cycle", "platform",
        ("Make an ineligible outcome block later dosing in the same cycle unless a supervisor clears it.", f"{N(R['V32'])} children ineligible then dosed."))
if R["V24_rate"] > 0.02 or R["V24_hf_over20"] > 0:
    add("A02", "Probable duplicate registrations", f"{P(R['V24_rate'],2)} of children; {R['V24_hf_over20']} facilities above 20%", "High",
        "Search-and-match check at registration; facility-level audit", "platform", ("Add search-and-match before new registration.", f"{N(R['V24_n'])} children in probable duplicate groups."))
if c4["R10"] / c4["R05"] > 0.01:
    add("A07", "Registered children with no outcome", f"{N(c4['R10'])} ({P(c4['R10']/c4['R05'])})" + (" under a stale denominator" if STALE else ""), "High",
        "Mandatory outcome and reason; daily exception list", "state", ("Require an outcome for every registered child and send a daily exception list.", f"{N(c4['R10'])} registered children with no outcome."), "During", "Medium" if STALE else "High")
if R["n_low_supply"] > 0 or c4["K12"] < 1:
    add("A08", "Supply below registered children", f"{R['n_low_supply']} facilities under 90% supply; their coverage {P(R['cov_low_supply'])}", "High",
        "Size stock to registration plus buffer; re-supply during the campaign", "supply",
        (f"Top up the {R['n_low_supply']} facilities under 90% supply" + (f" and investigate the {N(c4['K02_rej'])} rejected units." if c4["K02_rej"] else "."), f"Low-supply facilities reached {P(R['cov_low_supply'])} coverage."), "Before / during", "Medium")
for i, cond, name, ev, act in [
        ("A09", R["K03_n"] > 0, "Exact-50% stock rejection pattern", f"{R['K03_n']} facilities", "State supply-chain investigation (two-batch or packaging convention)"),
        ("A10", R["K04_n"] > 0, "Zero-receipt facilities with doses administered", f"{R['K04_n']} facilities, {N(R['K04_doses'])} doses", "Ledger reconciliation at State store; verify receipt entries"),
        ("A11", R["K05_n"] > 0, "Doses exceed stock received", f"{R['K05_n']} facilities ({R['K05_over200']} above 200%)", "Document transfers; fix stock entry timing"),
        ("A13", c4["K02_rej_pct"] > 0.01, "High State-to-facility rejection", f"{P(c4['K02_rej_pct'],2)} of units sent", "Investigate batch quality and delivery procedure")]:
    if cond:
        add(i, name, ev, "High", act, "supply")
if T["E10.6"]["mis"].sum() > 0 or (T["E10.6"]["Over-issuance %"] > 0.01).any():
    add("A12", "HF-to-staff mismatch or over-issuance", f"{int(T['E10.6']['mis'].sum())} facilities with mismatch", "High", "Reconcile sent vs received at the worker tier", "supply")
if not isnan(R["K10_rej"]) and R["K10_rej"] > 0.05:
    add("A14", "High rejection of returned stock at State", P(R["K10_rej"]), "Medium", "Agree return-quantity verification rule", "supply")
if len(T["E5.2"]):
    add("A15", "High-redose facility", f"{len(T['E5.2'])} facilities above 12%", "Medium", "Facility audit; compare with prior cycle", "state")
if c4["F01"] < 0.9:
    add("A17", "Idle worker accounts", f"Activation {P(c4['F01'])}", "Medium", "Retrain or reallocate; remove unused accounts", "partners")
if R["F05_days200"] > 0 or R["F05_tot600"] > 0:
    add("A18", "Implausible worker volumes or shared accounts", f"{R['F05_days200']} CDD-days above 200; {R['F05_tot600']} CDDs above 600", "Medium", "Shared device / credentials review", "partners")
if R["F07"] > 0:
    add("A19", "Stationary or test-entry accounts", f"{R['F07']} CDDs", "Medium", "Review GPS and supervise accounts", "partners")
if R["F09"] > 0:
    add("A20", "CDD duplicate rate above threshold", f"{R['F09']} CDDs above 20% duplicate-record share", "High", "Targeted retraining by tenure group; supervisory review", "partners",
        (f"Retrain the {R['F09']} CDDs above the 20% duplicate-record share" + (f" and supervise the {R['F07']} stationary accounts." if R["F07"] else "."), f"{R['F09']} CDDs above threshold."), "Before next cycle", "Medium")
if any(vstat(i) in ("Watch", "Action") for i in ("V13", "V14", "V17")):
    add("A21", "Missing or implausible GPS", f"{P(v13,2)} without GPS; {N(R['V17_n'])} more than 30 km from facility", "Medium", "Mandatory GPS; distance and stationary exception report", "platform")
if c4["T04"] < 0.10:
    add("A25", "Day-1 mobilisation failure", f"Day 1 = {P(c4['T04'])} of cycle", "High", "Review activation and communication on launch day", "state")
if vstat("V27") in ("Watch", "Action"):
    add("A22", "Records at night", f"{N(c4['T10_n'])} dose records 20:00-05:59", "Low", "Cross-check device sync logs and clocks", "platform")
if c4["T05"] > 0.03 or c4["tail_days"] > 7:
    add("A24", "Long post-campaign tail", f"{P(c4['T05'])} of doses over {c4['tail_days']} days after the window", "Low", "Enforce close-out; reconcile late sync", "state",
        ("Enforce a two-day close-out and sync deadline after the core window.", f"{P(c4['T05'])} of doses over {c4['tail_days']} days after the window (Cycle {PREV}: {P(c3['T05'])})."), "During", "Medium")
if c4["T06_n"] > 0 or c4["T05_n"] > 0:
    add("A23", "Out-of-window records", f"{N(c4['T06_n'])} before, {N(c4['T05_n'])} after the window", "Low", "Tag and exclude test data; enforce date window", "state")
if any(vstat(f"V0{i}") in ("Watch", "Action") for i in range(1, 8)):
    add("A28", "Child and household name quality", f"{N(vfail('V05'))} child names equal household head", "Low", "Field validation; keep name format constant", "partners")
if vfail("V38") > 0:
    add("A29", "Facility-name identity defects", f"{N(vfail('V38'))} names", "Low", "Maintain a machine-checkable facility registry", "platform")
if vfail("V22") > 0:
    add("A30", "Duplicate technical IDs", N(vfail("V22")), "Low", "Fix ID generation or merge process", "platform")
if vfail("V36") > 0 or vfail("V37") > 0:
    add("A31", "Report vs raw mismatch", f"HF summary {R['hfsum_succ']-c4['R04']:+,.0f} vs raw export", "Medium", "Align definitions and extraction; check against source index", "platform",
        ("Reconcile the HF summary to the raw export and document the counting rule.", f"{R['hfsum_succ']-c4['R04']:+,.0f} doses in the HF summary vs raw."), "After", "Medium")
if STALE:
    add("A33", "Registered count not refreshed", f"{stale_n} of {c4['hfs_n']} facilities unchanged since Cycle {PREV}", "Medium", "Refresh registered counts including new children", "state",
        ("Refresh the Registered Count per facility, including new and returning children.", f"{stale_n} of {c4['hfs_n']} facilities unchanged; coverage {P(c4['R07'])}."), "Before next cycle", "High")
rt = cfg.get("thresholds", {}).get("retention_target")
if rt and R["G06"] < rt:
    add("A35", "Low cohort retention", f"{P(R['G06'])} against target {P(rt,0)}", "High", "Follow up missed children; check device sync", "state")
if R["A36_share"] > 0.5:
    add("A36", "Supply shortfall explains large share of gap", P(R["A36_share"], 0), "High", "Stock plan by facility", "supply")
if R["F12"] > 0:
    add("A37", "Single-user facilities", f"{R['F12']} facilities", "Medium", "Add a back-up user", "partners")
if R["F11_over"] > 0:
    add("A38", "High workload per user", f"{R['F11_over']} facilities above {cfg.get('thresholds', {}).get('workload', 400)} registered per user", "Medium", "Add users where load exceeds threshold", "partners")
if other_n > 0:
    add("A39", "Low-coverage facility without matching driver", f"{other_n} facilities below 90% with adequate supply and workload", "Medium", "Local review of access, conduct and attendance", "lga",
        (f"Run a facility review at the {R['R12']} facilities below 90%, starting with {names(low4)}.", f"They hold {P(R['gap_share_below'],0)} of children not reached."), "Before next cycle", "High")
if R.get("has_wm") and not isnan(R["K15_adoption"]) and R["K15_adoption"] < 0.95:
    add("A40", "Stock module under-adoption", f"{P(R['K15_adoption'])} of OIC accounts transacting", "Medium", "Train OICs; pre-position stock", "partners")
SEV_ORDER = {"High (safety)": 0, "High": 1, "Medium": 2, "Low": 3}
A.sort(key=lambda a: (SEV_ORDER[a["sev"]], a["id"]))

# recommendations: anomaly-linked first, then data-gap ones
OWNER = {"platform": "DIGIT HCM platform team", "state": f"{STATE} State SMC programme", "supply": "Supply-chain partners",
         "partners": "Implementing partners", "lga": "LGA SMC focal persons", "data": "DIGIT HCM / data team"}
RECS = []
for a in A:
    if a["rec"]:
        RECS.append(dict(action=a["rec"][0], ev=a["rec"][1], owner=a["owner"], timing=a["timing"], prio=a["prio"], aid=a["id"]))
if len(rep_lgas) and c4["S01"] > 0:
    top2 = rep_lgas.head(2)
    RECS.append(dict(action=f"Review overlapping CDD territories where repeat doses cluster ({names(top2.LGA.tolist())}).",
                     ev="; ".join(f"{r.LGA} {P(r['Share of LGA doses'])}" for _, r in top2.iterrows()) + " of doses are extra records.", owner="partners", timing="Before next cycle", prio="Medium", aid="A01"))
miss = [k for k, v in R["elig_fields"].items() if not v]
if miss:
    RECS.append(dict(action="Add SP/CTX, sickness/fever and first-time-SPAQ fields to the children-treated export.", ev="V30, V31 or V34 cannot be run.",
                     owner="data", timing="Before next cycle", prio="Medium", aid="-"))
if R["G06"] < 0.95:
    RECS.append(dict(action=f"Follow up Cycle {PREV} children not dosed again, by LGA.", ev=f"{N(R['G06_prev']-R['G06_n'])} children; retention {P(R['G06'])}.",
                     owner="lga", timing="During", prio="Low", aid="G06"))
PR = {"High": 0, "Medium": 1, "Low": 2}
RECS = sorted(RECS, key=lambda r: PR.get(r["prio"], 2))[:12]
for i, r in enumerate(RECS, 1):
    r["no"] = f"R{i}"


def recs_of(*owners):
    return [r for r in RECS if r["owner"] in owners]


# ------------------------------------------------------------------ scorecard rules (Rating_Scales)
def rate(strong, weak):
    return "Weak" if weak else ("Strong" if strong else "Mixed")


safety_any = c4["S01"] > 0 or c4["oor"] > 0 or R["V32"] > 0
n_action = R["V_counts"].get("Action", 0)
SC = [("Reach", f"{N(c4['R03'])} children; {R['R12']} facilities below 90%", rate(R["R12"] == 0 and not STALE, c4["R07"] < 0.8),
       "True coverage cannot be measured while the denominator is stale" if STALE else f"{P(R['gap_share_below'],0)} of the gap sits in facilities below 90%"),
      ("Speed and tempo", f"{P(c4['T02'])} in {c4['window_days']}-day window; tail {P(c4['T05'])} over {c4['tail_days']} days", rate(c4["T02"] >= .95 and c4["T05"] <= .03, c4["T02"] < .85),
       "Late entries delay complete reporting" if c4["T05"] > .03 else "Delivery compressed into the planned window"),
      ("Supply adequacy", f"Supply ratio {P(c4['K12'],0)}; {R['n_low_supply']} facilities under 90%", rate(R["n_low_supply"] == 0, c4["K12"] < .9),
       f"Low-supply facilities reached {P(R['cov_low_supply'],0)} of registered" if R["n_low_supply"] else "No facility short of stock"),
      ("Field force", f"{P(c4['F01'])} active; median {F(c4['F03_med'],0)} children per CDD", rate(c4["F01"] >= .95 and R["F05_days200"] == 0, c4["F01"] < .8),
       "Workload even and within threshold" if R["F11_over"] == 0 else f"{R['F11_over']} facilities above workload threshold"),
      ("Equity", f"Boys {P(c4['male_pct'])} of children dosed; girls shortfall {N(R['D02_girls_short'])}", rate(R["D02_girls_short"] <= .01 * c4["R03"], False),
       "No sign of under-reach of girls" if R["D02_girls_short"] <= .01 * c4["R03"] else "Fewer girls than expected"),
      ("Eligibility and safety", f"{N(c4['oor'])} records above {AGE_MAX} months; {N(c4['S01'])} repeat-dosed children", "Weak" if safety_any else "Strong",
       "Possible harm to children; protocol breach" if safety_any else "No safety exceptions"),
      ("Stock accountability", f"{R['K14']} ledger violations; {R['K11_within']}/{c4['hfs_n']} facilities within 10%",
       rate(R["K14"] == 0 and R["K04_n"] == 0 and c4["K02_rej"] == 0 and R["K11_within"] >= .9 * c4["hfs_n"], R["K14"] > 0 or R["K04_n"] > 0),
       f"{N(c4['K02_rej'])} units rejected on receipt" if c4["K02_rej"] else "Ledger reconciles"),
      ("Data quality", f"{R['V_counts'].get('Pass',0)} Pass, {R['V_counts'].get('Watch',0)} Watch, {n_action} Action", rate(n_action <= 3, n_action >= 12),
       "Action items limit trust in summaries" if n_action > 3 else "Core data sound"),
      ("Registry reuse", f"{P(R['G01_reused']/c4['R03'])} reused; retention {P(R['G06'])}", rate(R["G06"] >= .9, R["G06"] < .7),
       f"{N(R['G06_prev']-R['G06_n'])} Cycle {PREV} children not dosed again")]

# ------------------------------------------------------------------ title block
for txt, sz, col, b in [(f"SMC CAMPAIGN {YEAR}  |  {STATE.upper()} STATE, {COUNTRY.upper()}", 11, BLUE, True), (f"Cycle {CUR} Activity Report", 24, NAVY, True),
                        (f"How far did Cycle {CUR} reach, how safely, and what should change before the next cycle?", 12, GREY, False)]:
    run(para(), txt, bold=b, size=sz, color=col)
for a_, b_ in [("Audience:", cfg.get("audience", f"{NAT}, {STATE} State SMC programme, funders and implementing partners")),
               ("Data:", f"Cycle {CUR} administration export ({N(c4['R01'])} records) and DIGIT HCM reports; compared with Cycle {PREV} ({N(c3['R01'])} records)"),
               ("Period:", f"{D(c4['first_date'])} - {D(c4['last_date'])} (main window {win})"),
               ("Coverage:", f"{c4['lga_n']} LGAs  |  {c4['ward_n']} wards  |  {c4['hfs_n']} health facilities  |  {N(c4['accounts'])} CDD accounts"),
               ("Prepared:", f"{date.today():%B %Y}  |  Version {cfg.get('version', '1.0')}  |  Data cut-off {D(c4['last_date'])}"),
               ("Data protection:", "Prepared from a de-identified copy of the data. No child, household-head or CDD names, no household GPS points and no system IDs appear in this report or its input files. CDDs are shown as account codes (CDD-0001 ...).")]:
    para(b_, bold_lead=a_, size=9.5, space_after=3)
doc.add_page_break()

# ------------------------------------------------------------------ 1
h1("1. Executive summary")
h3("E1.1  Headline KPI tiles")
tiles = [(N(c4["R03"]), "Children dosed with SPAQ"), (P(c4["R07"]), f"Coverage of {N(c4['R05'])} registered children" + ("*" if STALE else "")),
         (P(c4["T02"]), f"Doses in {c4['window_days']}-day main window"), (N(c4["workers_active"]), f"CDDs who recorded work ({P(c4['F01'])} of accounts)"),
         (N(c4["S01"]), "Repeat-dosed children (2+ dose records)"), (N(c4["oor"]), f"Dose records above {AGE_MAX} months"),
         (P(R["G06"]), f"Cycle {PREV} children dosed again (retention)"), (str(R["R12"]), "Facilities below 90% coverage")]
t = doc.add_table(rows=2, cols=4)
for i, (v, l) in enumerate(tiles):
    c = t.rows[i // 4].cells[i % 4]; c.text = ""; shade(c, "EAF3FB" if i < 4 else "F7F9FC"); borders(c)
    run(c.paragraphs[0], v, bold=True, size=16, color=NAVY); c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2 = c.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; run(p2, l, size=8, color=GREY)
if STALE:
    note(f"* The registered denominator was not refreshed for Cycle {CUR} ({stale_n} of {c4['hfs_n']} facilities carry the same Registered Count as Cycle {PREV}); {over100} facilities dosed more children than registered. Read coverage as a volume indicator, not true coverage (section 2.3).")

h3("E1.2  Headline narrative")
d_pts = (c4["R07"] - c3["R07"]) * 100
para(f"{STATE} State reached {N(c4['R03'])} children with SPAQ in Cycle {CUR} across {c4['lga_n']} LGAs, {c4['ward_n']} wards and {c4['hfs_n']} facilities, which is {P(c4['R07'])} of {N(c4['R05'])} registered children. "
     f"This is {abs(d_pts):.1f} points {'up' if d_pts >= 0 else 'down'} from Cycle {PREV} ({N(c3['R03'])} children, {P(c3['R07'])}). {P(c4['T02'])} of dose records fell in the {c4['window_days']}-day main window ({win}), "
     f"peaking on {D(c4['peak_date'])} with {N(c4['peak_doses'])} doses. {N(c4['workers_active'])} of {N(c4['accounts'])} CDD accounts recorded work.")
top = A[0] if A else None
if top and top["sev"] == "High (safety)":
    para(f"The most important finding is a child-safety signal. Dose records for children above {AGE_MAX} months moved from {N(c3['oor'])} in Cycle {PREV} to {N(c4['oor'])} in Cycle {CUR}"
         + (f", and {P(oor_share60,0)} of them sit at exactly 60 months, consistent with ages being carried forward on reused registrations [ASSOCIATION]" if c4["oor"] and oor_share60 >= .3 else "")
         + f". In addition, {N(c4['S01'])} children ({P(c4['S01']/c4['R03'],2)}) have two or more SPAQ dose records; {P(RM['diff_worker'])} of the extra records were entered by a different CDD and {P(RM['same_day'])} on the same day.")
else:
    para(f"The most important finding is the concentration of the gap: {R['R12']} facilities below 90% hold {P(R['gap_share_below'],0)} of the {N(R['R09'])} registered children not reached.")
strong = [s[0].lower() for s in SC if s[2] == "Strong"]; weak = [s[0].lower() for s in SC if s[2] != "Strong"]
para(f"Strong: {names(strong) if strong else 'no dimension rated Strong'}. Needs attention: {names(weak)}. "
     f"Against Cycle {PREV}, of {len(X01)} headline metrics {n_better} improved, {n_worse} worsened and {n_same} stayed the same.")

h3("E1.3  What the data tells us")
chg = (c4["R03"] - c3["R03"]) / c3["R03"]
bullets([
    ("Reach.", f"{N(c4['R03'])} children were dosed, {P(abs(chg))} {'more' if chg >= 0 else 'fewer'} than in Cycle {PREV} ({N(c3['R03'])})."),
    ("The gaps are concentrated.", f"{R['R12']} of {c4['hfs_n']} facilities are below 90% and hold {P(R['gap_share_below'],0)} of the {N(R['R09'])} registered children not reached; the top 10% of facilities hold {P(R['R14_top10'],0)}."),
    ("Supply.", f"The State sent {P(c4['K12'],0)} of registered need overall; the {R['n_low_supply']} facilities below 90% supply reached {P(R['cov_low_supply'])} against {P(R['cov_high_supply'],0)} where supply was 95% or more [ASSOCIATION]."),
    ("Field force.", f"{N(c4['workers_active'])} of {N(c4['accounts'])} accounts recorded doses, median {F(c4['F03_med'],0)} children per CDD (P10-P90 {F(c4['F03_p10'],0)}-{F(c4['F03_p90'],0)})."),
    ("Registry.", f"{P(R['G01_reused']/c4['R03'])} of children dosed were already dosed in Cycle {PREV}; {P(R['G06'])} of Cycle {PREV} children were dosed again."),
    ("Location data.", f"Median distance from the facility median point is {F(R['gps_med_km'])} km; {P(v13,1)} of dose records have no GPS and {N(R['V17_n'])} are more than 30 km away."),
    ("Safety.", f"{N(c4['oor'])} dose records above {AGE_MAX} months (Cycle {PREV}: {N(c3['oor'])}); {N(R['V32'])} children recorded ineligible and later dosed; {N(c4['S01'])} children dosed more than once."),
    ("Reporting.", f"The HF summary shows {N(R['hfsum_succ'])} successful administrations against {N(c4['R04'])} in the raw export ({R['hfsum_succ']-c4['R04']:+,.0f})."),
    ("Stock.", f"Net stock issued to CDDs differs from dose records by {R['K11_diff']:+,.0f} units ({P(R['K11_diff']/R['K11_doses'])}); {R['K11_within']} of {c4['hfs_n']} facilities reconcile within +/-10%; {N(c4['K02_rej'])} units rejected on receipt (Cycle {PREV}: {N(c3['K02_rej'])})."),
])
h3("E1.4  Effectiveness scorecard")
table(["Dimension", "Result", "Rating", "Why it matters"], SC, [3.2, 5.5, 1.8, 6.0], "Effectiveness scorecard", "this report", f"Cycle {CUR}", first_bold=True, fills=fills_for(SC, 2))
h3("E1.5  Cycle-over-cycle headline table")
xr = [(m, P(a) if m == "Coverage" else N(a), P(b) if m == "Coverage" else N(b), f"{ch*100:+.1f} pts" if m == "Coverage" else f"{ch:+,.0f}", d) for m, a, b, ch, d in X01]
table(["Metric", f"Cycle {PREV}", f"Cycle {CUR}", "Change", "Direction"], xr, [5.2, 2.6, 2.6, 2.6, 2.6], "Cycle-over-cycle headline table", SRC_HF, "registered children; children dosed", first_bold=True, fills=fills_for(xr, 4))

POS = []
if c4["F01"] >= .95:
    POS.append(("Full activation.", f"{N(c4['workers_active'])} of {N(c4['accounts'])} CDD accounts active (Cycle {PREV}: {N(c3['workers_active'])})."))
if c4["T02"] >= .9:
    POS.append(("Compressed delivery.", f"{P(c4['T02'])} of doses in the {c4['window_days']}-day window, {F(dpw_min,0)}-{F(dpw_max,0)} doses per CDD per day."))
if c4["S01"] < c3["S01"]:
    POS.append(("Repeat dosing down.", f"{N(c3['S01']-c4['S01'])} fewer repeat-dosed children than Cycle {PREV} ({P((c3['S01']-c4['S01'])/c3['S01'],0)} drop)."))
if R["V11"] == 0:
    POS.append(("Product compliance.", "Zero SPAQ 1 / SPAQ 2 age mismatches."))
if R["K14"] == 0 and R["K04_n"] == 0 and R["K03_n"] == 0:
    POS.append(("Ledger integrity.", f"Stock ledger identity holds everywhere; no zero-receipt or exact-50% facilities; {R['K11_within']} of {c4['hfs_n']} reconcile within 10%."))
if "Ineligible only" in fu.index and fu.loc["Ineligible only", "Children"] and fu.loc["Ineligible only", "Share"] >= .7:
    POS.append(("Follow-up of previous non-delivery.", f"{P(fu.loc['Ineligible only','Share'],0)} of Cycle {PREV} ineligible-only children were dosed in Cycle {CUR}."))
if vfail("V38") == 0 and vfail("V22") == 0:
    POS.append(("Clean master data.", "Facility names match across all reports and both cycles; no duplicate technical IDs."))
POS.append(("Benchmark LGA.", f"{bench}: retention {P(e125.loc['Retention', bench])}, repeat-dose rate {P(e125.loc['Repeat-dose rate', bench])}, out-of-range age {P(e125.loc['OOR age rate', bench])}."))
h3("E1.6  Positive findings summary")
bullets([(None, f"{a} {b}") for a, b in POS[:5]])

# ------------------------------------------------------------------ 2
h1("2. About this report: scope, data and method")
h3("E2.1  Operational overview")
table(["Item", "Value"], [("Campaign period (core window / tail)", f"{win} / post-window records to {D(c4['last_date'])}"),
                          ("LGAs / wards / facilities", f"{c4['lga_n']} / {c4['ward_n']} / {c4['hfs_n']}"),
                          ("CDD accounts created / active", f"{N(c4['accounts'])} / {N(c4['workers_active'])}"),
                          ("Products and age bands", f"SPAQ 1: 3-11 months; SPAQ 2: 12-{AGE_MAX} months"),
                          ("Stock tracking", "Commodity module (State, facility and CDD tiers)" + (" with warehouse-manager activity" if R.get("has_wm") else ""))],
      [6, 10.5], "Operational overview", "IN-1 to IN-5", f"Cycle {CUR}", first_bold=True)
h3("E2.2  Definitions")
table(["Term", "Definition used in this report"], [
    ("Child dosed", f"Unique child key with at least one ADMINISTRATION_SUCCESS record in Cycle {CUR}"),
    ("Registered", "Registered Count in the HF summary (IN-2)" + (f"; not refreshed for Cycle {CUR} at {stale_n} of {c4['hfs_n']} facilities" if STALE else "")),
    ("Coverage", "Children dosed (unique) / registered"), ("Main window", f"Days carrying at least 10% of peak-day doses: {win}"),
    ("Supply ratio", "Units sent by the State to a facility / registered children there"),
    ("Timing", f"Timestamps converted to local time (UTC{cfg.get('tz_offset_hours', 1):+d})"),
    ("Child key / CDD code", f"Keyed pseudonyms replacing system IDs and usernames; the same child or CDD has the same key in Cycles {PREV} and {CUR}")],
    [4, 12.5], "Definitions", "SMC report standard v1, Term_Mapping", None, first_bold=True)
h3("E2.3  Method, comparability and limitations")
lim = []
if STALE:
    lim.append(("Denominator.", f"Coverage is of registered, not eligible, children. The Registered Count is identical to Cycle {PREV} at {stale_n} of {c4['hfs_n']} facilities (G09), while children dosed exceed it at {over100} facilities. Coverage above 100% reflects an outdated registry total; facility rankings remain informative, absolute coverage does not."))
else:
    lim.append(("Denominator.", "Coverage is of registered, not eligible, children; no microplan target was supplied."))
lim += [("Windows.", f"Cycle {PREV} main window {Ds(c3['win_start'])} - {D(c3['win_end'])}; Cycle {CUR} {win}. Definitions unchanged between cycles."),
        ("Nothing removed or imputed.", "All records are kept." + (" Facility names match exactly across reports and cycles (V38 = 0)." if vfail("V38") == 0 else f" {N(vfail('V38'))} facility names differ across reports or cycles (V38).")),
        ("De-identified inputs.", "The analysis ran only on de-identified files. Name, GPS and ID checks were computed before removal and carried as flags; checks needing original values that were not pre-computed are marked Not assessable (Annex E).")]
if miss:
    lim.append(("Not in the export.", "Some eligibility questions (SP/CTX, sickness/fever, first-time SPAQ) are not in this export, so the matching checks cannot be run."))
lim.append(("Association is not cause.", "Supply, workload and age-carry-forward findings are labelled [ASSOCIATION]."))
bullets(lim)

# ------------------------------------------------------------------ 3
h1("3. Reach and coverage")
h3("E3.1  Headline figures")
table(["Indicator", "Value"], [
    ("Records in export", N(c4["R01"])), ("Unique individuals (all status)", N(c4["R02"])), ("Children registered", N(c4["R05"])),
    ("Children dosed with SPAQ", f"{N(c4['R03'])} ({P(c4['R07'])} of registered)"),
    ("SPAQ 1 (3-11 months)", f"{N(c4['spaq1'])} ({P(c4['spaq1']/c4['R03'])})"), (f"SPAQ 2 (12-{AGE_MAX} months)", f"{N(c4['spaq2'])} ({P(c4['spaq2']/c4['R03'])})"),
    ("Dose success rate", P(c4["R08"], 2)), ("Redose visits", f"{N(c4['VISITED'])} ({F(c4['VISITED']/c4['R03']*1000,2)} per 1,000)"),
    ("Ineligible contacts", N(c4["INELIGIBLE"])),
    ("Absent / refused / died / migrated", f"{c4['BENEFICIARY_ABSENT']} / {c4['BENEFICIARY_REFUSED']} / {c4['BENEFICIARY_DIED']} / {c4['BENEFICIARY_MIGRATED']}"),
    ("Referrals", N(c4["Children Referred"])), ("Registered with no recorded outcome", f"{N(c4['R10'])} ({P(c4['R10']/c4['R05'])})")],
    [8, 8.5], "Headline figures", SRC_HF, f"all Cycle {CUR} records", first_bold=True)
figure("fig01_coverage_lga.png", f"Coverage by LGA, Cycle {PREV} (grey) and Cycle {CUR} (blue), with 90% threshold", SRC_HF)
h3("E3.3  LGA performance profile")
rows = [(r.LGA, N(r.Registered), N(r.Dosed), P(r.Coverage), P(r["Male %"]), F(r["Redose /1k"], 2), F(r["Inelig. /1k"], 2), F(r["Referred /1k"], 2), F(r["Repeat /100"], 2), N(r["GPS invalid or blank"]), P(r["After window %"])) for _, r in e33.iterrows()]
table(["LGA", "Registered", "Dosed", "Coverage", "Male %", "Redose /1k", "Inelig. /1k", "Referred /1k", "Repeat /100", "GPS invalid or missing", "After window %"], rows,
      [2.6, 1.6, 1.6, 1.4, 1.2, 1.3, 1.3, 1.3, 1.3, 1.5, 1.4], "LGA performance profile, sorted by registered children", SRC_HF, "per-1,000 and per-100 rates use children dosed")
para(f"Coverage ranges from {P(e33.Coverage.min())} in {e33.loc[e33.Coverage.idxmin(),'LGA']} to {P(e33.Coverage.max())} in {e33.loc[e33.Coverage.idxmax(),'LGA']}; {(e33.Coverage >= .9).sum()} of {c4['lga_n']} LGAs reached the 90% threshold."
     + (f" {names(lga_over)} dosed far more children than registered, which points to outdated registered counts there." if lga_over else ""))
h3("E3.5  Facilities by coverage band")
table(["Coverage band", "Facilities", "Registered", "Dosed", "Coverage"], [(r.Band, N(r.Facilities), N(r.Registered), N(r.Dosed), P(r.Coverage)) for _, r in T["E3.5"].iterrows()],
      [4, 3, 3, 3, 3], "Facilities by coverage band", SRC_HF, f"{c4['hfs_n']} facilities")
h3("E3.6  Lowest-coverage wards (10)")
table(["Ward", "LGA", "Registered", "Dosed", "Coverage"], [(r.Ward, r.LGA, N(r.Registered), N(r.Dosed), P(r.Coverage)) for _, r in T["E3.6"].iterrows()],
      [4.5, 3.5, 2.8, 2.8, 2.8], "Lowest-coverage wards (10)", SRC_HF, f"{c4['ward_n']} wards")
h3("E3.7  Lowest-coverage facilities (10) with supply and users")
table(["Facility", "LGA", "Registered", "Dosed", "Coverage", "Supply ratio", "Users"],
      [(short(r["Health Facility"]), r.LGA, N(r.registered), N(r.dosed), P(r.coverage), P(r.supply_ratio, 0), N(r.users_active)) for _, r in T["E3.7"].iterrows()],
      [5.2, 2.6, 1.8, 1.7, 1.7, 1.8, 1.4], "Lowest-coverage facilities (10) with supply and users", f"{SRC_HF}; {SRC_ST}", f"{c4['hfs_n']} facilities")
figure("fig02_concentration.png", "Concentration of children not reached across facilities", SRC_HF)
figure("fig03_age_sex.png", "Children dosed by age band and sex", SRC_RAW)
figure("fig04_age_dist.png", f"Age distribution by month of age (red = exact years from 12 months, or above {AGE_MAX})", SRC_RAW)
para(f"{R['R12']} facilities are below 90% and hold {P(R['gap_share_below'],0)} of the children not reached; the top 10% of facilities hold {P(R['R14_top10'],0)} and the top 20% hold {P(R['R14_top20'],0)}. "
     f"Boys make up {P(c4['male_pct'])} of children dosed" + (f"; about {N(R['D02_girls_short'])} more girls would be expected" if R["D02_girls_short"] else "; no girls shortfall is expected")
     + f". {P(R['D03_heap'])} of 12-{AGE_MAX} month children are recorded at an exact number of years.")

# ------------------------------------------------------------------ 4
h1("4. Campaign tempo and timeline")
figure("fig05_daily.png", f"Daily dose records (blue = main window) with Cycle {CUR} and Cycle {PREV} cumulative share", SRC_RAW)
h3("E4.2  Cumulative coverage by day")
table(["Day", f"Cycle {PREV} cumulative", f"Cycle {CUR} cumulative", "Difference (pts)"], [(r.Day, P(r.iloc[1]), P(r.iloc[2]), F(r.iloc[3])) for _, r in T["E4.2"].iterrows()],
      [3, 4.5, 4.5, 4.5], "Cumulative coverage by day (unique children / registered)", SRC_HF, "registered children")
h3("E4.3  Cumulative coverage by LGA and day")
e = T["E4.3"]; ncol = len(e.columns)
table(list(e.columns), [[r.LGA] + [P(x) for x in r.iloc[1:-1]] + [F(r.iloc[-1])] for _, r in e.iterrows()],
      [3.4] + [min(2.0, 11.0 / (ncol - 2))] * (ncol - 2) + [2.6], "Cumulative coverage by LGA and day", SRC_HF, "registered children per LGA")
h3("E4.4  Active CDDs and doses per CDD by day")
table(["Day", "Active CDDs", "Doses per CDD", f"Active CDDs (Cycle {PREV})", f"Doses per CDD (Cycle {PREV})"],
      [(r.Day, N(r.iloc[1]), F(r.iloc[2]), N(r.iloc[3]), F(r.iloc[4])) for _, r in e44.iterrows()], [2.5, 3.3, 3.3, 3.6, 3.8], "Active CDDs and doses per CDD by main-window day", SRC_RAW, "dose records")
figure("fig06_hour.png", "Dose records by local hour (red = 20:00-05:59)", SRC_RAW)
h3("E4.6  Launch and tail")
table(["Metric", f"Cycle {PREV}", f"Cycle {CUR}"], [tuple(r) for r in T["E4.6"].values], [6.5, 5, 5], "Launch and tail", SRC_RAW, "dose records")
para(f"{P(c4['T02'])} of doses fall in the main window; {N(c4['T05_n'])} ({P(c4['T05'])}) were recorded after it over {c4['tail_days']} days, against {N(c3['T05_n'])} ({P(c3['T05'])}) over {c3['tail_days']} days in Cycle {PREV}. "
     f"{N(c4['T06_n'])} records pre-date the window. {N(c4['T10_n'])} dose records ({P(c4['T10'],2)}) fall between 20:00 and 05:59; treat these as late sync until verified with device logs. Day 1 carried {P(c4['T04'])} of the cycle (Cycle {PREV}: {P(c3['T04'])}).")

# ------------------------------------------------------------------ 5
h1("5. Household contact outcomes")
h3("E5.1  Contact outcomes")
table(["Outcome", "Records", "Share of contacts"], [(r.Outcome, N(r.Records), P(r["Share of contacts"], 3)) for _, r in T["E5.1"].iterrows()],
      [7, 4, 4], "Contact outcomes", SRC_RAW, f"{N(c4['contacts'])} contact records (referrals counted separately)")
h3("E5.2  Facility redose audit list")
if len(T["E5.2"]):
    table(list(T["E5.2"].columns), [(r.iloc[0], short(r.iloc[1]), P(r.iloc[2]), P(r.iloc[3]), r.iloc[4]) for _, r in T["E5.2"].iterrows()], [3, 5.5, 2.6, 2.6, 2.8], "Facility redose audit list (above 12%)", SRC_RAW, "dose records per facility")
else:
    para(f"No facility exceeds the 12% redose threshold. The highest facility redose rate is {P(R['max_hf_redose'],2)} ({short(R['max_hf_redose_name'])}).")
h3("E5.3  Referral symptom mix")
table(["Symptom", f"Cycle {PREV}", f"Cycle {CUR}"], [(r.iloc[0], N(r.iloc[1]), N(r.iloc[2])) for _, r in T["E5.3"].iterrows()], [6, 4, 4], "Referral symptom mix", SRC_RAW, "referral records")
h3("E5.4  Children per household proxy")
para(f"Using a household key of head name + CDD + GPS at about 11 m (computed before de-identification), {N(R['H06_clusters'])} household clusters received doses, {F(R['H06_mean'],2)} children per cluster on average (P90: {F(R['H06_p90'],0)}). "
     f"{N(R['H06_over10'])} clusters hold more than 10 children; these are likely shared compounds or a common recording point and are worth a spot check.")
para(f"Of {N(c4['contacts'])} contact records, {P(c4['R04']/c4['contacts'],2)} were successful administrations. Ineligible contacts were {N(c4['INELIGIBLE'])} ({F(c4['INELIGIBLE']/c4['R03']*1000,2)} per 1,000); referrals {N(c4['Children Referred'])}; redose visits {N(c4['VISITED'])}.")

# ------------------------------------------------------------------ 6
h1("6. Drivers of performance gaps")
figure("fig07_cov_supply.png", "Facility coverage vs supply ratio (bubble size = registered children)", f"{SRC_HF}; {SRC_ST}")
h3("E6.2  Coverage by supply ratio")
table(["Supply ratio band", "Facilities", "Registered", "Dosed", "Coverage"], [(r.Band, N(r.Facilities), N(r.Registered), N(r.Dosed), P(r.Coverage)) for _, r in T["E6.2"].iterrows()],
      [4, 3, 3, 3, 3], "Coverage by supply ratio", f"{SRC_HF}; {SRC_ST}", f"{c4['hfs_n']} facilities")
h3("E6.3  Funnel and supply shortfall by LGA")
table(["LGA", "Registered", "Outcome recorded", "Dosed", "Not reached", "No outcome, % of gap", "Supply shortfall (units)", "Shortfall, % of gap"],
      [(r.LGA, N(r.Registered), N(r.Outcome), N(r.Dosed), N(r.NotReached), P(r["No outcome, % of gap"], 0), N(r.Short), P(r["Shortfall, % of gap"], 0)) for _, r in T["E6.3"].iterrows()],
      [2.8, 1.9, 2, 1.8, 1.8, 2.1, 2.1, 2], "Funnel and supply shortfall by LGA (not reached summed over facilities below their registered count)", f"{SRC_HF}; {SRC_ST}", "facilities")
h3("E6.4  Coverage by registered children per user")
table(["Registered per user", "Facilities", "Registered", "Dosed", "Coverage"], [(r.iloc[0], N(r.Facilities), N(r.Registered), N(r.Dosed), P(r.Coverage)) for _, r in T["E6.4"].iterrows()],
      [4, 3, 3, 3, 3], "Coverage by registered children per active user", SRC_HF, f"{c4['hfs_n']} facilities")
h3("E6.5  Facility size, supply and coverage")
table(["Registered per facility", "Facilities", "Registered", "Dosed", "Coverage", "Supply ratio", "Below threshold"],
      [(r.iloc[0], N(r.Facilities), N(r.Registered), N(r.Dosed), P(r.Coverage), P(r["Supply ratio"], 0), N(r.Below)) for _, r in T["E6.5"].iterrows()],
      [3.4, 2, 2.2, 2.2, 2, 2.2, 2.4], "Facility size, supply and coverage", f"{SRC_HF}; {SRC_ST}", f"{c4['hfs_n']} facilities")
h3("E6.6  Main driver of facilities below threshold")
table(["Main driver", "Facilities", "Registered", "Dosed", "Not reached"], [(r.driver, N(r.Facilities), N(r.Registered), N(r.Dosed), N(r.NotReached)) for _, r in T["E6.6"].iterrows()],
      [5.5, 2.7, 2.7, 2.7, 2.7], "Main driver of facilities below 90% (supply = ratio under 90%; workload = more than 450 registered per user)", f"{SRC_HF}; {SRC_ST}", f"{R['R12']} facilities below 90%")
para(f"Facility coverage rises with supply ratio (rank correlation {F(R['rank_corr'],2)}, n = {R['rank_n']}). Facilities with supply under 90% reached {P(R['cov_low_supply'])} against {P(R['cov_high_supply'],0)} where supply was 95% or more [ASSOCIATION]. "
     "Where shipments are sized to expected doses, part of this correlation runs from coverage to supply; distance and facility size were not tested separately. "
     + ("Workload is not a driver: every facility has 400 or fewer registered children per active user. " if R["F11_over"] == 0 else f"{R['F11_over']} facilities exceed the workload threshold. ")
     + (f"{other_n} of the {R['R12']} low facilities have adequate supply and workload and need a local review (A39). " if other_n else "")
     + ("Scenario estimates (E6.7) are omitted because the stale denominator would overstate modelled gains." if STALE else "Scenario estimates (E6.7) are not included in this version."))

# ------------------------------------------------------------------ 7
h1("7. Field force: CDD activity")
h3("E7.1  Activation and workload by LGA")
table(["LGA", "Accounts created", "Active", "Active %", "Doses per active CDD", "Registered per user", "Facilities <100% active"],
      [(r.LGA, N(r["Accounts created"]), N(r.Active), P(r["Active %"]), F(r["Doses per active worker"], 0), F(r["Registered per user"], 0), N(r["Facilities <100% active"])) for _, r in T["E7.1"].iterrows()],
      [3, 2.2, 1.8, 1.8, 2.5, 2.5, 2.7], "Activation and workload by LGA", "IN-5 user activity and IN-1", "CDD accounts")
h3("E7.2  Workforce overview")
table(["Metric", f"Cycle {PREV}", f"Cycle {CUR}", "Growth"], [(r.Metric, F(r.iloc[1], 0), F(r.iloc[2], 0), P(r.Growth) if not isnan(r.Growth) else "") for _, r in T["E7.2"].iterrows()],
      [7, 3, 3, 3], "Workforce overview", "IN-1", "CDD accounts with doses")
figure("fig08_per_worker.png", "Children dosed per active CDD", SRC_RAW)
h3("E7.4  Days active vs duplicate rate")
table(["Active days", "Average duplicate rate", "CDDs at this level"], [(r.dband, P(r.rate), N(r.workers)) for _, r in T["E7.4"].iterrows()],
      [4, 5, 5], "Days active vs duplicate rate (share of a CDD's records flagged Record Type = duplicate in the export)", SRC_RAW, "active CDDs")
h3("E7.5  CDDs with the most repeat dosing (account codes)")
table(["CDD code", "Repeat SPAQ records", "Total doses"], [(r.worker_code, N(r.iloc[1]), N(r.iloc[2])) for _, r in T["E7.5"].iterrows()],
      [4, 5, 5], "CDDs with the most repeat dosing (account codes; usernames held only in the private key file)", SRC_RAW, "extra dose records")
h3("E7.6  Credibility and hygiene")
bullets([("Workload.", f"Median {F(c4['F03_med'],0)} children per active CDD (P10-P90 {F(c4['F03_p10'],0)}-{F(c4['F03_p90'],0)}); median {F(R['F03_perday_med'],0)} doses per CDD-day over a median working span of {F(R['F13_span_med'],1)} hours."),
         ("Light-touch and implausible volumes.", f"{R['F04_50']} CDDs dosed fewer than 50 children and {R['F04_100']} fewer than 100; {R['F05_days200']} CDD-days above 200 doses (maximum {R['F05_max_day']}); {R['F05_tot600']} CDDs above 600 in total."),
         ("Territory discipline.", f"{R['F06_multi_hf']} CDDs recorded in more than one facility and {R['F06_multi_lga']} in more than one LGA."),
         ("Stationary entry.", f"{R['F07']} CDDs recorded at least 50 doses at one identical GPS point covering more than half their work."),
         ("Duplicate-record share.", f"{R['F09']} CDDs have more than 20% of their records flagged Record Type = duplicate."),
         ("Account hygiene.", f"{R['F10_nonstandard']} active accounts do not follow the standard CDD/OIC username pattern; {R['F12']} single-user facilities.")])
h3("E7.7  Commodity module adoption")
if R.get("has_wm"):
    table(["Metric", "Value"], [tuple(r) for r in T["E7.7"].values], [8, 8.5], "Commodity module adoption (warehouse-manager activity report)", "IN-5 warehouse-manager tab", "OIC accounts")
else:
    para("The warehouse-manager activity report was not supplied, so commodity module adoption cannot be shown.")
para(f"The user-activity report lists {N(c4['accounts'])} CDD accounts; {N(c4['workers_active'])} ({P(c4['F01'])}) recorded at least one dose. Median workload is {F(c4['F03_med'],0)} children per active CDD.")

# ------------------------------------------------------------------ 8
h1("8. Registration and registry reuse")
h3("E8.1  Registration source of children dosed")
table(["Source", f"Cycle {PREV}", f"Cycle {CUR}", f"Share (Cycle {CUR})"], [(r.Source, N(r.iloc[1]), N(r.iloc[2]), P(r.Share)) for _, r in T["E8.1"].iterrows()],
      [7, 3.5, 3, 3], f"Registration source of children dosed (Cycle {PREV} 'reused' needs the Cycle {PREV-1} file, not in this input set)", f"IN-1 Cycles {PREV} and {CUR}", "children dosed")
figure("fig09_reg_source.png", f"Registration source of children dosed in Cycle {CUR}, by LGA", f"IN-1 Cycles {PREV} and {CUR}")
h3("E8.3  Reuse by LGA")
table(["LGA", "New registrations", "Reused records", "Reuse rate"], [(r.LGA, N(r["New registrations"]), N(r["Reused records"]), P(r["Reuse rate"])) for _, r in T["E8.3"].iterrows()],
      [4.5, 4, 4, 4], "Reuse by LGA", f"IN-1 Cycles {PREV} and {CUR}", f"children dosed in Cycle {CUR}")
h3("E8.4  Consistency of reused fields")
table(["Field", "Match between cycles", "Reading"], [tuple(r) for r in T["E8.4"].values], [4, 6.5, 6], "Consistency of reused fields", f"IN-1 Cycles {PREV} and {CUR}", f"{N(R['G06_n'])} children dosed in both cycles")
h3(f"E8.5  Retention of Cycle {PREV} children")
table(["LGA", f"Cycle {PREV} children", "Dosed again", "Not dosed again", "Retention", "New in this cycle"],
      [(r.LGA, N(r.prev), N(r.again), N(r.not_again), P(r.Retention), N(r["New in this cycle"])) for _, r in T["E8.5"].iterrows()],
      [3.4, 2.6, 2.6, 2.6, 2.4, 2.8], f"Retention of Cycle {PREV} children", f"IN-1 Cycles {PREV} and {CUR}", f"children dosed in Cycle {PREV}")
h3("E8.6  Follow-up of previous non-delivery")
table([f"Cycle {PREV} group", "Children", "Dosed this cycle", "Share"], [(r.iloc[0], N(r.Children), N(r.iloc[2]), P(r.Share)) for _, r in T["E8.6"].iterrows()],
      [6, 3.5, 3.5, 3.5], f"Follow-up of Cycle {PREV} non-delivery", f"IN-1 Cycles {PREV} and {CUR}", "children with that outcome")
h3("E8.7  Timing and date quality")
bullets([("Registration-to-dose time.", f"Median {F(R['G03_med_days'],0)} days between ID creation and dose (P90 {F(R['G03_p90_days'],0)} days)."),
         ("Registration-date quality.", f"{N(R['G04_regdate_diff'])} dose records carry a registration date different from the dose date."),
         ("New children.", f"{N(R['G01_new'])} newly registered, first dosed on {D(R['G07_first_day'])}; {P(R['G07_infant_share'],0)} are infants (3-11 months)." if R["G01_new"] else "No newly registered children."),
         ("Denominator refresh (G09).", f"Registered Count unchanged at {stale_n} of {c4['hfs_n']} facilities."),
         ("Age update on reuse (G05).", f"{P(R['G05_unchanged'])} of matched children kept the same age between cycles {R['G05_gap_days']} days apart.")])
para(f"{N(R['G01_reused'])} of {N(c4['R03'])} children dosed ({P(R['G01_reused']/c4['R03'])}) carry a key already dosed in Cycle {PREV}; {N(R['G01_new'])} ({P(R['G01_new']/c4['R03'])}) were newly registered. "
     f"Of {N(R['G06_prev'])} children dosed in Cycle {PREV}, {N(R['G06_n'])} ({P(R['G06'])}) were dosed again; the {N(R['G06_prev']-R['G06_n'])} not dosed again are concentrated in {names(R['G06_lost_top'])}. "
     "Some will have aged out or moved, and some may have been re-registered under a new key.")

# ------------------------------------------------------------------ 9
h1("9. Safety, eligibility and data-integrity risks to children")
h3("E9.1  Repeat dosing summary")
table(["Metric", "Value"], [tuple(r) for r in T["E9.1"].values], [8, 8.5], "Repeat dosing summary", SRC_RAW, "children dosed")
h3("E9.2  Repeat dosing by LGA")
table(["LGA", "Extra dose records", "Share of LGA doses"], [(r.LGA, N(r.iloc[1]), P(r.iloc[2], 2)) for _, r in T["E9.2"].iterrows()], [5, 5, 5], "Repeat dosing by LGA", SRC_RAW, "dose records per LGA")
h3("E9.3  Out-of-range age by LGA")
table(["LGA", f"Cycle {PREV} rate", f"Cycle {CUR} rate", "Share at exactly 60 months"], [(r.LGA, P(r.iloc[1], 2), P(r.iloc[2], 2), P(r.iloc[3], 0)) for _, r in T["E9.3"].iterrows()],
      [4.5, 3.5, 3.5, 4.5], f"Out-of-range age (above {AGE_MAX} months) by LGA", SRC_RAW, "dose records")
h3("E9.4  Age near the eligibility limit")
table(["Age (months)", f"Cycle {PREV}", f"Cycle {CUR}"], [(str(r.iloc[0]), N(r.iloc[1]), N(r.iloc[2])) for _, r in T["E9.4"].iterrows()], [4, 4, 4], "Dose records at ages 58-62 months", SRC_RAW, "dose records")
h3("E9.5  Probable duplicate registrations by facility")
table(["Facility", "Ward", "Children dosed", "Duplicate rate", "Out-of-range rate"], [(short(r["Health Facility"]), r.Ward, N(r.total), P(r.dup, 2), P(r.oor, 2)) for _, r in T["E9.5"].iterrows()],
      [5.5, 3, 2.6, 2.6, 2.8], "Facilities with the highest probable duplicate registration rate (same name, head, age, sex and ward under different keys)", SRC_RAW, "children dosed per facility")
note(f"Name-based duplicate rules overstate duplication where names are common. State-wide, {N(R['V24_n'])} children ({P(R['V24_rate'],2)}) fall in such groups; {R['V24_hf_over20']} facilities exceed 20%.")
h3("E9.6  Eligibility exceptions and product compliance")
table(["Check", f"Cycle {PREV}", f"Cycle {CUR}"], [(r.iloc[0], N(r.iloc[1]), N(r.iloc[2])) for _, r in T["E9.6"].iterrows()], [8, 4, 4], "Eligibility exceptions and product compliance", SRC_RAW, "children and dose records")
para(f"{N(RM['children'])} children ({P(RM['children']/c4['R03'],2)}) have two or more SPAQ dose records ({N(RM['extra'])} extra); {P(RM['diff_worker'])} of the extra records were entered by a different CDD and {P(RM['within24'])} within 24 hours. "
     f"{N(RM['workers'])} CDDs are involved; the top 50 account for {P(RM['top50'],0)} of extra records.")
if c4["oor"]:
    para(f"{N(c4['oor'])} dose records ({P(c4['oor']/c4['R04'],2)}) are for children above {AGE_MAX} months; {P(oor_share60,0)} sit at exactly 60 months. "
         + ("This indicates an age-cliff: reused records age forward each cycle and dosing is not blocked once a child passes the limit. " if oor_share60 >= .3 else "")
         + "All safety items are rated Action and the case lists should go to pharmacovigilance.")

# ------------------------------------------------------------------ 10
h1("10. Stock management")
h3("E10.1  Pipeline by stage")
table(["Pipeline stage", f"Cycle {PREV}", f"Cycle {CUR}"], [tuple(r) for r in T["E10.1"].values], [6.5, 5, 5], "Stock pipeline by stage (SPAQ units)", SRC_ST, "units")
figure("fig10_stock_tiers.png", f"Units moved at each tier, Cycle {CUR}", f"{SRC_ST}; IN-1")
h3("E10.3  Facilities with exact-50% rejection")
if len(T["E10.3"]):
    table(["LGA", "Facility", "Sent", "Received", "Rejected", "Rate"], [(r.LGA, short(r["Health Facility"]), N(r.iloc[2]), N(r.iloc[3]), N(r.iloc[4]), P(r.iloc[5])) for _, r in T["E10.3"].iterrows()],
          [2.8, 5.4, 2, 2, 2, 2], "Facilities with exact-50% rejection", SRC_ST, "facilities")
else:
    para("No facility rejected exactly 50% of a State shipment (K03 = 0).")
h3("E10.4  Zero-receipt facilities with doses administered")
if len(T["E10.4"]):
    table(["LGA", "Facility", "Stock sent (HF to staff)", "Doses administered"], [(r.LGA, short(r["Health Facility"]), N(r.iloc[2]), N(r.iloc[3])) for _, r in T["E10.4"].iterrows()],
          [3, 6, 3.5, 3.5], "Zero-receipt facilities with doses administered", f"{SRC_ST}; IN-1", "facilities")
else:
    para("No facility administered doses without a recorded State receipt (K04 = 0).")
h3("E10.5  Doses administered vs stock received")
table(["LGA", "Facility", "Stock received", "Doses given", "Utilisation"], [(r.LGA, short(r["Health Facility"]), N(r.iloc[2]), N(r.dose_records), P(r.util)) for _, r in T["E10.5"].iterrows()],
      [2.8, 5.8, 2.6, 2.6, 2.6], "Highest utilisation facilities (doses / stock received from State)", f"{SRC_ST}; IN-1", f"{c4['hfs_n']} facilities")
h3("E10.6  HF-to-staff mismatch")
table(["LGA", "Facilities with mismatch", "Units", "Over-issuance %"], [(r.LGA, N(r.mis), N(r.units), P(r.iloc[3], 2)) for _, r in T["E10.6"].iterrows()], [4.5, 4, 3.5, 4], "HF-to-staff mismatch and over-issuance", SRC_ST, "facilities")
h3("E10.7  Net stock issued vs doses")
table(["Facility", "LGA", "Net issued", "Dosed", "Surplus / short"], [(short(r["Health Facility"]), r.LGA, N(r.net), N(r.dose_records), f"{r['diff']:+,.0f}") for _, r in T["E10.7"].iterrows()],
      [6.2, 3, 2.4, 2.4, 2.5], "Net stock issued to CDDs vs dose records (largest surplus, then largest shortfall)", f"{SRC_ST}; IN-1", f"{c4['hfs_n']} facilities")
figure("fig11_net_vs_doses.png", "Dose records vs net stock issued, facility level (log scale, 1:1 line)", f"{SRC_ST}; IN-1")
para(f"{N(c4['K01'])} units were sent by the State store; {P(c4['K02_recv']/c4['K01'])} were confirmed received and {P(c4['K02_rej_pct'],2)} rejected. "
     f"Net stock issued to CDDs is {N(R['K11_net'])} units against {N(R['K11_doses'])} doses, a difference of {R['K11_diff']:+,.0f} ({P(R['K11_diff']/R['K11_doses'])}); {R['K11_within']} of {c4['hfs_n']} facilities reconcile within +/-10%. "
     f"Ledger identity violations: {R['K14']}. {N(R['K10_transit'])} returned units are still in transit to the State. The stock report carries no dates, so movement timing against Day 1 cannot be shown.")

# ------------------------------------------------------------------ 11
h1("11. Data validation and anomalies")
h3("E11.1  Validation results")
vrows = [(r.ID + "  " + r["Validation check"], N(r["Failing records"]), N(r.Base), P(r["% of base"], 2), r.Status) for _, r in T["E11.1"].iterrows()]
table(["Validation check", "Failing", "Base", "% of base", "Status"], vrows, [8.6, 2, 2.2, 1.8, 2], "Validation results (40 standard checks)", "IN-1 to IN-4", "as stated per check", fills=fills_for(vrows, 4))
h3("E11.2  What passed")
para("Zero or near-zero failures on: " + "; ".join(T["E11.1"][T["E11.1"].Status == "Pass"]["Validation check"]) + ".")
h3("E11.3  Location data")
para(f"{P(1 - VT.loc['V14','% of base'],2)} of points with GPS fall inside the state bounding box; median distance from the facility median point is {F(R['gps_med_km'])} km; "
     f"{N(R['V17_n'])} records lie more than 30 km away; {P(R['V16_rate'],2)} have 3 or fewer decimals; {R['F07']} CDDs show stationary entry. Distances were computed from the original coordinates before de-identification.")
h3("E11.4  Report vs raw reconciliation")
table(["Source A", "Source B", "Difference", "Likely cause"], [tuple(r) for r in T["E11.4"].values], [3.8, 3.8, 3, 6], "Report vs raw reconciliation", "IN-1, IN-2, IN-3b, IN-4", "dose records / facilities")
if R["hfsum_succ"] != c4["R04"]:
    note("The cause of the HF summary vs raw difference is not established from these files (measured, not inferred); check it against the source index.")
h3("E11.5  Anomaly register")
arows = [(a["id"], a["name"], a["ev"], a["sev"], a["action"]) for a in A]
table(["ID", "Anomaly", "Evidence", "Severity", "Recommended action"], arows, [1.1, 3.8, 4.6, 2, 5], "Anomaly register (triggered anomalies, ordered by severity)", "this report", None, fills=fills_for(arows, 3))
h3("E11.6  Findings recurring across LGAs")
nl = c4["lga_n"]; rec_rows = []
e93 = T["E9.3"]; up = int((e93.iloc[:, 2] > e93.iloc[:, 1]).sum())
for cond, lead, text, owner in [
        (up >= 3, f"Out-of-range age rose in {up} of {nl} LGAs.", "State-level pattern.", "DIGIT HCM platform team"),
        (int((T["E9.2"]["Share of LGA doses"] > .005).sum()) >= 3, f"Repeat dosing above 0.5% of doses in {int((T['E9.2']['Share of LGA doses'] > .005).sum())} of {nl} LGAs.", "Programme pattern.", f"{STATE} State SMC programme with the platform team"),
        (int((e33.Coverage > 1).sum()) >= 3, f"Coverage above 100% in {int((e33.Coverage > 1).sum())} of {nl} LGAs.", "Denominator process pattern.", f"{STATE} State SMC programme"),
        (int((e33["After window %"] > .025).sum()) >= 3, f"Post-window share above 2.5% in {int((e33['After window %'] > .025).sum())} of {nl} LGAs.", "Close-out pattern.", f"{STATE} State SMC programme"),
        (int((e33.Coverage < .9).sum()) >= 3, f"Coverage below 90% in {int((e33.Coverage < .9).sum())} of {nl} LGAs.", "Reach pattern.", f"{STATE} State SMC programme")]:
    if cond:
        rec_rows.append((lead, f"{text} Owner: {owner}."))
bullets(rec_rows if rec_rows else [(None, "No finding recurs in three or more LGAs.")])
para(f"{R['V_counts'].get('Pass',0)} checks pass, {R['V_counts'].get('Watch',0)} are on watch, {n_action} need action and {R['V_counts'].get('Not assessable',0)} cannot be run on this export.")

# ------------------------------------------------------------------ 12
h1("12. Trend, follow-through and benchmarking")
figure("fig12_trend_lga.png", f"Coverage by LGA: Cycle {PREV} vs Cycle {CUR}", SRC_HF)
h3("E12.2  Facility movement between bands")
table([f"Cycle {PREV} band", "Now below 80%", "Now 80-89%", "Now 90%+", "All"], [(r.iloc[0], N(r.iloc[1]), N(r.iloc[2]), N(r.iloc[3]), N(r.iloc[4])) for _, r in T["E12.2"].iterrows()],
      [4, 3, 3, 3, 3], "Facility movement between coverage bands", SRC_HF, f"{c4['hfs_n']} facilities")
h3("E12.3  Largest facility declines")
e = T["E12.3"]
table(["Facility", "LGA", "Registered", f"Cycle {PREV}", f"Cycle {CUR}", "Change (pts)", "Supply ratio"],
      [(short(r["Health Facility"]), r.LGA, N(r.registered), P(r.coverage_3), P(r.coverage_4), F(r.chg), P(r.supply_ratio, 0)) for _, r in e.iterrows()],
      [5, 2.8, 1.8, 1.7, 1.7, 1.8, 1.8], "Largest facility declines", SRC_HF, f"{c4['hfs_n']} facilities")
endlow = [short(x) for x in e[e.coverage_4 < .9]["Health Facility"]]
note(f"Declines ending below 90%: {names(endlow)}." if endlow else "None of the largest declines ends below 90%.")
h3("E12.4  Follow-through on previous recommendations")
para(f"The Cycle {PREV} recommendation list is not part of this input set, so follow-through is scored on data-visible topics only.")
ST = {"Better": "Strong", "Same": "Mixed", "Worse": "Weak"}
xd = {m: d for m, a, b, ch, d in X01}
fr = [("Reduce repeat dosing", f"Repeat-dosed children {N(c3['S01'])} -> {N(c4['S01'])}", ST.get(xd["Repeat-dosed children"], "Mixed")),
      ("Lift facilities below 90%", f"{R['prio_recovered']} of {R['prio_prev']} Cycle {PREV} low facilities now at 90%+; total below 90% {R['R12_prev']} -> {R['R12']}", ST.get(xd["Facilities below threshold"], "Mixed")),
      ("Control out-of-range age", f"Dose records above {AGE_MAX} months {N(c3['oor'])} -> {N(c4['oor'])}", "Strong" if c4["oor"] < c3["oor"] else ("Mixed" if c4["oor"] == c3["oor"] else "Weak")),
      ("Keep stock rejection low", f"Rejected State-to-HF units {N(c3['K02_rej'])} -> {N(c4['K02_rej'])}", ST.get(xd["Stock rejected State to HF"], "Mixed"))]
table(["Topic", "What happened this cycle", "Status"], fr, [5, 8.5, 3], "Follow-through on data-visible topics", "this report", None, first_bold=True, fills=fills_for(fr, 2))
h3("E12.5  Cross-LGA benchmarking")
e = T["E12.5"]; rows, fl = [], {}
for i, (_, r) in enumerate(e.iterrows()):
    rows.append([r.Metric] + [P(r[l], 1) for l in lgas])
    if r.Metric == "Coverage" and STALE:
        continue
    vals = r[lgas].astype(float); hi_good = r.Metric in ("Coverage", "Retention")
    if vals.notna().sum() > 1:
        fl[(i, lgas.index(vals.idxmax() if hi_good else vals.idxmin()) + 1)] = "E2EFDA"; fl[(i, lgas.index(vals.idxmin() if hi_good else vals.idxmax()) + 1)] = "F8CBAD"
lw = max(1.6, min(2.4, 13.3 / max(len(lgas), 1)))
table(["Metric"] + lgas, rows, [3.2] + [lw] * len(lgas), "Cross-LGA benchmarking (green = best, red = worst" + ("; coverage not ranked because of the stale denominator)" if STALE else ")"), "this report", "LGA", first_bold=True, fills=fl)
para(f"Of {len(X01)} headline metrics, {n_better} improved, {n_worse} worsened and {n_same} stayed the same. {R['prio_recovered']} of {R['prio_prev']} Cycle {PREV} priority facilities recovered to 90% or more. {bench} is the internal benchmark LGA (best combined rank on out-of-range age, duplicates, repeat dosing and retention).")

# ------------------------------------------------------------------ 13-15
h1("13. Positive findings and good practice")
h3("E13.1  Positive findings list")
bullets(POS)
h1("14. What this means for each stakeholder")
h3("E14.1  Implications and asks")


def asks(*owners):
    rs = recs_of(*owners)
    return "; ".join(f"{r['action'].rstrip('.')} ({r['no']})" for r in rs[:3]) if rs else "No new ask this cycle."


def says(*owners):
    ev = [a for a in A if a["owner"] in owners][:2]
    return "; ".join(f"{a['name']}: {a['ev']}" for a in ev) if ev else "No anomaly owned this cycle."


high = [r for r in RECS if r["prio"] == "High"][:2]
table(["Stakeholder", "What the data says", "Decisions and asks"], [
    (NAT, f"{N(c4['R03'])} children dosed; " + ("coverage cannot be read at face value; " if STALE else "") + (f"{N(c4['oor'])} doses above {AGE_MAX} months." if c4["oor"] else "no age-ceiling breaches."),
     "Confirm national rules behind: " + "; ".join(f"{r['action'].rstrip('.')} ({r['no']})" for r in high) if high else "No new ask."),
    (f"{STATE} State SMC programme", says("state", "lga"), asks("state", "lga")),
    ("Funders", f"Activation {P(c4['F01'])}, {P(c4['T02'])} of doses in the window; {len([a for a in A if a['sev'].startswith('High')])} high-severity anomalies.",
     "Fund the top-priority fixes before the next cycle: " + ", ".join(r["no"] for r in high) + "." if high else "No new ask."),
    ("Implementing partners", says("partners"), asks("partners")),
    ("Supply-chain partners", says("supply"), asks("supply")),
    ("DIGIT HCM / platform team", says("platform"), asks("platform", "data"))],
    [3.5, 6.3, 6.7], "Implications and asks", "this report", None, first_bold=True)
h1("15. Recommendations and next-cycle plan")
h3("E15.1  Recommendations")
rr = [(r["no"], r["action"], r["ev"], OWNER[r["owner"]], r["timing"], r["prio"]) for r in RECS]
table(["#", "Action", "Evidence", "Owner", "Timing", "Priority"], rr, [0.9, 5, 4.4, 2.8, 2, 1.5], "Recommendations", "this report", None, fills=fills_for(rr, 5))
h3("E15.2  Baselines and proposed targets")
table(["Indicator", f"Cycle {PREV}", f"Cycle {CUR}", "Proposed next-cycle target"], [
    ("Coverage (unique children)", P(c3["R07"]), P(c4["R07"]), "Re-baseline after denominator refresh; 90%+" if STALE else "90%+"),
    ("Facilities below threshold", N(R["R12_prev"]), N(R["R12"]), "Under 5"), ("Repeat-dosed children", N(c3["S01"]), N(c4["S01"]), "Zero"),
    ("Supply ratio", P(c3["K12"], 0), P(c4["K12"], 0), "100% or more at every facility"),
    ("Registered with no outcome", P(c3["R10"] / c3["R05"]), P(c4["R10"] / c4["R05"]), "Under 1%"),
    (f"Dose records above {AGE_MAX} months", N(c3["oor"]), N(c4["oor"]), "Zero"),
    ("State-to-HF stock rejected", P(c3["K02_rej_pct"], 2), P(c4["K02_rej_pct"], 2), "Under 1%")],
    [6, 2.5, 2.5, 5.5], "Baselines and proposed targets (proposals for the national programme to confirm)", "this report", None, first_bold=True)

# ------------------------------------------------------------------ annexes
doc.add_page_break(); h1("Annexes")
h3(f"Annex A  Priority facilities for Cycle {CUR + 1}")
ct = pd.read_csv(ROOT / "01_input_deidentified" / f"IN-1_children_treated_C{CUR}.csv", dtype=str, usecols=["child_key", "Health Facility", "Administration Status"])
s = ct[ct["Administration Status"] == "ADMINISTRATION_SUCCESS"]
rep_share = (s.groupby("Health Facility").child_key.value_counts() - 1).clip(lower=0).groupby(level=0).sum() / s.groupby("Health Facility").size()
prev_cov = hf3.set_index("Health Facility").coverage
table(["Facility", "LGA", "Coverage", f"Cycle {PREV} status", "Supply ratio", "Users", "Repeat share", "Driver"],
      [(short(r["Health Facility"]), r.LGA, P(r.coverage), "Persistent" if prev_cov.get(r["Health Facility"], 1) < .9 else "New", P(r.supply_ratio, 0), N(r.users_active),
        P(rep_share.get(r["Health Facility"], 0)), r.driver) for _, r in below.sort_values("coverage").iterrows()],
      [4.4, 2.2, 1.6, 1.8, 1.6, 1.2, 1.6, 2.6], f"Priority facilities (below 90% in Cycle {CUR})", f"{SRC_HF}; {SRC_ST}", f"{R['R12']} facilities")
h3("Annex B  Ward detail")
ward = hf4.groupby(["LGA", "Ward"]).agg(f=("Health Facility", "count"), reg=("registered", "sum"), dosed=("dosed", "sum")).reset_index()
ward["cv"] = ward.dosed / ward.reg
table(["LGA", "Ward", "Facilities", "Registered", "Dosed", "Coverage"], [(r.LGA, r.Ward, N(r.f), N(r.reg), N(r.dosed), P(r.cv)) for _, r in ward.sort_values(["LGA", "cv"]).iterrows()],
      [3.4, 4.6, 2, 2.4, 2.4, 2], "Ward detail", SRC_HF, f"{c4['ward_n']} wards")
h3("Annex C  Companion data files")
table(["File", "Content"], [("01_input_deidentified/", f"De-identified inputs for Cycles {PREV} and {CUR}, README and de-identification audit"),
                            (f"02_analysis/{STATE}_SMC_C{CUR}_analysis.xlsx", "All exhibit tables, cycle metrics and the facility table"),
                            ("02_analysis/fig/", "Figures as PNG"), ("smc-dne-reporter skill scripts", "deidentify.py, verify.py, analysis.py, build_report.py")], [6, 10.5], "Companion data files")
h3("Annex D  Metric dictionary extract")
para("Metric IDs as defined in SMC_Report_Standard_Spec_v1. Local changes: F08/F09 duplicate rate uses the export's Record Type = duplicate flag; "
     f"G01 Cycle {PREV} split uses ID creation date only; V39 counts facilities whose Registered Count is unchanged.")
h3("Annex E  Method notes, limitations and de-identification")
table(["Field in source", "What was done", "Effect on this report"], [
    ("Child name, household head name", "Removed. Quality flags and keyed duplicate / household keys computed first", "V01-V07, V24 and H06 run normally; name match between cycles not assessable"),
    ("Latitude / longitude", "Removed. GPS status, precision, distance from facility median and keyed point ID computed first; facility median points kept (3 decimals)", "V13-V17, F07 run normally; GPS recapture between cycles not assessable; no household map"),
    ("Individual ID, Client Reference ID", "Replaced with keyed hashes (same key in both cycles); ID creation date kept", "Reach, repeat dosing, retention and reuse identical to raw"),
    ("Username (CDD / OIC)", "Replaced with account codes", "Field-force metrics identical; codes map back only with the private key file"),
    ("Person name in user reports", "Removed", "None"),
    ("Facility, ward, LGA, age, sex, dates, product, status", "Kept", "None")], [3.6, 6.6, 6.3], "De-identification applied to the inputs")
h3("Annex F  Glossary")
table(["Term", "Meaning"], [("LGA", "Local Government Area (admin level 2)"), ("Ward", "Admin level 3"), ("HF", "Health facility"), ("CDD", "Community drug distributor (app user account)"),
                            ("OIC", "Officer in charge (facility warehouse manager account)"), ("SPAQ 1 / SPAQ 2", f"SP + amodiaquine for 3-11 / 12-{AGE_MAX} months"),
                            ("Redose visit", "VISITED status: child re-dosed after vomiting"), ("[ASSOCIATION]", "Pattern in aggregate data; cause not proven")], [4, 12.5], "Glossary", first_bold=True)
h3("Annex G  Sign-off")
table(["Role", "Name", "Date"], [("Prepared by", "", ""), ("Reviewed by", "", ""), ("Approved by", "", ""), ("Version / data cut-off", f"{cfg.get('version', '1.0')} / {D(c4['last_date'])}", "")], [5, 7, 4.5], "Sign-off", first_bold=True)

for sec in doc.sections:
    for hp in list(sec.header.paragraphs) + list(sec.footer.paragraphs):
        if "{{" in hp.text:
            new = hp.text.replace("SMC Cycle Report Template v1.0", "SMC Cycle Report v1.0").replace("{{ADMIN1_NAME}}", f"{STATE} State").replace("{{CYCLE_NO}}", str(CUR))
            for i, r_ in enumerate(hp.runs):
                r_.text = new if i == 0 else ""
out = OUTDIR / f"{STATE}_SMC_Cycle{CUR}_Report_v{cfg.get('version', '1.0')}.docx"
try:
    doc.save(out)
except PermissionError:
    out = out.with_name(out.stem + "_new.docx"); doc.save(out)
print("saved", out, "| tables", TN[0], "| figures", FN[0], "| anomalies", len(A), "| recommendations", len(RECS))
