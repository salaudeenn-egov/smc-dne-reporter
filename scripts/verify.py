"""Verification gates.

Usage:
  python -X utf8 verify.py <config.json> parity   # raw vs de-identified: core metrics must be identical
  python -X utf8 verify.py <config.json> report   # finished .docx: no names, IDs, usernames or {{placeholders}}

Exit code 0 = pass, 2 = fail. Never print matched personal values; print counts only.
"""
import json, re, sys
from pathlib import Path
import pandas as pd

cfg = json.load(open(sys.argv[1], encoding="utf-8")); mode = sys.argv[2]
ROOT = Path(cfg["out_root"]); IN = ROOT / "01_input_deidentified"
CT_PAT = cfg.get("patterns", {}).get("ct", r"(children_treated|child_treatment)")
COLS = ["Individual ID", "Username", "Administration Status", "latitude", "Child Name", "Household Head Name", "LGA", "Ward", "Health Facility"]


def raw_ct(c):
    spec = cfg["cycles"][str(c)]
    if isinstance(spec, dict) and "ct" in spec.get("files", {}):
        v = spec["files"]["ct"]; paths = [Path(x) for x in (v if isinstance(v, list) else [v])]
    else:
        folder = Path(spec if isinstance(spec, str) else spec["dir"])
        paths = sorted(p for p in folder.rglob("*.xls*") if re.search(CT_PAT, p.name, re.I) and "chunk" not in str(p).lower() and not p.name.startswith("~$"))
    fr = []
    for p in paths:
        for _, df in pd.read_excel(p, sheet_name=None, engine="calamine", dtype=str).items():
            if set(COLS) <= set(df.columns):
                fr.append(df[COLS])
    return pd.concat(fr, ignore_index=True)


def metrics(df, idc, uc):
    s = df[df["Administration Status"] == "ADMINISTRATION_SUCCESS"]; vc = s[idc].value_counts(); per = s.groupby(uc).size()
    return {"rows": len(df), "children_dosed": s[idc].nunique(), "repeat_children": int((vc > 1).sum()), "extra_records": int((vc - 1).clip(lower=0).sum()),
            "workers": df[uc].nunique(), "max_doses_one_worker": int(per.max()), "median_doses_per_worker": float(per.median())}


if mode == "parity":
    cyc = [c for c in (cfg.get("prev_cycle"), cfg["cycle"]) if c]
    raw = {c: raw_ct(c) for c in cyc}
    de = {c: pd.read_csv(IN / f"IN-1_children_treated_C{c}.csv", dtype=str) for c in cyc}
    rows = []
    for c in cyc:
        a, b = metrics(raw[c], "Individual ID", "Username"), metrics(de[c], "child_key", "worker_code")
        rows += [(f"C{c}", k, a[k], b[k]) for k in a]
        rows.append((f"C{c}", "gps_missing", int(raw[c]["latitude"].isna().sum()), int((de[c]["gps_status"] == "blank").sum())))
    if len(cyc) == 2:
        ok = lambda d, i: set(d[d["Administration Status"] == "ADMINISTRATION_SUCCESS"][i])
        rows.append(("both", "children dosed in both cycles", len(ok(raw[cyc[0]], "Individual ID") & ok(raw[cyc[1]], "Individual ID")),
                     len(ok(de[cyc[0]], "child_key") & ok(de[cyc[1]], "child_key"))))
    t = pd.DataFrame(rows, columns=["cycle", "metric", "raw", "de-identified"]); t["check"] = (t.raw == t["de-identified"]).map({True: "SAME", False: "DIFF"})
    print(t.to_string(index=False))
    (ROOT / "02_analysis").mkdir(exist_ok=True); t.to_csv(ROOT / "02_analysis" / "parity_check.csv", index=False)
    sys.exit(0 if (t.check == "SAME").all() else 2)

if mode == "report":
    import docx
    rep = sorted((ROOT / "03_report").glob("*.docx"), key=lambda p: p.stat().st_mtime)[-1]
    d = docx.Document(rep)
    parts = [p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells]
    for s in d.sections:
        parts += [p.text for p in s.header.paragraphs + s.footer.paragraphs]
    txt = "\n".join(parts); low = txt.lower()
    raw = pd.concat([raw_ct(c) for c in [x for x in (cfg.get("prev_cycle"), cfg["cycle"]) if x]])
    names = {n.strip() for n in pd.concat([raw["Child Name"], raw["Household Head Name"]]).dropna() if len(n.strip().split()) >= 2}
    admin = [a.lower() for a in pd.concat([raw.LGA, raw.Ward, raw["Health Facility"]]).dropna().unique()]
    names = {n for n in names if n.lower() in low}                          # fast pre-filter: only names present in the text
    names = {n for n in names if not any(n.lower() in a for a in admin)}   # a "name" that is also a place name is not personal data
    users = set(raw["Username"].dropna())
    words = re.findall(r"[a-z']+", low); wc = {}
    for w in words:
        wc[w] = wc.get(w, 0) + 1

    # vocabulary of the report's own prose: build_report.py strings + template + spec. A "name" made only of
    # these words (junk entries such as "may have", "at least") is report wording, not personal data.
    skill = Path(__file__).resolve().parent.parent
    prose = (skill / "scripts" / "build_report.py").read_text(encoding="utf-8").lower()
    prose += "\n".join(p.text for p in docx.Document(skill / "assets" / "SMC_Cycle_Report_Template_v1.docx").paragraphs).lower()
    for _, sdf in pd.read_excel(skill / "assets" / "SMC_Report_Standard_Spec_v1.xlsx", sheet_name=None, header=None, dtype=str).items():
        prose += " " + " ".join(sdf.fillna("").values.ravel()).lower()
    vocab = set(re.findall(r"[a-z']+", prose))

    def leaked(n):
        if set(re.findall(r"[a-z']+", n.lower())) <= vocab:
            return False
        n = re.sub(r"\s+", " ", n.lower().strip())
        k = len(re.findall(r"(?<![a-z])" + re.escape(n) + r"(?![a-z])", low))   # whole-word match only
        if not k:
            return False
        # junk "names" made of ordinary words ("what the") also occur elsewhere in the report;
        # a real leaked name has at least one word that appears only inside that name
        return any(wc.get(t, 0) <= k for t in re.findall(r"[a-z']+", n) if len(t) >= 3)

    res = {"full_name_hits": sum(1 for n in names if leaked(n)),
           "username_hits": sum(1 for u in users if u in txt),
           "uuid_hits": len(re.findall(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", txt)),
           "placeholders": len(re.findall(r"\{\{", d.element.xml)) + txt.count("{{"),
           "nan_text": len(re.findall(r"\bnan\b", txt))}
    print(rep.name, res)
    sys.exit(0 if not any(res.values()) else 2)
