"""One-command pipeline: de-identify -> parity check -> analysis -> report -> report scan -> PDF.

Usage:  python -X utf8 run_all.py <config.json> [--from analysis|report] [--no-pdf]
Stops at the first failing gate. Re-run with --from to skip finished stages.
"""
import subprocess, sys, json
from pathlib import Path

HERE = Path(__file__).resolve().parent
cfg_path = sys.argv[1]; args = sys.argv[2:]
start = args[args.index("--from") + 1] if "--from" in args else "deidentify"
STAGES = [("deidentify", ["deidentify.py"]), ("parity", ["verify.py", "parity"]), ("analysis", ["analysis.py"]),
          ("report", ["build_report.py"]), ("scan", ["verify.py", "report"])]
names = [s for s, _ in STAGES]
for name, cmd in STAGES[names.index(start):]:
    print(f"\n===== {name} =====", flush=True)
    r = subprocess.run([sys.executable, "-X", "utf8", "-W", "ignore", str(HERE / cmd[0]), cfg_path] + cmd[1:])
    if r.returncode != 0:
        sys.exit(f"STOPPED at stage '{name}' (exit {r.returncode}). Fix and re-run with --from {name}.")
if "--no-pdf" not in args and sys.platform == "win32":
    out = Path(json.load(open(cfg_path, encoding="utf-8"))["out_root"]) / "03_report"
    docx = sorted(out.glob("*.docx"), key=lambda p: p.stat().st_mtime)[-1]
    ps = (f"$w=New-Object -ComObject Word.Application;$w.Visible=$false;try{{$d=$w.Documents.Open('{docx}',$false,$true);"
          f"$d.ExportAsFixedFormat('{docx.with_suffix('.pdf')}',17);'pages: '+$d.ComputeStatistics(2);$d.Close($false)}}finally{{$w.Quit()}}")
    subprocess.run(["powershell", "-NoProfile", "-Command", ps])
print("\nALL STAGES PASSED")
