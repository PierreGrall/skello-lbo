"""Rebuilds the whole project from data/inputs.json. Stops at the first failed guard."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
STEPS = [
    ["src/analysis.py"],
    ["src/charts.py"],
    ["src/build_model.py"],
    ["src/recalc.py", "model/Skello_LBO_Model.xlsx"],
    ["src/reconcile.py"],
    ["src/build_memo.py"],
]
for step in STEPS:
    print(">", " ".join(step))
    subprocess.run([sys.executable, *step], cwd=ROOT, check=True, stdout=None if step[0] != "src/analysis.py" else subprocess.DEVNULL)
subprocess.run([sys.executable, "-m", "pytest", "tests", "-q"], cwd=ROOT, check=True)
print("Build complete.")
