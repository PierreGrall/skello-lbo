"""Recalculates a workbook in place with headless LibreOffice, so cached values exist."""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def recalc(path: Path) -> None:
    path = Path(path)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            ["soffice", "--headless", f"-env:UserInstallation=file://{tmp}/profile", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
             "--outdir", tmp, str(path)],
            check=True, capture_output=True, timeout=180,
        )
        shutil.copy(Path(tmp) / path.name, path)


if __name__ == "__main__":
    recalc(Path(sys.argv[1]))
