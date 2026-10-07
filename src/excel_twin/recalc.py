"""
Recalculate a copy of a workbook with LibreOffice, so its formulas carry values.

LibreOffice does not recalculate .xlsx files on load by default, and openpyxl stores formulas without
results. A one-line Basic macro, installed in a throwaway profile, calls calculateAll() and saves.
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def recalc(source: Path, out_dir: Path, timeout: int = 600) -> Path:
    """Copy ``source`` into ``out_dir``, recalculate it with LibreOffice and return the copy's path."""
    exe = soffice()
    if not exe:
        raise RuntimeError("LibreOffice (soffice) is not on PATH")
    source, out_dir = Path(source), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / source.name
    shutil.copyfile(source, target)
    before = target.stat().st_mtime_ns
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile:
        url = Path(profile).as_uri()
        subprocess.run([exe, "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       check=True, capture_output=True, timeout=120)
        macro_dir = Path(profile) / "user" / "basic" / "Standard"
        (macro_dir / "Module1.xba").write_text(MACRO)
        subprocess.run([exe, "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                        str(target.resolve())], check=True, capture_output=True, timeout=timeout)
    if target.stat().st_mtime_ns == before:
        raise RuntimeError("LibreOffice exited without saving: nothing was recalculated")
    return target
