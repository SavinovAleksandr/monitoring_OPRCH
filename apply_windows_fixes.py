#!/usr/bin/env python3
"""Apply narrowly-scoped VBA compatibility fixes to a local monitoring_OPRCH clone.

Does not need extra Python packages, GitHub credentials, or Excel.
Makes backup copies of files changed in the current execution.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import re
import shutil
import sys

MODULES = ("Common", "Engine", "Station", "Sheets", "Charts", "Monitor")

# Exact replacements intentionally fail on an unknown revision; don't silently alter unrelated code.
REPLACEMENTS = {
    "OPRCH_Sheets.bas": [
        (
            "        res.PReq, res.PFact, _\n        Round(res.AmplPctPnom, 2), res.AmplitudeTag, _",
            "        res.PReq, res.PFact, Round(res.AmplPctPnom, 2), res.AmplitudeTag, _",
        ),
        (
            "        res.TransientType, res.NumExtrema, _\n        qualStatusStr, qualT5Str, qualT10Str, qualStStr, _",
            "        res.TransientType, res.NumExtrema, qualStatusStr, qualT5Str, qualT10Str, qualStStr, _",
        ),
        (
            "        res.PMaxEff, res.PMinEff, res.ReservePlus, res.ReserveMinus, _\n        res.PReqOrig, limitStr, _",
            "        res.PMaxEff, res.PMinEff, res.ReservePlus, res.ReserveMinus, res.PReqOrig, limitStr, _",
        ),
        (
            "Public Sub CollectOldOutputSheets(ByRef names As Collection)\n    Dim ws As Worksheet\n    For Each ws In ThisWorkbook.Worksheets\n        ' Универсальная книга: сохраняем только рабочие листы ввода/журнала.\n        ' Лист «Запрос» и прочие служебные вкладки при очистке/пересчёте удаляются.\n        If ws.Name <> SH_RAW And ws.Name <> SH_CFG And ws.Name <> SH_SUM And ws.Name <> SH_LOG _\n           And ws.Name <> SH_INSTR Then\n            names.Add ws.Name\n        End If\n    Next ws\nEnd Sub",
            "Public Sub CollectOldOutputSheets(ByRef names As Collection)\n    Dim wsSummary As Worksheet, ws As Worksheet\n    Dim r As Long, c As Long, nm As String\n    ' Удаляем только листы, перечисленные в старом Summary.\n    ' Нельзя удалять посторонние листы пользователя (в том числе «Запрос»).\n    On Error Resume Next\n    Set wsSummary = ThisWorkbook.Worksheets(SH_SUM)\n    On Error GoTo 0\n    If Not wsSummary Is Nothing Then\n        For r = 2 To LastUsedRow(wsSummary)\n            For c = 27 To 28  ' AA/AB: лист данных и графика\n                If Not IsError(wsSummary.Cells(r, c).Value) Then\n                    nm = Trim$(CStr(wsSummary.Cells(r, c).Value))\n                    If Len(nm) > 0 And nm <> SH_RAW And nm <> SH_CFG _\n                       And nm <> SH_SUM And nm <> SH_LOG And nm <> SH_INSTR Then\n                        Set ws = Nothing\n                        On Error Resume Next\n                        Set ws = ThisWorkbook.Worksheets(nm)\n                        If Not ws Is Nothing Then names.Add ws.Name, ws.Name\n                        On Error GoTo 0\n                    End If\n                End If\n            Next c\n        Next r\n    End If\n    Set ws = Nothing\n    On Error Resume Next\n    Set ws = ThisWorkbook.Worksheets(SH_PROT)\n    If Not ws Is Nothing Then names.Add ws.Name, ws.Name\n    On Error GoTo 0\nEnd Sub",
        ),
        (
            '    Application.DisplayAlerts = False\n    For i = names.Count To 1 Step -1\n        On Error Resume Next\n        ThisWorkbook.Worksheets(CStr(names(i))).Delete\n        On Error GoTo 0\n    Next i\n    Application.DisplayAlerts = True',
            '    Dim alertsBefore As Boolean, failureText As String\n    alertsBefore = Application.DisplayAlerts\n    Application.DisplayAlerts = False\n    On Error GoTo CleanFail\n    For i = names.Count To 1 Step -1\n        ThisWorkbook.Worksheets(CStr(names(i))).Delete\n    Next i\n    Application.DisplayAlerts = alertsBefore\n    Exit Sub\nCleanFail:\n    failureText = Err.Description\n    Application.DisplayAlerts = alertsBefore\n    Err.Raise vbObjectError + 2130, "DeleteOutputSheets", _\n        "Не удалось удалить старый выходной лист: " & failureText',
        ),
    ],
    "OPRCH_Monitor.bas": [
        (
            '    Dim fCol As Long\n\n    t0Run = Timer',
            '    Dim fCol As Long\n    Dim previousCalculation As XlCalculation, previousUpdating As Boolean\n    Dim stateSaved As Boolean, failureNumber As Long, failureText As String\n\n    previousCalculation = Application.Calculation\n    previousUpdating = Application.ScreenUpdating\n    stateSaved = True\n    t0Run = Timer',
        ),
        (
            '    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n\n    OPRCHMsg',
            '    Application.Calculation = previousCalculation\n    Application.ScreenUpdating = previousUpdating\n\n    OPRCHMsg',
        ),
        (
            'EH:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    OPRCHMsg "Ошибка AnalyzeOPRCH (" & stepName & "): " & Err.Description, vbCritical\nEnd Sub',
            'EH:\n    failureNumber = Err.Number\n    failureText = Err.Description\n    On Error Resume Next\n    If stateSaved Then\n        Application.Calculation = previousCalculation\n        Application.ScreenUpdating = previousUpdating\n    End If\n    AppendLog "ERROR", stepName, "Ошибка #" & CStr(failureNumber) & ": " & failureText\n    On Error GoTo 0\n    OPRCHMsg "Ошибка AnalyzeOPRCH (" & stepName & ", #" & failureNumber & "): " & _\n             failureText & vbCrLf & "Подробности: лист Log.", vbCritical\nEnd Sub',
        ),
        (
            'Public Sub InitOPRCHUISilent()\n    On Error Resume Next\n    gOPRCHQuiet = True\n    WriteInstructionSheet\n    EnsureControlButtons EnsureSheet(SH_CFG)\n    EnsureControlButtons EnsureSheet(SH_INSTR)\nEnd Sub',
            'Public Sub InitOPRCHUISilent()\n    Dim previousQuiet As Boolean, errorInfo As String\n    previousQuiet = gOPRCHQuiet\n    gOPRCHQuiet = True\n    On Error GoTo InitFail\n    WriteInstructionSheet\n    EnsureControlButtons EnsureSheet(SH_CFG)\n    EnsureControlButtons EnsureSheet(SH_INSTR)\n    gOPRCHQuiet = previousQuiet\n    Exit Sub\nInitFail:\n    errorInfo = "Ошибка инициализации UI #" & CStr(Err.Number) & ": " & Err.Description\n    On Error Resume Next\n    AppendLog "ERROR", "InitOPRCHUISilent", errorInfo\n    gOPRCHQuiet = previousQuiet\nEnd Sub',
        ),
        (
            '    Application.DisplayAlerts = False\n    DeleteOutputSheets targetSheets\n    Application.DisplayAlerts = True\n    MsgBox',
            '    DeleteOutputSheets targetSheets\n    MsgBox',
        ),
    ],
    "build_oprch_1_8.py": [
        (
            'import subprocess\nimport zipfile',
            'import subprocess\nimport tempfile\nimport zipfile',
        ),
        (
            'ROOT = Path("/Users/asavinov/ОПРЧ/ОПРЧ")',
            'ROOT = Path(__file__).resolve().parent',
        ),
        (
            'shell = Path("/tmp/oprch_vba_shell.xlsm")',
            'shell = Path(tempfile.gettempdir()) / "oprch_vba_shell.xlsm"',
        ),
    ],
    "OPRCH_Install.bas": [
        (
            'Установщик OPRCH v1.8.2',
            'Установщик OPRCH v1.8.6 (Windows fixes)',
        ),
        (
            'Макросы OPRCH v1.8.2 установлены.',
            'Макросы OPRCH v1.8.6 установлены.',
        ),
    ],
}


def apply_replacements(filename: str, source: str) -> tuple[str, int]:
    count = 0
    for old, new in REPLACEMENTS.get(filename, ()):
        if old in source:
            source = source.replace(old, new, 1)
            count += 1
        elif new in source:
            # Already patched; do not duplicate a change on a second run.
            continue
        else:
            raise RuntimeError(f"Unknown file revision in {filename}; missing patch anchor:\n{old[:180]}")
    return source, count


def check_vba_source(path: Path, source: str) -> None:
    runs = []
    pending = 0
    for line_number, line in enumerate(source.splitlines(), 1):
        if re.search(r"\s_\s*$", line):
            pending += 1
        else:
            if pending > 24:
                runs.append((line_number - pending, pending))
            pending = 0
    if runs:
        raise RuntimeError(f"VBA has >24 line continuations in {path.name}: {runs}")

    too_long = [(i, len(line)) for i, line in enumerate(source.splitlines(), 1) if len(line) > 1023]
    if too_long:
        raise RuntimeError(f"VBA has >1023-char lines in {path.name}: {too_long}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path.cwd(), help="Path to local monitoring_OPRCH checkout")
    parser.add_argument("--check", action="store_true", help="Check patched source only (do not write)")
    args = parser.parse_args()
    root: Path = args.repo.resolve()
    if not (root / "OPRCH_Monitor.bas").is_file():
        parser.error(f"Not a monitoring_OPRCH root: {root}")

    sources = {}
    replacements = 0
    for module in MODULES:
        name = f"OPRCH_{module}.bas"
        source = (root / name).read_text("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
        if not args.check:
            source, modified = apply_replacements(name, source)
            replacements += modified
        check_vba_source(root / name, source)
        sources[name] = source

    extra = {}
    for name in ("OPRCH_Install.bas", "build_oprch_1_8.py"):
        source = (root / name).read_text("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
        if not args.check:
            source, modified = apply_replacements(name, source)
            replacements += modified
        extra[name] = source

    assert "gOPRCHQuiet = previousQuiet" in sources["OPRCH_Monitor.bas"]
    if args.check:
        print("PASS: 6 VBA modules, no continuation/physical-line violations; silent state is restored")
        return 0

    # Generate real cp1251, not UTF-8 text under a *_win1251.bas name.
    generated = {}
    for module in MODULES:
        name = f"OPRCH_{module}.bas"
        slim = "\n".join(line for line in sources[name].splitlines() if not line.lstrip().startswith("'"))
        data = (slim.rstrip() + "\n").replace("\n", "\r\n").encode("cp1251", errors="strict")
        if len(data) >= 64000:
            raise RuntimeError(f"{module}: Windows VBA module is too large ({len(data)} bytes)")
        generated[f"OPRCH_{module}_win1251.bas"] = data
    generated["OPRCH_Install_win1251.bas"] = extra["OPRCH_Install.bas"].replace("\n", "\r\n").encode("cp1251", errors="strict")

    writes: dict[str, bytes] = {
        **{n: t.encode("utf-8") for n, t in sources.items()},
        **{n: t.encode("utf-8") for n, t in extra.items()},
        **generated,
    }
    to_change = {n: data for n, data in writes.items() if not (root / n).exists() or (root / n).read_bytes() != data}
    # Preserve a clean repo when already patched: only make backups for edits.
    if to_change:
        backup_dir = root / (".oprch_windows_backup_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
        backup_dir.mkdir()
        for name in to_change:
            if (root / name).is_file():
                shutil.copy2(root / name, backup_dir / name)
        for name, content in to_change.items():
            (root / name).write_bytes(content)
        print(f"Backups: {backup_dir}")
    print(f"Patched {replacements} source fragments; wrote {len(to_change)} files")
    for name, data in generated.items():
        print(f"  {name}: {len(data)} bytes; Windows-1251")
    print("Next: run make_windows_workbook.ps1 in desktop Excel for Windows")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (RuntimeError, UnicodeError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
