#!/usr/bin/env python3
"""Split OPRCH_Monitor_all.bas → 6 modules (лимит VBA ~64 КБ), rebuild xlsx, build xlsm."""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
MASTER = ROOT / "OPRCH_Monitor_all.bas"

MONITOR_PROCS = {
    "SetupOPRCHTemplate", "ExportOPRCHReport", "AnalyzeOPRCH",
    "AnalyzeOPRCHSilent",
    "ClearOPRCHResults", "ApplyPresetsToConfig",
    "EnsureControlButtons", "AddOrReplaceButton", "WriteInstructionSheet",
    "InitOPRCHUI", "InitOPRCHUISilent",
}
ENGINE_PROCS = {
    "AnalyzeOneGenerator", "EvaluateReaction15", "React15ThresholdMW",
    "ClassifyReaction15", "EvaluateDeficit", "BuildGeneratorVerdict",
    "EvaluateApplicability", "SelectEvaluationPoint", "EvaluateOscillations",
    "EvaluateTelemetry", "ClassifyReactionType",
    "EvaluateQualitative", "ComputeSteadyMean", "ComputeSteadyPReqMean",
    "EvaluateTransient", "FindEvents", "AvgPowerBackward", "ResolveF0",
}
# Станционные агрегаты — отдельный модуль (Sheets был ~63 КБ и не импортировался)
STATION_PROCS = {
    "BuildStationAggregates", "BuildParoprovodSubAggregates", "SumPowerAtRow",
    "BuildOneStationAggregate", "BuildOneStationAggregateEvent",
    "AnnotateGeneratorRowsWithStationVerdict", "BuildStationVerdict",
    "StationSheetName", "StationChartSheetName", "WriteStationChartSheet",
}
CHARTS_PROCS = {
    "WriteGeneratorChartSheet", "ApplyChartAxesFormat",
    "GetRangeMinMaxByCols", "MaxAbsInCols", "WriteChartVerdictBlock", "FindChartEndRow",
    "AddSeries", "AddMarkerSeries", "AddLimitLineSeries",
    "FillMarkerColumn", "FillMarkerColumnAtTime", "ApplyGeneratorSheetFormats",
    "LoadChartDisplayOptions", "ChartOpt", "EnsureChartOptionCheckboxes",
}
SHEETS_PROCS = {
    "WriteGeneratorSheet",
    "WriteSummaryRow", "WriteSummaryInvalid", "WriteSummaryStationRow",
    "ApplySummaryConditionalFormat", "ApplyStatusCF", "WriteVersionStamp",
    "InitLog", "AppendLog", "FinalizeLog", "ValidateInputs",
    "WriteProtocolSheet", "AddProtocolLine", "AddProtocolGenerator", "AddProtocolStation",
    "CollectOldOutputSheets", "DeleteOutputSheets",
    "GeneratorSheetName", "GeneratorChartSheetName",
}

PROC_RE = re.compile(r"^(Public |Private )(Sub |Function )([A-Za-z0-9_]+)")

MODULE_ORDER = ["Common", "Engine", "Station", "Sheets", "Charts", "Monitor"]


def split_master(text: str) -> dict[str, str]:
    lines = text.splitlines()
    first_proc = next(i for i, l in enumerate(lines) if PROC_RE.match(l))
    preamble = lines[:first_proc]
    pre = []
    for l in preamble:
        if l.startswith("Private Const "):
            l = "Public" + l[7:]
        elif l.startswith("Private Type "):
            l = "Public" + l[7:]
        elif l.startswith("Private m_"):
            l = "Public" + l[7:]
        if l.startswith("Attribute VB_Name"):
            continue
        pre.append(l)
    while pre and pre[0] == "":
        pre.pop(0)
    if pre and pre[0].strip() == "Option Explicit":
        pre = pre[1:]
        if pre and pre[0] == "":
            pre = pre[1:]

    idxs = [i for i, l in enumerate(lines) if PROC_RE.match(l)]
    idxs.append(len(lines))
    buckets = {k: [] for k in MODULE_ORDER}
    for a, b in zip(idxs, idxs[1:]):
        name = PROC_RE.match(lines[a]).group(3)
        block = lines[a:b]
        while block and block[-1] == "":
            block.pop()
        block.append("")
        block[0] = re.sub(r"^Private (Sub|Function) ", r"Public \1 ", block[0])
        if name in MONITOR_PROCS:
            buckets["Monitor"].append(block)
        elif name in ENGINE_PROCS:
            buckets["Engine"].append(block)
        elif name in STATION_PROCS:
            buckets["Station"].append(block)
        elif name in CHARTS_PROCS:
            buckets["Charts"].append(block)
        elif name in SHEETS_PROCS:
            buckets["Sheets"].append(block)
        else:
            buckets["Common"].append(block)

    headers = {
        "Common": "OPRCH_Common",
        "Engine": "OPRCH_Engine",
        "Station": "OPRCH_Station",
        "Sheets": "OPRCH_Sheets",
        "Charts": "OPRCH_Charts",
        "Monitor": "OPRCH_Monitor",
    }
    notes = {
        "Common": "Типы, константы, общие функции. Импорт 1.",
        "Engine": "Расчёт генератора. Импорт 2.",
        "Station": "Суммы станций (п. 4.2.10). Импорт 3.",
        "Sheets": "Summary/Log/Протокол/листы генераторов. Импорт 4.",
        "Charts": "Графики. Импорт 5.",
        "Monitor": "Точки входа AnalyzeOPRCH. Импорт 6.",
    }
    out = {}
    for key in MODULE_ORDER:
        attr = headers[key]
        body = [
            f'Attribute VB_Name = "{attr}"',
            "Option Explicit",
            "",
            f"' {notes[key]}",
            "' OPRCH_Monitor v1.8.6 — раздельный модуль (лимит VBA ~64 КБ).",
            "",
        ]
        if key == "Common":
            body.extend(pre)
            if body[-1] != "":
                body.append("")
        for block in buckets[key]:
            body.extend(block)
        text_out = "\n".join(body).rstrip() + "\n"
        out[key] = text_out
        slim = "\n".join(
            ln for ln in text_out.splitlines()
            if not (ln.lstrip().startswith("'") and not ln.lstrip().startswith("Attribute"))
        ) + "\n"
        nbytes = len(slim.replace("\n", "\r\n").encode("cp1251"))
        print(f"{attr:16} lines={text_out.count(chr(10)):4}  win1251_slim={nbytes:6}")
        if nbytes > 58000:
            print("  WARN: close to 64K — риск отказа импорта")
        if nbytes > 64000:
            raise SystemExit(f"{attr} slim > 64K ({nbytes}) — нужно дробить дальше")
    return out


def write_bas(modules: dict[str, str]) -> list[Path]:
    mapping = {
        "Common": "OPRCH_Common.bas",
        "Engine": "OPRCH_Engine.bas",
        "Station": "OPRCH_Station.bas",
        "Sheets": "OPRCH_Sheets.bas",
        "Charts": "OPRCH_Charts.bas",
        "Monitor": "OPRCH_Monitor.bas",
    }
    # удалить устаревшие одиночные большие файлы при пересборке
    win_paths = []
    for key in MODULE_ORDER:
        name = mapping[key]
        utf = modules[key]
        (ROOT / name).write_text(utf, encoding="utf-8")
        slim_lines = []
        for line in utf.splitlines():
            s = line.lstrip()
            if s.startswith("'") and not s.startswith("Attribute"):
                continue
            slim_lines.append(line)
        slim = "\n".join(slim_lines).rstrip() + "\n"
        cp = slim.replace("\n", "\r\n").encode("cp1251")
        wpath = ROOT / name.replace(".bas", "_win1251.bas")
        wpath.write_bytes(cp)
        win_paths.append(wpath)
        print("wrote", name, "win1251", len(cp))
    return win_paths


def load_raw_from_template() -> list[dict]:
    src = ROOT / "Шаблон_мониторинг_ОПРЧ.xlsx"
    if not src.exists():
        src = ROOT / "Шаблон_мониторинг_ОПРЧ_ВТЭЦ-2_31.10.2026.xlsx"
    if not src.exists():
        raise FileNotFoundError("Нет шаблона с демо-данными RawData")
    wb = load_workbook(src, data_only=True)
    ws = wb["RawData"]
    rows = []
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, 1).value
        if t is None:
            continue
        if not isinstance(t, datetime):
            continue
        rec = {
            "t": t.replace(microsecond=0) if hasattr(t, "replace") else t,
            "f": ws.cell(r, 2).value,
            "ТГ-1": ws.cell(r, 3).value,
            "ТГ-3": ws.cell(r, 4).value,
            "ТГ-4": ws.cell(r, 5).value,
            "ТГ-7": ws.cell(r, 6).value,
        }
        if rec["f"] is None:
            continue
        rows.append(rec)
    wb.close()
    return rows


def rebuild_xlsx(rows: list[dict]) -> Path:
    wb = Workbook()
    wsRaw = wb.active
    wsRaw.title = "RawData"
    wsInstr = wb.create_sheet("Инструкция", 0)
    wsCfg = wb.create_sheet("Config")
    wsSum = wb.create_sheet("Summary")
    wsLog = wb.create_sheet("Log")

    hdr_font = Font(bold=True)
    hdr_fill = PatternFill("solid", fgColor="D6DCE4")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    lines = [
        ("Мониторинг ОПРЧ v1.8.6 — готовый шаблон (любая станция / любое событие)", True, 14),
        ("", False, 11),
        ("Откройте этот файл как .xlsm с уже вшитыми макросами (Шаблон_мониторинг_ОПРЧ.xlsm).", False, 11),
        ("Подставьте свои данные в RawData и Config → кнопка «Запустить мониторинг ОПРЧ».", False, 11),
        ("Вкладки «Запрос» нет. Один запуск = колич. + кач. сразу.", False, 11),
        ("Режим Config W13: Авария (Прил. 14) / Испытание (Прил. 1).", False, 11),
    ]
    for i, (txt, bold, size) in enumerate(lines, 1):
        cell = wsInstr.cell(i, 1, txt)
        cell.font = Font(bold=bold, size=size)
    wsInstr.column_dimensions["A"].width = 120

    headers = ["Время", "f ВТЭЦ-2", "ТГ-1", "ТГ-3", "ТГ-4", "ТГ-7"]
    for i, h in enumerate(headers, 1):
        cell = wsRaw.cell(1, i, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = thin
    for i, rec in enumerate(rows, start=2):
        c = wsRaw.cell(i, 1, rec["t"])
        c.number_format = "DD.MM.YYYY HH:MM:SS"
        c.border = thin
        vals = [rec["f"], rec["ТГ-1"], rec["ТГ-3"], rec["ТГ-4"], rec["ТГ-7"]]
        for j, v in enumerate(vals, start=2):
            cell = wsRaw.cell(i, j, v)
            cell.number_format = "0.00"
            cell.border = thin
    for col, w in zip(range(1, 7), (20, 12, 10, 10, 10, 10)):
        wsRaw.column_dimensions[get_column_letter(col)].width = w

    cfg_headers = [
        "Станция", "Генератор", "Колонка_мощности", "Колонка_частоты", "Тип_оборудования",
        "Pном, МВт", "S, %", "fнч, Гц", "Вкл (1/0)", "Кач_вкл (1/0)",
        "t5, c", "dP5, %Pном", "t10, c", "dP10, %Pном", "Уст_допуск, %Pном",
        "В сумму станции (1/0)", "Контр_уст (1/0)", "Паропровод",
        "Pmax, МВт", "Pmin, МВт", "РМ (1/0)", "Компоновка (блок/неблок)",
    ]
    for i, h in enumerate(cfg_headers, 1):
        cell = wsCfg.cell(1, i, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
    gens = [
        ["Воркутинская ТЭЦ-2", "ТГ-1", "ТГ-1", "f ВТЭЦ-2", "ПТУ_неблок", 35, 6, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 35, 0, 0, "неблок"],
        ["Воркутинская ТЭЦ-2", "ТГ-3", "ТГ-3", "f ВТЭЦ-2", "ПТУ_неблок", 25, 6, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 25, 0, 0, "неблок"],
        ["Воркутинская ТЭЦ-2", "ТГ-4", "ТГ-4", "f ВТЭЦ-2", "ПТУ_неблок", 25, 6, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 25, 0, 0, "неблок"],
        ["Воркутинская ТЭЦ-2", "ТГ-7", "ТГ-7", "f ВТЭЦ-2", "ПТУ_неблок", 47, 6, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 47, 0, 0, "неблок"],
    ]
    for r, row in enumerate(gens, start=2):
        for c, v in enumerate(row, start=1):
            wsCfg.cell(r, c, v if v != "" else None)

    wsCfg["W1"] = "Глобальные настройки"
    wsCfg["W1"].font = hdr_font
    settings = [
        (2, "fном, Гц", 50),
        (3, "Время начала события", datetime(2026, 10, 31, 17, 3, 38)),
        (4, "Автопоиск старта (1/0)", 1),
        (5, "Запасной интервал tр (если нет QS), с", 300),
        (6, "Допуск количеств., %", 10),
        (7, "Порог включения в работу, МВт", 1),
        (8, "Pre-start буфер, с (реком. ≥300)", 15),
        (9, "Интервал графика, с", 240),
        (10, "Окно установив./конец события, с", 30),
        (11, "Порог реакции 15 с, %Pном", 0.5),
        (12, "Усреднение P0/Pтек, с (0=текущие)", 0),
        (13, "Режим (Авария/Испытание)", "Авария"),
        (14, "Макс. окно события, с", 600),
        (15, "Макс. число событий (1=только первое)", 3),
        (16, "Ограничение Pтреб, %Pном (0=нет)", 10),
        (17, "Строгий дефицит 4.2.14.1 (1/0)", 1),
    ]
    for r, name, val in settings:
        wsCfg.cell(r, 23, name)
        wsCfg.cell(r, 24, val)
    wsCfg["X3"].number_format = "DD.MM.YYYY HH:MM:SS"
    wsCfg["X13"].number_format = "@"

    wsCfg["AA1"] = "Профили Kd(t)"
    wsCfg["AA1"].font = hdr_font
    kd_hdr = ["Тип_оборудования", "t0, с", "Kд0", "t1, с", "Kд1", "t2, с", "Kд2", "t3, с", "Kд3"]
    for i, h in enumerate(kd_hdr, start=27):
        wsCfg.cell(2, i, h)
        wsCfg.cell(2, i).font = hdr_font
        wsCfg.cell(2, i).fill = hdr_fill
    profiles = [
        ["ПТУ_блок", 0, 0.5, 15, 0.5, 300, 1, 600, 1],
        ["ПТУ_блок_уголь", 0, 0.5, 15, 0.5, 360, 1, 600, 1],
        ["ПТУ_неблок", 0, 0.5, 15, 0.5, 420, 1, 600, 1],
        ["ГТУ", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1],
        ["ПГУ_утил", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1],
        ["ПГУ_сбросн", 0, 0.5, 15, 0.5, 300, 1, 600, 1],
        ["ГПА", 0, 0.5, 15, 0.5, 60, 1, 600, 1],
    ]
    for i, p in enumerate(profiles, start=3):
        for j, v in enumerate(p, start=1):
            wsCfg.cell(i, 26 + j, v)

    wsCfg["AJ1"] = "Ряды графиков (1/0)"
    wsCfg["AJ1"].font = hdr_font
    wsCfg["AJ2"] = "Ряд"
    wsCfg["AK2"] = "Относит."
    wsCfg["AL2"] = "Абсолют."
    series = ["Pфакт", "Pтреб", "+Допуск", "-Допуск", "Pmax/Pmin", "Частота", "Маркеры t5/t10/fнч", "Резерв"]
    for i, name in enumerate(series, start=3):
        wsCfg.cell(i, 36, name)
        wsCfg.cell(i, 37, 1)
        wsCfg.cell(i, 38, 0 if name == "Резерв" else 1)

    wsCfg["W19"] = "Готовый .xlsm v1.8.6 — макросы внутри"
    wsCfg["W19"].font = hdr_font
    wsCfg["W20"] = "Любая станция/событие: замените RawData и Config."
    wsCfg["W21"] = "Кнопка «Запустить мониторинг ОПРЧ» или Alt+F8 → AnalyzeOPRCH."
    wsCfg["W22"] = "НЕ запускать SetupOPRCHTemplate — сотрёт данные."

    for col in range(1, 39):
        wsCfg.column_dimensions[get_column_letter(col)].width = 14
    wsCfg.column_dimensions["A"].width = 24
    wsCfg.column_dimensions["E"].width = 16
    wsCfg.column_dimensions["W"].width = 55

    wsSum["A1"] = "После AnalyzeOPRCH здесь сводная таблица."
    wsLog["A1"] = "Журнал появится после AnalyzeOPRCH."

    out = ROOT / "Шаблон_мониторинг_ОПРЧ.xlsx"
    wb.save(out)

    tmp = out.with_suffix(".tmp.xlsx")
    with ZipFile(out, "r") as zin, ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/workbook.xml":
                data = data.decode("utf-8").replace("<workbookProtection/>", "").encode("utf-8")
            elif info.filename == "docProps/app.xml":
                data = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                    "<Application>Microsoft Excel</Application>"
                    "<AppVersion>16.0300</AppVersion>"
                    "</Properties>"
                ).encode("utf-8")
            zout.writestr(info, data)
    tmp.replace(out)
    print("xlsx rebuilt", out, out.stat().st_size)
    return out


def build_xlsm_with_macros(xlsx: Path, modules: dict[str, str] | None = None) -> Path:
    """Собрать готовый .xlsm: данные из xlsx + VBA через pyOpenVBA (без Excel SaveAs)."""
    from pyopenvba import ExcelFile

    xlsm = ROOT / "Шаблон_мониторинг_ОПРЧ.xlsm"
    shell = Path(tempfile.gettempdir()) / "oprch_vba_shell.xlsm"
    if shell.exists():
        shell.unlink()
    if xlsm.exists():
        xlsm.unlink()

    names = [f"OPRCH_{m}" for m in MODULE_ORDER]

    def _module_src(name: str) -> str:
        utf = ROOT / f"{name}.bas"
        if not utf.exists():
            raise FileNotFoundError(utf)
        src = utf.read_text("utf-8").replace("\r\n", "\n").replace("\r", "\n")
        if "Attribute VB_Name" not in "\n".join(src.splitlines()[:3]):
            src = f'Attribute VB_Name = "{name}"\n' + src
        return src

    ef = ExcelFile.create_new(shell)
    proj = ef.vba_project()
    # Русский исходник должен кодироваться как Windows-1251.
    proj.code_page = 1251
    for drop in ("Module1", "Sheet1"):
        if drop in proj.module_names():
            try:
                proj.delete_module(drop)
            except Exception as e:
                print("warn delete", drop, e)

    for name in names:
        src = _module_src(name)
        print(f"embed {name}: {len(src)} chars")
        proj.add_module(name, src)
    ef.save()

    # pyOpenVBA оставляет в DIR code page 1252 от шаблона — из‑за этого Excel Mac
    # показывает «An error occurred while loading 'OPRCH_Sheets'». Патчим DIR и
    # перезаписываем исходники уже при code_page=1251.
    ef = ExcelFile(shell)
    proj = ef.vba_project()
    d = bytearray(proj.dir_raw)
    needle = b"\x03\x00\x02\x00\x00\x00"  # PROJECTCODEPAGE
    idx = d.find(needle)
    if idx < 0:
        raise RuntimeError("В vbaProject не найден PROJECTCODEPAGE")
    d[idx + 6 : idx + 8] = (1251).to_bytes(2, "little")
    proj.dir_raw = bytes(d)
    proj.code_page = 1251
    proj.dir_structure_dirty = True
    for name in names:
        ef.set_module(name, _module_src(name))
    ef.save()

    check0 = ExcelFile(shell)
    if check0.vba_project().code_page != 1251:
        raise RuntimeError(f"code_page={check0.vba_project().code_page}, ожидался 1251")
    if "Авария" not in check0.get_module("OPRCH_Sheets"):
        raise RuntimeError("Кириллица в OPRCH_Sheets повреждена при сборке xlsm")

    # Автоинициализация кнопок при открытии книги
    tw = check0.get_module("ThisWorkbook")
    if "InitOPRCHUISilent" not in tw:
        if not tw.endswith("\n"):
            tw += "\n"
        tw += (
            "\nPrivate Sub Workbook_Open()\n"
            "    On Error Resume Next\n"
            "    Application.Run \"InitOPRCHUISilent\"\n"
            "    On Error GoTo 0\n"
            "End Sub\n"
        )
        check0.set_module("ThisWorkbook", tw)
        check0.save()
        check0 = ExcelFile(shell)

    if "InitOPRCHUI" not in check0.get_module("OPRCH_Monitor"):
        raise RuntimeError("В OPRCH_Monitor нет InitOPRCHUI")

    with ZipFile(shell, "r") as z:
        vba = z.read("xl/vbaProject.bin")
    print("vbaProject.bin", len(vba), "code_page", check0.vba_project().code_page)

    def patch_ct(data: bytes) -> bytes:
        s = data.decode("utf-8")
        s = s.replace(
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
            'ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"',
        )
        if "vbaProject" not in s:
            s = s.replace(
                "</Types>",
                '<Override PartName="/xl/vbaProject.bin" '
                'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
            )
        return s.encode("utf-8")

    def patch_rels(data: bytes) -> bytes:
        s = data.decode("utf-8")
        if "vbaProject" in s:
            return data
        ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', s)]
        nid = (max(ids) + 1) if ids else 1
        rel = (
            f'<Relationship Id="rId{nid}" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/>'
        )
        return s.replace("</Relationships>", rel + "</Relationships>").encode("utf-8")

    with ZipFile(xlsx, "r") as zin, ZipFile(xlsm, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = patch_ct(data)
            elif info.filename == "xl/_rels/workbook.xml.rels":
                data = patch_rels(data)
            zout.writestr(info, data)
        zout.writestr("xl/vbaProject.bin", vba)

    check = ExcelFile(xlsm)
    missing = [n for n in names if n not in check.module_names()]
    if missing:
        raise RuntimeError(f"В xlsm нет модулей: {missing}")
    with ZipFile(xlsm) as z:
        vsz = z.getinfo("xl/vbaProject.bin").file_size
    print("xlsm ready", xlsm, xlsm.stat().st_size, "vba", vsz, "modules", check.module_names())
    if vsz < 20000:
        raise RuntimeError("vbaProject.bin слишком мал")
    return xlsm


def main():
    text = MASTER.read_text("utf-8")
    # bump version stamp in comment if needed
    if 'OPRCH_VERSION As String = "1.8.6"' not in text:
        raise SystemExit("Unexpected version in master")

    # syntax sanity: unmatched quotes in slim is hard; check basic patterns
    for bad in ["\u0394", "\u2192", "\u2248", "\u2264", "\u2265"]:
        if bad in text:
            raise SystemExit(f"Non-cp1251 char {bad!r} in master")

    modules = split_master(text)
    write_bas(modules)
    # Installer win1251 (на случай ручной переустановки)
    inst = (ROOT / "OPRCH_Install.bas").read_text("utf-8")
    (ROOT / "OPRCH_Install_win1251.bas").write_bytes(
        inst.replace("\n", "\r\n").encode("cp1251")
    )
    rows = load_raw_from_template()
    print("raw rows", len(rows))
    xlsx = rebuild_xlsx(rows)
    xlsm = build_xlsm_with_macros(xlsx)
    print("DONE", xlsm)


if __name__ == "__main__":
    main()
