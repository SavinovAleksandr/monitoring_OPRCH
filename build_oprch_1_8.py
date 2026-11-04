#!/usr/bin/env python3
"""Split OPRCH_Monitor_all.bas → 5 modules, rebuild xlsx, try build xlsm via Excel."""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path("/Users/asavinov/ОПРЧ/ОПРЧ")
MASTER = ROOT / "OPRCH_Monitor_all.bas"

MONITOR_PROCS = {
    "SetupOPRCHTemplate", "ExportOPRCHReport", "AnalyzeOPRCH",
    "ClearOPRCHResults", "ApplyPresetsToConfig",
    "EnsureControlButtons", "AddOrReplaceButton", "WriteInstructionSheet",
}
ENGINE_PROCS = {
    "AnalyzeOneGenerator", "EvaluateReaction15", "React15ThresholdMW",
    "ClassifyReaction15", "EvaluateDeficit", "BuildGeneratorVerdict",
    "EvaluateApplicability", "SelectEvaluationPoint", "EvaluateOscillations",
    "EvaluateTelemetry", "ClassifyReactionType",
    "EvaluateQualitative", "ComputeSteadyMean", "ComputeSteadyPReqMean",
    "EvaluateTransient", "FindEvents", "AvgPowerBackward",
    "BuildStationVerdict", "ResolveF0",
}
# IdentifyH15 + SolveLinearSystem -> Common (разгрузка Engine под лимит ~64 КБ)
CHARTS_PROCS = {
    "WriteGeneratorChartSheet", "ApplyChartAxesFormat",
    "GetRangeMinMaxByCols", "MaxAbsInCols", "WriteChartVerdictBlock", "FindChartEndRow",
    "WriteStationChartSheet", "AddSeries", "AddMarkerSeries", "AddLimitLineSeries",
    "FillMarkerColumn", "FillMarkerColumnAtTime", "ApplyGeneratorSheetFormats",
    "LoadChartDisplayOptions", "ChartOpt", "EnsureChartOptionCheckboxes",
}
SHEETS_PROCS = {
    "WriteGeneratorSheet", "FillMarkerColumn", "FillMarkerColumnAtTime",
    "ApplyGeneratorSheetFormats", "WriteGeneratorChartSheet", "ApplyChartAxesFormat",
    "GetRangeMinMaxByCols", "MaxAbsInCols", "WriteChartVerdictBlock", "FindChartEndRow",
    "BuildStationAggregates", "BuildParoprovodSubAggregates", "SumPowerAtRow",
    "BuildOneStationAggregate", "BuildOneStationAggregateEvent",
    "AnnotateGeneratorRowsWithStationVerdict",
    "StationSheetName", "StationChartSheetName", "WriteStationChartSheet",
    "WriteSummaryRow", "WriteSummaryInvalid", "WriteSummaryStationRow",
    "ApplySummaryConditionalFormat", "ApplyStatusCF", "WriteVersionStamp",
    "InitLog", "AppendLog", "FinalizeLog", "ValidateInputs",
    "AddSeries", "AddMarkerSeries", "AddLimitLineSeries",
    "WriteProtocolSheet", "AddProtocolLine", "AddProtocolGenerator", "AddProtocolStation",
    "EnsureChartOptionCheckboxes", "LoadChartDisplayOptions", "ChartOpt",
    "CollectOldOutputSheets", "DeleteOutputSheets",
    "GeneratorSheetName", "GeneratorChartSheetName",
}

PROC_RE = re.compile(r"^(Public |Private )(Sub |Function )([A-Za-z0-9_]+)")


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
    buckets = {"Common": [], "Engine": [], "Sheets": [], "Charts": [], "Monitor": []}
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
        elif name in CHARTS_PROCS:
            buckets["Charts"].append(block)
        elif name in SHEETS_PROCS:
            buckets["Sheets"].append(block)
        else:
            buckets["Common"].append(block)

    headers = {
        "Common": "OPRCH_Common",
        "Engine": "OPRCH_Engine",
        "Sheets": "OPRCH_Sheets",
        "Charts": "OPRCH_Charts",
        "Monitor": "OPRCH_Monitor",
    }
    notes = {
        "Common": "Типы, константы, общие функции. Импортировать 1.",
        "Engine": "Расчёт участия в ОПРЧ. Импортировать 2.",
        "Sheets": "Summary, протокол, станционные листы. Импортировать 3.",
        "Charts": "Графики генераторов и станции. Импортировать 4.",
        "Monitor": "Точки входа AnalyzeOPRCH / SetupOPRCHTemplate. Импортировать 5.",
    }
    out = {}
    for key, attr in headers.items():
        body = [
            f'Attribute VB_Name = "{attr}"',
            "Option Explicit",
            "",
            f"' {notes[key]}",
            "' OPRCH_Monitor v1.8.0 — часть раздельного модуля (лимит VBA ~64 КБ на модуль).",
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
        nbytes = len(text_out.encode("cp1251"))
        print(f"{attr:16} lines={text_out.count(chr(10)):4}  cp1251={nbytes:6}")
        if nbytes > 60000:
            print("  WARN approaching 64K source")
    return out


def write_bas(modules: dict[str, str]) -> list[Path]:
    mapping = {
        "Common": "OPRCH_Common.bas",
        "Engine": "OPRCH_Engine.bas",
        "Sheets": "OPRCH_Sheets.bas",
        "Charts": "OPRCH_Charts.bas",
        "Monitor": "OPRCH_Monitor.bas",
    }
    win_paths = []
    for key, name in mapping.items():
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


NS = {"m": "http://purl.oclc.org/ooxml/spreadsheetml/main"}


def col_row(cell_ref):
    m = re.match(r"([A-Z]+)(\d+)", cell_ref)
    col, row = m.group(1), int(m.group(2))
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n, row


def parse_time(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if "T" in s:
        s = s.replace("Z", "")
        if "." in s:
            a, b = s.split(".", 1)
            b = re.sub(r"[^0-9].*", "", b)[:6]
            s = a + "." + b if b else a
        try:
            return datetime.fromisoformat(s)
        except Exception:
            return None
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(s))
    except Exception:
        return None


def load_raw_from_template() -> list[dict]:
    src = ROOT / "Шаблон_мониторинг_ОПРЧ_ВТЭЦ-2_31.10.2026.xlsx"
    if not src.exists():
        src = Path("/Users/asavinov/Downloads/Шаблон_мониторинг ОПРЧ.xlsx")
    wb = load_workbook(src, data_only=True)
    ws = wb["RawData"]
    rows = []
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, 1).value
        if t is None:
            continue
        if not isinstance(t, datetime):
            t = parse_time(t)
            if t is None:
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
    wsQ = wb.create_sheet("Запрос")
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

    # Инструкция
    lines = [
        ("Мониторинг ОПРЧ v1.8.0 — что вводить и как считать", True, 14),
        ("", False, 11),
        ("1. ИСХОДНЫЕ ДАННЫЕ (обязательно)", True, 12),
        ("• Лист RawData: колонка «Время» (шаг ≤ 5 с) + частота станции + активные мощности генераторов (МВт).", False, 11),
        ("• Окно данных по п. 3.1 МУ: желательно ≥ 5 мин до начала отклонения и до устойчивого возврата в 50±0.1 Гц более 5 мин.", False, 11),
        ("• Лист Config: станция, генератор, заголовки колонок мощности/частоты, тип, Pном, S, fнч, РМ (1/0), компоновка.", False, 11),
        ("• Пустые нормативные поля заполняются кнопкой «Применить пресеты типов».", False, 11),
        ("• Глобальные настройки — колонки W/X: режим Авария/Испытание, fном, автопоиск, строгий дефицит (W17).", False, 11),
        ("", False, 11),
        ("2. ОДИН ЗАПУСК = ВСЁ СРАЗУ", True, 12),
        ("Кнопка «Запустить мониторинг ОПРЧ» выполняет количественную и качественную оценку в одном проходе.", False, 11),
        ("Отдельно «только качественно» / «только количественно» запускать не нужно.", False, 11),
        ("• Колич. (Прил. 14 п. 4.2): P0, f0, Pтреб(t), tр, дефицит по периодам, реакция 15 с, сумма неблока.", False, 11),
        ("• Кач. (Прил. 14 п. 4.4): тип реакции 1–8. Ступени t5/t10 (Прил. 1) — при Кач_вкл=1.", False, 11),
        ("", False, 11),
        ("3. РЕЖИМЫ (Config W13)", True, 12),
        ("• Авария — Прил. 14. Без РМ: S=6 %, fнч=0.15. Дефицит в любом периоде → НАРУШЕНИЕ. t5/t10 справочно.", False, 11),
        ("• Испытание — Прил. 1 + формулы Прил. 14. S/fнч из Config. Несданные t5/t10 → НАРУШЕНИЕ.", False, 11),
        ("", False, 11),
        ("4. ЗАПУСК", True, 12),
        ("Откройте готовый .xlsm → разрешите макросы → Config → «Запустить мониторинг ОПРЧ».", False, 11),
        ("Не запускайте SetupOPRCHTemplate на этом шаблоне — сотрёт RawData/Config.", False, 11),
        ("Если макросов нет — импортируйте 5 файлов *_win1251.bas (см. КАК_ИМПОРТИРОВАТЬ_VBA.txt).", False, 11),
    ]
    for i, (txt, bold, size) in enumerate(lines, 1):
        cell = wsInstr.cell(i, 1, txt)
        cell.font = Font(bold=bold, size=size)
    wsInstr.column_dimensions["A"].width = 130

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

    wsCfg["W19"] = "Запуск v1.8.0"
    wsCfg["W19"].font = hdr_font
    wsCfg["W20"] = "Готовый файл: Шаблон_мониторинг_ОПРЧ_ВТЭЦ-2_31.10.2026.xlsm (макросы уже внутри)."
    wsCfg["W21"] = "Если открыли .xlsx — сохраните как .xlsm и импортируйте *_win1251.bas (см. Инструкция)."
    wsCfg["W22"] = "НЕ запускать SetupOPRCHTemplate — сотрёт данные."
    wsCfg["W23"] = "Воркутинская ТЭЦ-2, 31.10.2026 17:00-17:08, режим Авария, без РМ."

    for col in range(1, 39):
        wsCfg.column_dimensions[get_column_letter(col)].width = 14
    wsCfg.column_dimensions["A"].width = 24
    wsCfg.column_dimensions["E"].width = 16
    wsCfg.column_dimensions["W"].width = 55

    query = [
        "Код запроса архива ОИК — частотная авария 31.10.2026, Воркутинская ТЭЦ-2",
        "",
        "Объект: Воркутинская ТЭЦ-2 (общий паропровод)",
        "Интервал: 31.10.2026 16:55:00 — 31.10.2026 17:15:00",
        "Шаг: 1 с",
        "",
        "FROM  2026-10-31 16:55:00",
        "TO    2026-10-31 17:15:00",
        "STEP  1s",
        "SIGNALS",
        '  "Воркутинская ТЭЦ-2.Частота"  AS [f ВТЭЦ-2]',
        '  "Воркутинская ТЭЦ-2.ТГ-1.P"   AS [ТГ-1]',
        '  "Воркутинская ТЭЦ-2.ТГ-3.P"   AS [ТГ-3]',
        '  "Воркутинская ТЭЦ-2.ТГ-4.P"   AS [ТГ-4]',
        '  "Воркутинская ТЭЦ-2.ТГ-7.P"   AS [ТГ-7]',
    ]
    wsQ["A1"] = query[0]
    wsQ["A1"].font = Font(bold=True, size=12)
    for i, line in enumerate(query[1:], start=2):
        wsQ.cell(i, 1, line)
    wsQ.column_dimensions["A"].width = 80

    wsSum["A1"] = "После запуска AnalyzeOPRCH здесь появится сводная таблица."
    wsLog["A1"] = "Журнал появится после запуска AnalyzeOPRCH."

    out = ROOT / "Шаблон_мониторинг_ОПРЧ_ВТЭЦ-2_31.10.2026.xlsx"
    wb.save(out)

    tmp = out.with_suffix(".tmp.xlsx")
    with ZipFile(out, "r") as zin, ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/workbook.xml":
                s = data.decode("utf-8").replace("<workbookProtection/>", "")
                data = s.encode("utf-8")
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


def build_xlsm_via_excel(xlsx: Path, win_bas: list[Path]) -> Path | None:
    """Save as xlsm and import VBA modules through Microsoft Excel AppleScript."""
    xlsm = xlsx.with_suffix(".xlsm")
    if xlsm.exists():
        xlsm.unlink()

    # Escape paths for AppleScript
    def aq(p: Path) -> str:
        return str(p).replace("\\", "\\\\").replace('"', '\\"')

    import_lines = "\n".join(
        f'      do Visual Basic "ThisWorkbook.VBProject.VBComponents.Import \\"{aq(p)}\\""'
        for p in win_bas
    )

    script = f'''
tell application "Microsoft Excel"
  activate
  set wb to open workbook workbook file name "{aq(xlsx)}"
  delay 1
  set dest to "{aq(xlsm)}"
  save workbook as wb filename dest file format Excel macro enabled workbook file format with overwrite
  close wb saving no
  delay 1
  set wb2 to open workbook workbook file name dest
  delay 1
  try
{import_lines}
    do Visual Basic "Application.Run \\"WriteInstructionSheet\\""
  on error errMsg
    -- Trust access to VBA project may be off
    display dialog "VBA import failed: " & errMsg buttons {{"OK"}}
  end try
  save workbook wb2
  -- leave open for user, or close:
  -- close wb2 saving yes
end tell
'''
    script_path = ROOT / "_build_xlsm.applescript"
    script_path.write_text(script, encoding="utf-8")
    print("Running Excel AppleScript to build xlsm...")
    try:
        r = subprocess.run(
            ["osascript", str(script_path)],
            capture_output=True,
            text=True,
            timeout=180,
        )
        print("osascript rc=", r.returncode)
        if r.stdout:
            print("stdout:", r.stdout[:500])
        if r.stderr:
            print("stderr:", r.stderr[:800])
    except Exception as e:
        print("AppleScript failed:", e)
        return None

    if xlsm.exists() and xlsm.stat().st_size > 10000:
        # Check if vbaProject.bin present
        try:
            with ZipFile(xlsm) as z:
                has_vba = "xl/vbaProject.bin" in z.namelist()
            print("xlsm created", xlsm, "size", xlsm.stat().st_size, "has_vba", has_vba)
            return xlsm if has_vba else xlsm
        except Exception as e:
            print("xlsm zip check failed", e)
            return xlsm
    print("xlsm not created or too small")
    return None


def main():
    text = MASTER.read_text("utf-8")
    if 'OPRCH_VERSION As String = "1.8.0"' not in text:
        raise SystemExit("Master is not v1.8.0")
    modules = split_master(text)
    win_bas = write_bas(modules)
    rows = load_raw_from_template()
    print("raw rows", len(rows), rows[0]["t"], rows[-1]["t"])
    xlsx = rebuild_xlsx(rows)
    xlsm = build_xlsm_via_excel(xlsx, win_bas)
    print("DONE xlsx=", xlsx)
    print("DONE xlsm=", xlsm)


if __name__ == "__main__":
    main()
