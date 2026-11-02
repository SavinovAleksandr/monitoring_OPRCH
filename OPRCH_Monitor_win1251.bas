Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' ==========================================================
' OPRCH monitoring  (quantitative + qualitative)
' Per-generator:  data sheet <Ст>_<Ген>        +  chart sheet <Ст>_<Ген>_Граф
' Per-station:    data sheet <Ст>_Сумма[_<Пп>] +  chart sheet <...>_Граф
' Extra sheets:   Summary, Log
'
' Версия: 1.6.0
' Нормативная база: Методические указания по оценке участия генерирующего
'   оборудования в ОПРЧ (Приложение 14 к Техтребованиям, ред. 01.03.2026):
'   п. 4.2.1-4.2.3 (формулы dPп, Pтп, dFр), 4.2.6 (усреднение P/f),
'   4.2.7-4.2.8 (допуск 1 %Pном), 4.2.10 (оценка по сумме станции ->
'   индивидуальная оценка только при нарушении), 4.2.11 (наличие реакции
'   РЧВ турбины через 15 с), 4.2.14.1 (дефицит хотя бы в один из периодов),
'   Приложение (Kд(t): ТЭС 0.5 до 15 с -> 1.0 к 300/360/420 с;
'   ПГУ 0.25 до 15 с -> 0.5 к 30 с -> 1.0 к 120 с).
' Основные критерии:
'   - Количественный (|Pфакт|/|Pтреб|) с допуском и направлением;
'   - Перерегулирование (>100%+допуск);
'   - Качественные подкритерии: t5, t10, установившееся (среднее по хвосту);
'   - Характер переходного процесса: Монотонный / Апериодический / Колебательный;
'   - Амплитуда возмущения в %Pном и метка масштаба события;
'   - Пресеты параметров по типу оборудования (ПТУ_блок / ПТУ_неблок /
'     ГТУ / ПГУ_утил / ПГУ_сбросн / ГПА) подставляются при пустых ячейках;
'   - Учет Pmax/Pmin (диапазон регулирования): капинг Pтреб по располагаемому
'     резерву, статус 'Ограничен Pmax/Pmin' в Summary, WARN при резерве <5 %Pном,
'     горизонтальные уровни Pmax/Pmin и цветная зона за лимитом на графике.
'     Для станционных сумм Pmax_сум/Pmin_сум = сумма по включенным генераторам
'     (по паропроводам У/Э для ТЭЦ СЛПК).
'   - Реакция РЧВ через 15 с (п. 4.2.11): отдельный статус в Summary;
'   - Дефицит первичной мощности по точкам окна (п. 4.2.14.1): макс. %Pном и длительность;
'   - Вердикт по сумме станции (п. 4.2.10) со строкой в Summary и выводом о
'     необходимости индивидуальной оценки генераторов.
' ==========================================================

Public Const OPRCH_VERSION As String = "1.6.0"

Private Const SH_RAW As String = "RawData"
Private Const SH_CFG As String = "Config"
Private Const SH_SUM As String = "Summary"
Private Const SH_LOG As String = "Log"
Private Const CHART_SUFFIX As String = "_Граф"

Private m_LogRow As Long
Private m_KdProfiles As Object   ' key=EQUIPTYPE, value=Array(t0,kd0,t1,kd1,t2,kd2,t3,kd3)
Private m_ChartOpts As Object    ' key=REL_*/ABS_*, value=Boolean

Private Type TSettings
    FNom As Double
    EventStart As Double
    AutoStart As Boolean
    QuantIntervalSec As Double
    QuantTolPct As Double
    WorkThresholdMW As Double
    PreBufferSec As Double
    ChartIntervalSec As Double
    SteadyWindowSec As Double
    React15ThrPct As Double   ' порог наличия реакции РЧВ через 15 с, %Pном (п. 4.2.11)
    AvgWindowSec As Double    ' окно усреднения P0/Pтек, с (п. 4.2.6); 0 = текущие значения
End Type

Private Type TGenCfg
    Station As String
    Generator As String
    PowerHeader As String
    FreqHeader As String
    EquipType As String
    PNom As Double
    SPct As Double
    Fnch As Double
    Enabled As Boolean
    QualEnabled As Boolean
    T5Sec As Double
    Dp5Pct As Double
    T10Sec As Double
    Dp10Pct As Double
    SteadyTolPct As Double
    InStationSum As Boolean
    CheckSteady As Boolean
    Paroprovod As String
    PMax As Double          ' максимальная эксплуатационная мощность, МВт
    PMin As Double          ' технический минимум, МВт
End Type

Private Type TGenResult
    StartRow As Long
    EndQuantRow As Long
    EndQualRow As Long
    StartTime As Variant
    FirstExceedTime As Variant
    P0 As Double
    PTek As Double
    PsteadyAvg As Double
    Df As Double
    Dfr As Double
    PReq As Double
    PReqOrig As Double      ' исходный требуемый, до ограничения Pmax/Pmin
    PFact As Double
    AmplPctPnom As Double
    AmplitudeTag As String
    QuantPct As Double
    QuantPass As Boolean
    Overshoot As Boolean
    TransientType As String
    NumExtrema As Long
    PReqSteady As Double
    QualPass As Boolean
    QualT5Pass As Boolean
    QualT10Pass As Boolean
    QualSteadyPass As Boolean
    T5FactSec As Double
    T10FactSec As Double
    QualFailedList As String
    QualReason As String
    PMaxEff As Double       ' эффективная Pmax (g.PMax или g.PNom по умолчанию)
    PMinEff As Double       ' эффективная Pmin (g.PMin или 0)
    ReservePlus As Double   ' располагаемый резерв '+' = max(0, PMax - P0)
    ReserveMinus As Double  ' располагаемый резерв '-' = max(0, P0 - PMin)
    Limited As Boolean      ' Pтреб был ограничен диапазоном
    LimitType As String     ' 'Pmax' или 'Pmin'
    KdUsedQuant As Double   ' Kд, принятый в количественном расчёте
    KdProfile As String     ' справка по применённому профилю Kд(t)
    PReqStatic As Double    ' Pтреб при Kд=1 (статическая характеристика), с учётом резерва
    React15Status As String ' "ОК" / "Нарушение" / "Н/П" (не применимо) - п. 4.2.11
    React15Pass As Boolean
    DP15 As Double          ' изменение мощности через 15 с после выхода за fнч, МВт
    React15Note As String
    DeficitMaxPct As Double ' макс. дефицит первичной мощности в окне, %Pном (п. 4.2.14.1)
    DeficitSec As Double    ' суммарная длительность дефицита, с
    Verdict As String       ' итоговый вывод по генератору
End Type

' Результат оценки по сумме станции (п. 4.2.10)
Private Type TStationResult
    Station As String
    Paroprovod As String
    Count As Long
    StartTime As Variant
    FirstExceedTime As Variant
    P0 As Double
    PTek As Double
    Df As Double
    PReq As Double
    PReqOrig As Double
    PFact As Double
    QuantPct As Double
    QuantPass As Boolean
    Limited As Boolean
    LimitType As String
    PMaxSum As Double
    PMinSum As Double
    ReservePlus As Double
    ReserveMinus As Double
    PNomSum As Double
    React15Status As String
    React15Pass As Boolean
    DP15 As Double
    React15Note As String
    DeficitMaxPct As Double
    DeficitSec As Double
    Verdict As String
    SheetName As String
    ChartSheetName As String
End Type

' ==========================================================
' Точки входа
' ==========================================================

Public Sub SetupOPRCHTemplate()
    Dim wsRaw As Worksheet, wsCfg As Worksheet, wsSum As Worksheet

    Set wsRaw = EnsureSheet(SH_RAW)
    Set wsCfg = EnsureSheet(SH_CFG)
    Set wsSum = EnsureSheet(SH_SUM)

    wsRaw.Cells.Clear
    wsCfg.Cells.Clear
    wsSum.Cells.Clear

    wsRaw.Range("A1").Value = "Время"
    wsRaw.Range("B1").Value = "Частота"
    wsRaw.Range("C1").Value = "ТГ-1"
    wsRaw.Range("D1").Value = "ТГ-2"
    wsRaw.Range("E1").Value = "ТГ-3"

    wsCfg.Range("A1:T1").Value = Array( _
        "Станция", "Генератор", "Колонка_мощности", "Колонка_частоты", "Тип_оборудования", _
        "Pном, МВт", "S, %", "fнч, Гц", "Вкл (1/0)", "Кач_вкл (1/0)", _
        "t5, c", "dP5, %Pном", "t10, c", "dP10, %Pном", "Уст_допуск, %Pном", _
        "В сумму станции (1/0)", "Контр_уст (1/0)", "Паропровод", _
        "Pmax, МВт", "Pmin, МВт" _
    )

    wsCfg.Cells(2, 1).Resize(1, 20).Value = Array("Сосногорская ТЭЦ", "ТГ-5", "ТГ-5", "Частота", "ПТУ_неблок", 55, 4.2, 0.105, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 55, 15)
    wsCfg.Cells(3, 1).Resize(1, 20).Value = Array("Сосногорская ТЭЦ", "ТГ-7", "ТГ-7", "Частота", "ПТУ_неблок", 60, 4.5, 0.11, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 60, 20)
    wsCfg.Cells(4, 1).Resize(1, 20).Value = Array("ТЭЦ СЛПК", "ТГ-2Э", "ТГ-2Э", "Частота", "ПТУ_неблок", 50, 4.5, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "Э", 50, 15)
    wsCfg.Cells(5, 1).Resize(1, 20).Value = Array("ТЭЦ СЛПК", "ТГ-5У", "ТГ-5У", "Частота", "ПТУ_неблок", 87.7, 4.2, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "У", 87.7, 25)

    wsCfg.Range("W1").Value = "Глобальные настройки"
    wsCfg.Cells(2, 23).Resize(1, 2).Value = Array("fном, Гц", 50)
    wsCfg.Cells(3, 23).Resize(1, 2).Value = Array("Время начала события", "")
    wsCfg.Cells(4, 23).Resize(1, 2).Value = Array("Автопоиск старта (1/0)", 1)
    wsCfg.Cells(5, 23).Resize(1, 2).Value = Array("Колич. интервал, с", 82)
    wsCfg.Cells(6, 23).Resize(1, 2).Value = Array("Допуск количеств., %", 10)
    wsCfg.Cells(7, 23).Resize(1, 2).Value = Array("Порог включения в работу, МВт", 1)
    wsCfg.Cells(8, 23).Resize(1, 2).Value = Array("Pre-start буфер, с", 5)
    wsCfg.Cells(9, 23).Resize(1, 2).Value = Array("Интервал графика, с", 120)
    wsCfg.Cells(10, 23).Resize(1, 2).Value = Array("Окно установив., с", 30)
    wsCfg.Cells(11, 23).Resize(1, 2).Value = Array("Порог реакции 15 с, %Pном", 0.5)
    wsCfg.Cells(12, 23).Resize(1, 2).Value = Array("Усреднение P0/Pтек, с (0=текущие)", 0)

    ' Профили Kд(t) по Приложению к МУ оценки участия в ОПРЧ (Приложение 14 к Техтребованиям):
    '   ТЭС (кроме ПГУ): Kд = 0.5 на [0;15] с, далее линейно до 1.0 к T
    '       (T = 300 с газомазутные блоки, 360 с пылеугольные, 420 с общий паропровод);
    '   ПГУ/ГТУ: Kд = 0.25 на [0;15] с, 0.5 к 30 с, 1.0 к 120 с;
    '   ГПА: нормируется интервалом 1 мин (10 %Pном к 60 с) - ориентир.
    wsCfg.Range("AA1").Value = "Профили Kд(t) (абсолютные значения Kд)"
    wsCfg.Range("AA2:AI2").Value = Array("Тип_оборудования", "t0, с", "Kд0", "t1, с", "Kд1", "t2, с", "Kд2", "t3, с", "Kд3")
    wsCfg.Cells(3, 27).Resize(1, 9).Value = Array("ПТУ_блок", 0, 0.5, 15, 0.5, 300, 1, 600, 1)
    wsCfg.Cells(4, 27).Resize(1, 9).Value = Array("ПТУ_блок_уголь", 0, 0.5, 15, 0.5, 360, 1, 600, 1)
    wsCfg.Cells(5, 27).Resize(1, 9).Value = Array("ПТУ_неблок", 0, 0.5, 15, 0.5, 420, 1, 600, 1)
    wsCfg.Cells(6, 27).Resize(1, 9).Value = Array("ГТУ", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1)
    wsCfg.Cells(7, 27).Resize(1, 9).Value = Array("ПГУ_утил", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1)
    wsCfg.Cells(8, 27).Resize(1, 9).Value = Array("ПГУ_сбросн", 0, 0.5, 15, 0.5, 300, 1, 600, 1)
    wsCfg.Cells(9, 27).Resize(1, 9).Value = Array("ГПА", 0, 0.5, 15, 0.5, 60, 1, 600, 1)

    wsCfg.Range("AJ1").Value = "Отображение рядов графиков (1/0)"
    wsCfg.Range("AJ2:AL2").Value = Array("Ряд", "Относит.", "Абсолют.")
    wsCfg.Range("AJ3:AJ10").Value = WorksheetFunction.Transpose(Array( _
        "Pфакт", "Pтреб", "+Допуск", "-Допуск", "Pmax/Pmin", "Частота", "Маркеры t5/t10/fнч", "Резерв" _
    ))
    wsCfg.Range("AK3:AL10").Value = 1
    wsCfg.Range("AK10:AL10").Value = 0

    wsCfg.Columns("A:AL").AutoFit
    wsRaw.Columns("A:E").AutoFit
    EnsureControlButtons wsCfg
    EnsureChartOptionCheckboxes wsCfg

    MsgBox "Шаблон создан. Заполните RawData/Config и нажмите кнопку 'Запустить мониторинг ОПРЧ'.", vbInformation
End Sub

Public Sub AnalyzeOPRCH()
    On Error GoTo EH

    Dim wsRaw As Worksheet, wsCfg As Worksheet, wsSummary As Worksheet
    Dim st As TSettings
    Dim timeCol As Long, cfgLast As Long, r As Long, outRow As Long
    Dim g As TGenCfg, res As TGenResult
    Dim targetSheets As Collection
    Dim stepName As String
    Dim t0Run As Double

    t0Run = Timer
    stepName = "Подготовка листов"
    Set wsRaw = GetRequiredSheet(SH_RAW)
    Set wsCfg = GetRequiredSheet(SH_CFG)
    Set wsSummary = EnsureSheet(SH_SUM)

    stepName = "Чтение настроек"
    st = ReadSettings(wsCfg)
    LoadKdProfiles wsCfg
    LoadChartDisplayOptions wsCfg
    EnsureChartOptionCheckboxes wsCfg
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Err.Raise vbObjectError + 2001, , "В RawData не найдена колонка 'Время'."

    cfgLast = LastUsedRow(wsCfg)
    If cfgLast < 2 Then Err.Raise vbObjectError + 2002, , "В Config нет строк генераторов."

    stepName = "Подготовка лога"
    InitLog

    stepName = "Валидация Config/RawData"
    ValidateInputs wsRaw, wsCfg, st, timeCol

    stepName = "Очистка выходных листов"
    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    DeleteOutputSheets targetSheets

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    stepName = "Подготовка Summary"
    wsSummary.Cells.Clear
    wsSummary.Range("A1:AN1").Value = Array( _
        "Станция", "Генератор", "Тип", _
        "Старт (расч.)", "Время выхода за fнч", _
        "P0, МВт", "Pтек, МВт", "Pуст_сред, МВт", _
        "dF, Гц", "dFr, Гц", _
        "Pтреб, МВт", "Pфакт, МВт", _
        "Амплитуда, %Pном", "Масштаб события", _
        "Колич. %", "Колич. статус", "Перерегулирование", _
        "Характер процесса", "Экстремумов", _
        "Кач. статус", "Кач.t5", "Кач.t10", "Кач.уст", _
        "Проваленные подпункты", "t5 факт, c", "t10 факт, c", _
        "Лист", "Лист графика", _
        "Pmax, МВт", "Pmin, МВт", "Резерв '+', МВт", "Резерв '-', МВт", _
        "Pтреб исх., МВт", "Ограничение", _
        "Примечание", _
        "Реакция 15 с", "dP15, МВт", "Дефицит макс, %Pном", "Длит. дефицита, с", _
        "Вывод" _
    )
    outRow = 2

    For r = 2 To cfgLast
        stepName = "Чтение Config, строка " & r
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextGen

        g = ReadGenCfg(wsCfg, r)
        If Not g.Enabled Then GoTo NextGen

        If Not ValidateGenCfg(g) Then
            WriteSummaryInvalid wsSummary, outRow, g
            AppendLog "Config", g.Station & "/" & g.Generator, "Пропущен: не заполнены обязательные параметры"
            outRow = outRow + 1
            GoTo NextGen
        End If

        stepName = "Расчет генератора " & g.Generator
        res = AnalyzeOneGenerator(wsRaw, st, g)
        stepName = "Запись листа данных " & g.Generator
        WriteGeneratorSheet wsRaw, st, g, res
        stepName = "Запись графика " & g.Generator
        WriteGeneratorChartSheet st, g, res
        stepName = "Запись Summary для " & g.Generator
        WriteSummaryRow wsSummary, outRow, g, res
        outRow = outRow + 1

NextGen:
    Next r

    stepName = "Расчет суммарных листов станций"
    BuildStationAggregates wsRaw, wsCfg, wsSummary, st, outRow

    stepName = "Оформление Summary"
    wsSummary.Columns("A:AN").AutoFit
    wsSummary.Columns("AN").ColumnWidth = 90
    wsSummary.Columns("AI").ColumnWidth = 60
    wsSummary.Range("D:E").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ApplySummaryConditionalFormat wsSummary
    WriteVersionStamp wsSummary, wsRaw, t0Run

    stepName = "Финализация лога и кнопок"
    FinalizeLog
    EnsureControlButtons wsCfg

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Мониторинг ОПРЧ завершен (v" & OPRCH_VERSION & "). Время: " _
        & Format(Timer - t0Run, "0.0") & " c.", vbInformation
    Exit Sub

EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Ошибка AnalyzeOPRCH (" & stepName & "): " & Err.Description, vbCritical
End Sub

Public Sub ClearOPRCHResults()
    Dim targetSheets As Collection
    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    Application.DisplayAlerts = False
    DeleteOutputSheets targetSheets
    Application.DisplayAlerts = True
    MsgBox "Результаты очищены (" & targetSheets.Count & " вкладок).", vbInformation
End Sub

Public Sub ApplyPresetsToConfig()
    Dim wsCfg As Worksheet, r As Long, cfgLast As Long, changed As Long
    Dim et As String, g As TGenCfg, pr As TGenCfg
    Set wsCfg = GetRequiredSheet(SH_CFG)
    cfgLast = LastUsedRow(wsCfg)
    changed = 0
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        et = Trim$(CStr(wsCfg.Cells(r, 5).Value))
        If Len(et) = 0 Then GoTo NX
        pr = GetPreset(et)
        If NzD(wsCfg.Cells(r, 11).Value, 0) <= 0 Then wsCfg.Cells(r, 11).Value = pr.T5Sec: changed = changed + 1
        If NzD(wsCfg.Cells(r, 12).Value, 0) <= 0 Then wsCfg.Cells(r, 12).Value = pr.Dp5Pct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 13).Value, 0) <= 0 Then wsCfg.Cells(r, 13).Value = pr.T10Sec: changed = changed + 1
        If NzD(wsCfg.Cells(r, 14).Value, 0) <= 0 Then wsCfg.Cells(r, 14).Value = pr.Dp10Pct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 15).Value, 0) <= 0 Then wsCfg.Cells(r, 15).Value = pr.SteadyTolPct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 8).Value, -1) < 0 Then wsCfg.Cells(r, 8).Value = pr.Fnch: changed = changed + 1
NX:
    Next r
    MsgBox "Пресеты применены. Заполнено ячеек: " & changed, vbInformation
End Sub

' ==========================================================
' Анализ одного генератора
' ==========================================================

Private Function AnalyzeOneGenerator(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg) As TGenResult
    On Error GoTo EH

    Dim res As TGenResult
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim startRow As Long, endQ As Long, endQual As Long, firstExceedRow As Long
    Dim p0 As Double, ptek As Double, df As Double, dfr As Double, preq As Double, pfact As Double
    Dim kdQuant As Double, tQuantSec As Double
    Dim qpct As Double, qpass As Boolean
    Dim calcStep As String

    calcStep = "Поиск колонок"
    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderCol(wsRaw, g.PowerHeader)
    fCol = FindHeaderCol(wsRaw, g.FreqHeader)
    If pCol = 0 Then Err.Raise vbObjectError + 2101, , "Не найдена колонка мощности '" & g.PowerHeader & "' для " & g.Generator
    If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator

    calcStep = "Определение стартовой строки"
    startRow = ResolveStartRow(wsRaw, timeCol, fCol, st, g.Fnch, firstExceedRow)
    calcStep = "Определение конца количественного интервала"
    endQ = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)
    calcStep = "Определение конца качественного интервала"
    endQual = RowByTimeOffset(wsRaw, timeCol, startRow, g.T10Sec)

    calcStep = "Чтение P0/Pтек"
    ' п. 4.2.6: значения P вычисляются как средние на интервале вблизи момента
    ' оценки; при стабильных значениях допускается брать текущие (AvgWindowSec = 0).
    p0 = AvgPowerBackward(wsRaw, pCol, timeCol, startRow, st.AvgWindowSec)
    ptek = AvgPowerBackward(wsRaw, pCol, timeCol, endQ, st.AvgWindowSec)
    calcStep = "Расчет dF/dFr"
    df = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, endQ, st.FNom)
    dfr = DeadbandDeviation(df, g.Fnch)

    calcStep = "Расчет требуемой мощности"
    Dim preqOrig As Double
    tQuantSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(endQ, timeCol).Value)
    kdQuant = DynamicKdByTime(g.EquipType, tQuantSec)
    If dfr <> 0 Then
        preqOrig = -100# / g.SPct * g.PNom / st.FNom * kdQuant * dfr
    Else
        preqOrig = 0
    End If

    calcStep = "Учет Pmax/Pmin (резерв)"
    Dim pMaxEff As Double, pMinEff As Double
    Dim reservePlus As Double, reserveMinus As Double
    pMaxEff = g.PMax
    If pMaxEff <= 0 Then pMaxEff = g.PNom
    pMinEff = g.PMin
    If pMinEff < 0 Then pMinEff = 0
    reservePlus = pMaxEff - p0
    reserveMinus = p0 - pMinEff
    If reservePlus < 0 Then reservePlus = 0
    If reserveMinus < 0 Then reserveMinus = 0

    ' Капинг: реально достижимое значение с учётом эксплуатационного диапазона
    preq = preqOrig
    If preq > reservePlus Then
        preq = reservePlus
        res.Limited = True
        res.LimitType = "Pmax"
    ElseIf preq < -reserveMinus Then
        preq = -reserveMinus
        res.Limited = True
        res.LimitType = "Pmin"
    Else
        res.Limited = False
        res.LimitType = ""
    End If

    calcStep = "Расчет фактической мощности"
    pfact = ptek - p0

    calcStep = "Количественный критерий"
    If dfr = 0 Then
        qpct = 100
        qpass = True
    ElseIf Abs(preq) < 0.000001 Then
        ' Pтреб_исх есть, но капинг обнулил: резерв в нужную сторону равен нулю.
        ' Ничего требовать нельзя - считаем статус ОК (участие ограничено диапазоном).
        qpct = 100
        qpass = True
    ElseIf SgnNZ(pfact) <> SgnNZ(preq) Then
        qpct = 0
        qpass = False
    Else
        qpct = 100# * SafeDiv(Abs(pfact), Abs(preq), 0)
        qpass = (qpct >= (100# - st.QuantTolPct))
    End If

    res.StartRow = startRow
    res.EndQuantRow = endQ
    res.EndQualRow = endQual
    res.StartTime = wsRaw.Cells(startRow, timeCol).Value
    If firstExceedRow > 0 Then
        res.FirstExceedTime = wsRaw.Cells(firstExceedRow, timeCol).Value
    Else
        res.FirstExceedTime = ""
    End If
    res.P0 = p0
    res.PTek = ptek
    res.Df = df
    res.Dfr = dfr
    res.PReq = preq
    res.PReqOrig = preqOrig
    res.PFact = pfact
    res.KdUsedQuant = kdQuant
    res.KdProfile = KdProfileText(g.EquipType)

    ' Статический Pтреб (Kд = 1, окончание переходного процесса), ограниченный резервом.
    ' Используется как цель качественных ступеней t5/t10 (п. 4.2.2, 4.2.7).
    calcStep = "Статический Pтреб"
    Dim preqStatic As Double
    If dfr <> 0 Then
        preqStatic = -100# / g.SPct * g.PNom / st.FNom * 1# * dfr
    Else
        preqStatic = 0
    End If
    If preqStatic > reservePlus Then preqStatic = reservePlus
    If preqStatic < -reserveMinus Then preqStatic = -reserveMinus
    res.PReqStatic = preqStatic
    res.PMaxEff = pMaxEff
    res.PMinEff = pMinEff
    res.ReservePlus = reservePlus
    res.ReserveMinus = reserveMinus
    res.QuantPct = qpct
    res.QuantPass = qpass
    res.Overshoot = (qpass And qpct > (100# + st.QuantTolPct))

    ' WARN при недостаточном резерве в нужную сторону (порог 5 %Pном)
    calcStep = "Проверка резерва 5 %Pном"
    Dim minReservePct As Double, minReserveMW As Double, needSign As Integer
    minReservePct = 5#
    minReserveMW = minReservePct / 100# * g.PNom
    needSign = SgnNZ(preqOrig)
    If needSign = 1 And reservePlus < minReserveMW Then
        AppendLog "WARN", g.Station & "/" & g.Generator, _
                  "Резерв '+' = " & Format(reservePlus, "0.0") & " МВт < " & _
                  Format(minReserveMW, "0.0") & " МВт (5 %Pном). Возможно ограничение по Pmax."
    ElseIf needSign = -1 And reserveMinus < minReserveMW Then
        AppendLog "WARN", g.Station & "/" & g.Generator, _
                  "Резерв '-' = " & Format(reserveMinus, "0.0") & " МВт < " & _
                  Format(minReserveMW, "0.0") & " МВт (5 %Pном). Возможно ограничение по Pmin."
    End If
    If res.Limited Then
        AppendLog "INFO", g.Station & "/" & g.Generator, _
                  "Pтреб ограничен " & res.LimitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

    calcStep = "Амплитуда события"
    ' Амплитуда берётся по исходному (неограниченному) Pтреб - она характеризует
    ' масштаб возмущения, а не способность генератора его отработать.
    If g.PNom > 0 Then
        res.AmplPctPnom = 100# * Abs(preqOrig) / g.PNom
    Else
        res.AmplPctPnom = 0
    End If
    res.AmplitudeTag = AmplitudeTag(res.AmplPctPnom, dfr <> 0)

    calcStep = "Реакция РЧВ через 15 с (п. 4.2.11)"
    Dim r15Status As String, r15Pass As Boolean, dp15 As Double, r15Note As String
    EvaluateReaction15 wsRaw, st, pCol, fCol, timeCol, startRow, firstExceedRow, p0, g.PNom, g.SPct, _
                       g.Fnch, g.EquipType, SgnNZ(preqOrig), reservePlus, reserveMinus, _
                       r15Status, r15Pass, dp15, r15Note
    res.React15Status = r15Status
    res.React15Pass = r15Pass
    res.DP15 = dp15
    res.React15Note = r15Note
    If r15Status = "Нарушение" Then
        AppendLog "WARN", g.Station & "/" & g.Generator, "п. 4.2.11: " & r15Note
    End If

    calcStep = "Дефицит первичной мощности по точкам окна (п. 4.2.14.1)"
    Dim defMaxPct As Double, defSec As Double
    EvaluateDeficit wsRaw, st, pCol, fCol, timeCol, startRow, endQ, p0, g.PNom, g.SPct, g.Fnch, _
                    g.EquipType, reservePlus, reserveMinus, defMaxPct, defSec
    res.DeficitMaxPct = defMaxPct
    res.DeficitSec = defSec

    calcStep = "Характер переходного процесса"
    EvaluateTransient wsRaw, st, g, res, pCol, timeCol

    calcStep = "Качественный критерий"
    EvaluateQualitative wsRaw, st, g, res, pCol, fCol, timeCol

    calcStep = "Итоговый вывод"
    res.Verdict = BuildGeneratorVerdict(g, res, st)

    AnalyzeOneGenerator = res
    Exit Function

EH:
    Err.Raise vbObjectError + 2199, , "AnalyzeOneGenerator (" & g.Generator & " / " & calcStep & "): " & Err.Description
End Function

' ----------------------------------------------------------
' п. 4.2.6: среднее значение мощности на интервале [t(row) - windowSec ; t(row)].
' windowSec <= 0 -> текущее значение в строке row.
' ----------------------------------------------------------
Private Function AvgPowerBackward(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
                                  ByVal row As Long, ByVal windowSec As Double) As Double
    Dim r1 As Long, r As Long, sumP As Double, cnt As Long
    If windowSec <= 0 Then
        AvgPowerBackward = NzD(wsRaw.Cells(row, pCol).Value, 0)
        Exit Function
    End If
    r1 = RowByTimeOffset(wsRaw, timeCol, row, -windowSec)
    If r1 < 2 Then r1 = 2
    If r1 > row Then r1 = row
    For r = r1 To row
        If IsNumeric(wsRaw.Cells(r, pCol).Value) And Len(Trim$(CStr(wsRaw.Cells(r, pCol).Value))) > 0 Then
            sumP = sumP + CDbl(wsRaw.Cells(r, pCol).Value)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then AvgPowerBackward = sumP / cnt Else AvgPowerBackward = NzD(wsRaw.Cells(row, pCol).Value, 0)
End Function

' ----------------------------------------------------------
' п. 4.2.11: наличие реакции регулятора частоты вращения турбины на
' первоначальное отклонение частоты - обратное изменению частоты изменение
' мощности через 15 с после выхода частоты за fнч. Расчёт по текущим значениям.
' Отсутствие реакции НА ЗАГРУЗКУ допустимо только при отсутствии запаса на
' открытие РК (резерв '+' по Pmax); реакция НА РАЗГРУЗКУ обязательна всегда.
' ----------------------------------------------------------
Private Sub EvaluateReaction15(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                               ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                               ByVal startRow As Long, ByVal firstExceedRow As Long, _
                               ByVal p0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
                               ByVal fnch As Double, ByVal equipType As String, ByVal needSign As Integer, _
                               ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                               ByRef status As String, ByRef pass As Boolean, _
                               ByRef dp15 As Double, ByRef note As String)
    Dim baseRow As Long, row15 As Long, dfr15 As Double, preq15 As Double, thrMW As Double
    dp15 = 0
    If needSign = 0 Then
        status = "Н/П"
        pass = True
        note = "Отклонение частоты в пределах fнч - оценка реакции РЧВ не требуется"
        Exit Sub
    End If
    baseRow = IIf(firstExceedRow > 0, firstExceedRow, startRow)
    row15 = RowByTimeOffset(wsRaw, timeCol, baseRow, 15#)
    dp15 = NzD(wsRaw.Cells(row15, pCol).Value, 0) - p0
    ' Требуемая мощность на 15-й секунде по текущей частоте (Kд(15 с))
    dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, fCol).Value, st.FNom) - st.FNom, fnch)
    preq15 = 0
    If dfr15 <> 0 And sPct > 0 Then
        preq15 = -100# / sPct * pNom / st.FNom * DynamicKdByTime(equipType, 15#) * dfr15
        If preq15 > reservePlus Then preq15 = reservePlus
        If preq15 < -reserveMinus Then preq15 = -reserveMinus
    End If
    thrMW = React15ThresholdMW(st, pNom, preq15)
    ClassifyReaction15 st, dp15, pNom, needSign, reservePlus, reserveMinus, thrMW, (dfr15 = 0), status, pass, note
End Sub

' Порог "наличия реакции": не более React15ThrPct %Pном и не более половины |Pтреб(15 с)|
' (чтобы при малом отклонении частоты не требовать больше, чем положено по Kд(15 с)).
Private Function React15ThresholdMW(ByRef st As TSettings, ByVal pNom As Double, ByVal preq15 As Double) As Double
    Dim thr As Double
    thr = st.React15ThrPct / 100# * pNom
    If thr < 0 Then thr = 0
    If Abs(preq15) > 0 And 0.5 * Abs(preq15) < thr Then thr = 0.5 * Abs(preq15)
    React15ThresholdMW = thr
End Function

' Классификация реакции по уже вычисленному dP15 (общая для генератора и суммы станции)
Private Sub ClassifyReaction15(ByRef st As TSettings, ByVal dp15 As Double, ByVal pNom As Double, _
                               ByVal needSign As Integer, ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                               ByVal thrMW As Double, ByVal freqBackInside As Boolean, _
                               ByRef status As String, ByRef pass As Boolean, ByRef note As String)
    Dim dirText As String
    If needSign = 0 Then
        status = "Н/П"
        pass = True
        note = "Отклонение частоты в пределах fнч - оценка реакции РЧВ не требуется"
        Exit Sub
    End If
    dirText = IIf(needSign = 1, "загрузку", "разгрузку")

    If needSign * dp15 >= thrMW And Abs(dp15) > 0 Then
        status = "ОК"
        pass = True
        note = "Реакция РЧВ на " & dirText & " через 15 с есть: dP15 = " & Format(dp15, "0.000") & " МВт"
    ElseIf freqBackInside Then
        status = "Н/П"
        pass = True
        note = "К 15-й секунде частота вернулась в зону нечувствительности - реакция РЧВ не оценивается; dP15 = " & Format(dp15, "0.000") & " МВт"
    ElseIf needSign = 1 And reservePlus < thrMW Then
        status = "Н/П"
        pass = True
        note = "Нет запаса на открытие РК (резерв '+' = " & Format(reservePlus, "0.0") & _
               " МВт по Pmax) - отсутствие реакции на загрузку допустимо (п. 4.2.11); dP15 = " & Format(dp15, "0.000") & " МВт"
    Else
        status = "Нарушение"
        pass = False
        note = "Отсутствует реакция РЧВ на " & dirText & " через 15 с после выхода за fнч: dP15 = " & _
               Format(dp15, "0.000") & " МВт при пороге " & Format(needSign * thrMW, "0.000") & " МВт"
        If needSign = -1 And reserveMinus < thrMW Then
            note = note & " (резерв '-' = " & Format(reserveMinus, "0.0") & " МВт мал - проверьте Pmin; по п. 4.2.11 реакция на разгрузку обязательна)"
        End If
    End If
End Sub

' ----------------------------------------------------------
' п. 4.2.14.1 / Приложение (метод допустимых границ): дефицит первичной
' мощности по каждой точке окна, где частота вне fнч. Граница = Pтреб(t) по
' Kд(t), ограниченный резервом, минус допуск 1 %Pном (п. 4.2.7/4.2.8).
' Первые 15 с после старта (переходный участок) исключаются - они оцениваются
' по п. 4.2.11 (наличие реакции РЧВ).
' Возвращает макс. дефицит в %Pном и суммарную длительность дефицита, с.
' ----------------------------------------------------------
Private Sub EvaluateDeficit(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                            ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                            ByVal startRow As Long, ByVal endRow As Long, _
                            ByVal p0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
                            ByVal fnch As Double, ByVal equipType As String, _
                            ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                            ByRef defMaxPct As Double, ByRef defSec As Double)
    Dim r As Long, dFrT As Double, preqT As Double, allowed As Double, have As Double, deficit As Double
    Dim tSec As Double, dt As Double, tol As Double
    defMaxPct = 0
    defSec = 0
    If pNom <= 0 Or sPct <= 0 Then Exit Sub
    tol = 0.01 * pNom
    For r = startRow To endRow
        dFrT = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, fnch)
        If dFrT = 0 Then GoTo NextR
        tSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        ' первые 15 с - переходный участок, оценивается отдельно по п. 4.2.11 (реакция РЧВ)
        If tSec < 15 Then GoTo NextR
        preqT = -100# / sPct * pNom / st.FNom * DynamicKdByTime(equipType, tSec) * dFrT
        If preqT > reservePlus Then preqT = reservePlus
        If preqT < -reserveMinus Then preqT = -reserveMinus
        allowed = Abs(preqT) - tol
        If allowed <= 0 Then GoTo NextR
        have = SgnNZ(preqT) * (NzD(wsRaw.Cells(r, pCol).Value, 0) - p0)
        deficit = allowed - have
        If deficit > 0 Then
            If 100# * deficit / pNom > defMaxPct Then defMaxPct = 100# * deficit / pNom
            If r < endRow Then
                dt = SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(r + 1, timeCol).Value)
            ElseIf r > startRow Then
                dt = SecBetween(wsRaw.Cells(r - 1, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
            Else
                dt = 1
            End If
            If dt <= 0 Or dt > 60 Then dt = 1
            defSec = defSec + dt
        End If
NextR:
    Next r
End Sub

' ----------------------------------------------------------
' Итоговый вывод по генератору (текст для Summary / листа графика)
' ----------------------------------------------------------
Private Function BuildGeneratorVerdict(ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String
    bad = ""
    remark = ""

    If res.Limited Then
        remark = remark & "Pтреб ограничен " & res.LimitType & " (оценка по располагаемому резерву); "
    ElseIf Not res.QuantPass Then
        bad = bad & "колич. критерий (" & Format(res.QuantPct, "0") & " % от Pтреб); "
    End If
    If res.React15Status = "Нарушение" Then
        bad = bad & "нет реакции РЧВ через 15 с (п. 4.2.11); "
    ElseIf res.React15Status = "Н/П" And Len(res.React15Note) > 0 And res.Dfr <> 0 Then
        remark = remark & "реакция 15 с не оценивается - нет запаса по Pmax; "
    End If
    If res.DeficitSec > 0 Then
        remark = remark & "дефицит до " & Format(res.DeficitMaxPct, "0.0") & " %Pном в течение " & _
                 Format(res.DeficitSec, "0") & " с (п. 4.2.14.1); "
    End If
    If g.QualEnabled And res.Dfr <> 0 And Not res.QualPass Then
        remark = remark & "кач. подпункты: " & res.QualFailedList & "; "
    End If
    If res.TransientType = "Колебательный" Then
        remark = remark & "колебательный процесс (" & res.NumExtrema & " экстремумов, п. 4.2.14.4); "
    End If

    If Len(bad) > 0 Then
        txt = "НАРУШЕНИЕ: " & bad
    Else
        txt = "ОК: количественно и по реакции 15 с соответствует; "
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    If Right$(txt, 2) = "; " Then txt = Left$(txt, Len(txt) - 2)
    BuildGeneratorVerdict = txt
End Function

' ==========================================================
' Качественная оценка
' ==========================================================

Private Sub EvaluateQualitative(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long)
    On Error GoTo EH

    Dim signReq As Integer, t5 As Double, t10 As Double
    Dim r As Long, dP As Double, target5 As Double, target10 As Double
    Dim hit5 As Boolean, hit10 As Boolean, row5 As Long, row10 As Long
    Dim steadyMean As Double, steadyTolMW As Double
    Dim reason As String, failed As String, qStep As String

    qStep = "Проверка включения"
    If Not g.QualEnabled Then
        res.QualPass = True
        res.QualT5Pass = True
        res.QualT10Pass = True
        res.QualSteadyPass = True
        res.T5FactSec = -1
        res.T10FactSec = -1
        res.QualFailedList = ""
        res.QualReason = "Качественная проверка отключена"
        res.PsteadyAvg = 0
        Exit Sub
    End If

    qStep = "Направление требуемого отклика"
    signReq = SgnNZ(res.PReq)
    If signReq = 0 Then
        res.QualPass = True
        res.QualT5Pass = True
        res.QualT10Pass = True
        res.QualSteadyPass = True
        res.T5FactSec = -1
        res.T10FactSec = -1
        res.QualFailedList = ""
        res.QualReason = "Вне зоны отклонения"
        res.PsteadyAvg = 0
        Exit Sub
    End If

    qStep = "Целевые ступени dP5/dP10"
    ' Цели масштабируются к статическому |Pтреб| (Kд = 1, окончание переходного
    ' процесса, п. 4.2.2/4.2.7), ограниченному резервом; dP5/dP10 задают пропорции
    ' (5%/10% = 0.5 = половина Pтреб к моменту t5 - это соответствует Kд(15 с) = 0.5
    ' для ТЭС по Приложению к МУ; полный Pтреб к t10).
    ' Для контрольных испытаний (ступень 10 % Pном) это даёт тот же результат,
    ' т. к. Pтреб = Pном * dp10% / 100.
    Dim pReqAbs As Double, ratio5 As Double, tolMW As Double
    pReqAbs = Abs(res.PReqStatic)
    If pReqAbs < 0.000001 Then pReqAbs = Abs(res.PReq)
    If g.Dp10Pct > 0 Then
        ratio5 = g.Dp5Pct / g.Dp10Pct
    Else
        ratio5 = 0.5
    End If
    tolMW = 0.01 * g.PNom
    target5 = signReq * pReqAbs * ratio5
    target10 = signReq * pReqAbs

    row5 = 0
    row10 = 0
    hit5 = False
    hit10 = False

    qStep = "Поиск t5/t10"
    ' Ступень считается достигнутой с учётом допуска 1 %Pном (п. 4.2.7).
    For r = res.StartRow To res.EndQualRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        If Not hit5 Then
            If signReq * dP >= signReq * target5 - tolMW Then
                hit5 = True
                row5 = r
            End If
        End If
        If Not hit10 Then
            If signReq * dP >= signReq * target10 - tolMW Then
                hit10 = True
                row10 = r
            End If
        End If
    Next r

    ' Если частота к t10 уже вернулась (Pтреб(t) снизился), полный статический Pтреб
    ' объективно не нужен: ступень t10 засчитывается, если в конце качественного окна
    ' нет дефицита относительно текущего Pтреб(t) по Kд(t) (с допуском 1 %Pном).
    If Not hit10 Then
        Dim dFrEnd As Double, preqEnd As Double, haveEnd As Double
        dFrEnd = DeadbandDeviation(NzD(wsRaw.Cells(res.EndQualRow, fCol).Value, st.FNom) - st.FNom, g.Fnch)
        preqEnd = -100# / g.SPct * g.PNom / st.FNom * _
                  DynamicKdByTime(g.EquipType, SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, _
                                                          wsRaw.Cells(res.EndQualRow, timeCol).Value)) * dFrEnd
        If preqEnd > res.ReservePlus Then preqEnd = res.ReservePlus
        If preqEnd < -res.ReserveMinus Then preqEnd = -res.ReserveMinus
        haveEnd = NzD(wsRaw.Cells(res.EndQualRow, pCol).Value, 0) - res.P0
        If SgnNZ(preqEnd) = 0 Or SgnNZ(preqEnd) <> signReq Then
            ' частота вернулась в зону нечувствительности (или сменила знак) -
            ' требование к t10 снято, если генератор не остался с дефицитом
            hit10 = True
            row10 = res.EndQualRow
        ElseIf signReq * haveEnd >= Abs(preqEnd) - tolMW Then
            hit10 = True
            row10 = res.EndQualRow
        End If
    End If

    qStep = "Расчет фактических времен"
    If hit5 Then
        t5 = SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row5, timeCol).Value)
    Else
        t5 = -1
    End If
    If hit10 Then
        t10 = SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row10, timeCol).Value)
    Else
        t10 = -1
    End If
    res.T5FactSec = t5
    res.T10FactSec = t10

    qStep = "Среднее установившееся"
    steadyMean = ComputeSteadyMean(wsRaw, pCol, timeCol, res.EndQualRow, st.SteadyWindowSec)
    res.PsteadyAvg = steadyMean
    steadyTolMW = g.SteadyTolPct / 100# * g.PNom

    ' В реальных событиях частота может восстановиться к концу качественного окна.
    ' Поэтому цель установившегося = среднее Pтреб(t) в том же хвостовом окне.
    ' Если частота вернулась, Pтреб_ср близок к нулю и генератор тоже должен вернуться к P0.
    res.PReqSteady = ComputeSteadyPReqMean(wsRaw, fCol, timeCol, res.EndQualRow, _
                                           st.SteadyWindowSec, st.FNom, g.SPct, g.PNom, g.Fnch, _
                                           g.EquipType, wsRaw.Cells(res.StartRow, timeCol).Value)

    ' допуск 0.5 с на дискретность выборки
    res.QualT5Pass = (hit5 And t5 <= g.T5Sec + 0.5)
    res.QualT10Pass = (hit10 And t10 <= g.T10Sec + 0.5)
    If g.CheckSteady Then
        res.QualSteadyPass = (Abs((steadyMean - res.P0) - res.PReqSteady) <= steadyTolMW)
    Else
        res.QualSteadyPass = True
    End If

    reason = ""
    failed = ""
    If Not res.QualT5Pass Then
        reason = reason & "Не достигнута 1-я ступень (" & Format(ratio5 * 100#, "0") & " % Pтреб) к t5=" & g.T5Sec & "c; "
        failed = failed & "t5; "
    End If
    If Not res.QualT10Pass Then
        reason = reason & "Не достигнут Pтреб к t10=" & g.T10Sec & "c; "
        failed = failed & "t10; "
    End If
    If g.CheckSteady And (Not res.QualSteadyPass) Then
        reason = reason & "Установившееся отклонение от Pтреб_ср (" & Format(res.PReqSteady, "0.000") _
                        & ") выходит за допуск ±" & Format(steadyTolMW, "0.000") & " МВт; "
        failed = failed & "уст; "
    End If
    If Not g.CheckSteady Then reason = reason & "Контроль установившегося отключен; "

    res.QualPass = (res.QualT5Pass And res.QualT10Pass And res.QualSteadyPass)
    If Len(failed) > 0 Then
        If Right$(failed, 2) = "; " Then failed = Left$(failed, Len(failed) - 2)
        res.QualFailedList = failed
    Else
        res.QualFailedList = ""
    End If

    If res.QualPass Then
        res.QualReason = "Качественно: ОК; t5=" & IIf(t5 >= 0, Format(t5, "0.0"), "н/д") & "с; t10=" & IIf(t10 >= 0, Format(t10, "0.0"), "н/д") & "с"
    Else
        res.QualReason = reason & "t5=" & IIf(t5 >= 0, Format(t5, "0.0"), "н/д") & "с; t10=" & IIf(t10 >= 0, Format(t10, "0.0"), "н/д") & "с"
    End If
    Exit Sub

EH:
    Err.Raise vbObjectError + 2299, , "EvaluateQualitative (" & g.Generator & " / " & qStep & "): " & Err.Description
End Sub

Private Function ComputeSteadyMean(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
                                    ByVal endRow As Long, ByVal windowSec As Double) As Double
    Dim startRow As Long, r As Long
    Dim sumP As Double, cnt As Long
    If windowSec <= 0 Then windowSec = 30
    startRow = RowByTimeOffset(wsRaw, timeCol, endRow, -windowSec)
    If startRow < 2 Then startRow = 2
    If startRow >= endRow Then startRow = endRow
    For r = startRow To endRow
        If IsNumeric(wsRaw.Cells(r, pCol).Value) Then
            sumP = sumP + CDbl(wsRaw.Cells(r, pCol).Value)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then ComputeSteadyMean = sumP / cnt Else ComputeSteadyMean = 0
End Function

Private Function ComputeSteadyPReqMean(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal timeCol As Long, _
                                        ByVal endRow As Long, ByVal windowSec As Double, _
                                        ByVal fNom As Double, ByVal sPct As Double, ByVal pNom As Double, _
                                        ByVal fnch As Double, _
                                        ByVal equipType As String, ByVal tStart As Variant) As Double
    ' Усреднённый Pтреб(t) за окно [endRow-windowSec ; endRow].
    ' Pтреб(t) = -100/S * Pном/fном * Kd * dFr(t), где dFr - отклонение за fнч.
    Dim startRow As Long, r As Long
    Dim sumP As Double, cnt As Long
    Dim dFr As Double, fv As Double, kdEff As Double, tSec As Double
    If windowSec <= 0 Then windowSec = 30
    If sPct <= 0 Then Exit Function
    startRow = RowByTimeOffset(wsRaw, timeCol, endRow, -windowSec)
    If startRow < 2 Then startRow = 2
    If startRow >= endRow Then startRow = endRow
    For r = startRow To endRow
        If IsNumeric(wsRaw.Cells(r, freqCol).Value) Then
            fv = CDbl(wsRaw.Cells(r, freqCol).Value) - fNom
            dFr = DeadbandDeviation(fv, fnch)
            tSec = SecBetween(tStart, wsRaw.Cells(r, timeCol).Value)
            kdEff = DynamicKdByTime(equipType, tSec)
            sumP = sumP + (-100# / sPct * pNom / fNom * kdEff * dFr)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then ComputeSteadyPReqMean = sumP / cnt Else ComputeSteadyPReqMean = 0
End Function

' ==========================================================
' Характер переходного процесса: Монотонный / Апериодический / Колебательный
' ==========================================================

Private Sub EvaluateTransient(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                              ByVal pCol As Long, ByVal timeCol As Long)
    Dim r As Long, n As Long, i As Long
    Dim arr() As Double
    Dim extrCount As Long
    Dim prevDir As Integer, curDir As Integer
    Dim dP As Double, prev As Double
    Dim amps() As Double, ampCount As Long
    Dim lastExtrVal As Double
    Dim hasLast As Boolean
    Dim maxAbs As Double, noise As Double
    Dim smoothed() As Double

    If res.EndQualRow <= res.StartRow Then
        res.TransientType = "н/д"
        res.NumExtrema = 0
        Exit Sub
    End If

    n = res.EndQualRow - res.StartRow + 1
    ReDim arr(1 To n)
    i = 0
    For r = res.StartRow To res.EndQualRow
        i = i + 1
        arr(i) = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        If Abs(arr(i)) > maxAbs Then maxAbs = Abs(arr(i))
    Next r

    If n < 5 Or maxAbs < 0.0001 Then
        res.TransientType = "н/д"
        res.NumExtrema = 0
        Exit Sub
    End If

    ' Сглаживание окном 3 точки (убираем шум дискретизации)
    ReDim smoothed(1 To n)
    smoothed(1) = arr(1)
    smoothed(n) = arr(n)
    For i = 2 To n - 1
        smoothed(i) = (arr(i - 1) + arr(i) + arr(i + 1)) / 3#
    Next i

    noise = maxAbs * 0.05   ' Порог "значимого" экстремума: 5 % от амплитуды
    extrCount = 0
    ReDim amps(1 To n)
    ampCount = 0
    prevDir = 0
    lastExtrVal = smoothed(1)
    hasLast = False

    For i = 2 To n - 1
        dP = smoothed(i) - smoothed(i - 1)
        If dP > noise Then
            curDir = 1
        ElseIf dP < -noise Then
            curDir = -1
        Else
            curDir = 0
        End If

        If prevDir <> 0 And curDir <> 0 And prevDir <> curDir Then
            ' Локальный экстремум в точке i-1
            If Abs(smoothed(i - 1) - lastExtrVal) >= noise Or Not hasLast Then
                extrCount = extrCount + 1
                ampCount = ampCount + 1
                amps(ampCount) = Abs(smoothed(i - 1) - lastExtrVal)
                lastExtrVal = smoothed(i - 1)
                hasLast = True
            End If
        End If
        If curDir <> 0 Then prevDir = curDir
    Next i

    res.NumExtrema = extrCount

    Dim decaying As Boolean
    decaying = True
    For i = 2 To ampCount
        If amps(i) > amps(i - 1) * 1.1 Then
            decaying = False
            Exit For
        End If
    Next i

    If extrCount <= 1 Then
        res.TransientType = "Монотонный"
    ElseIf extrCount <= 3 And decaying Then
        res.TransientType = "Апериодический"
    Else
        res.TransientType = "Колебательный"
    End If
End Sub

' ==========================================================
' Запись листа данных генератора
' ==========================================================

Private Sub WriteGeneratorSheet(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim ws As Worksheet
    Dim shName As String
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim endRow As Long, r As Long, outR As Long
    Dim dP As Double, dFr As Double
    Dim displayStartRow As Long, chartEndRow As Long
    Dim target5Val As Double, target10Val As Double
    Dim targetPreq As Double
    Dim tolPreq As Double
    Dim signReq As Integer

    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderCol(wsRaw, g.PowerHeader)
    fCol = FindHeaderCol(wsRaw, g.FreqHeader)

    shName = GeneratorSheetName(g)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    ' Шапка (колонки A:B и D:E)
    Dim quantStatusStr As String
    If res.Limited Then
        quantStatusStr = "Ограничен " & res.LimitType
    ElseIf res.QuantPass Then
        quantStatusStr = "ОК"
    Else
        quantStatusStr = "Нарушение"
    End If

    ws.Range("A1:B1").Value = Array("Станция", g.Station)
    ws.Range("A2:B2").Value = Array("Генератор", g.Generator)
    ws.Range("A3:B3").Value = Array("Тип", g.EquipType)
    ws.Range("A4:B4").Value = Array("Старт (расч.)", res.StartTime)
    ws.Range("A5:B5").Value = Array("Выход за fнч", res.FirstExceedTime)
    ws.Range("A6:B6").Value = Array("Колич. статус", quantStatusStr)
    ws.Range("A7:B7").Value = Array("Кач. статус", IIf(res.QualPass, "ОК", "Нарушение"))
    ws.Range("A8:B8").Value = Array("Характер", res.TransientType)
    ws.Range("A9:B9").Value = Array("Амплитуда, %Pном", Round(res.AmplPctPnom, 2))
    ws.Cells(1, 4).Resize(1, 2).Value = Array("P0, МВт", res.P0)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pтек, МВт", res.PTek)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Pуст_сред, МВт", res.PsteadyAvg)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("dF, Гц", res.Df)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("dFr, Гц", res.Dfr)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Pтреб, МВт", res.PReq)
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Pфакт, МВт", res.PFact)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Колич. %", res.QuantPct)
    ws.Cells(9, 4).Resize(1, 2).Value = Array("Экстремумов", res.NumExtrema)

    ws.Cells(1, 7).Resize(1, 2).Value = Array("Pmax, МВт", res.PMaxEff)
    ws.Cells(2, 7).Resize(1, 2).Value = Array("Pmin, МВт", res.PMinEff)
    ws.Cells(3, 7).Resize(1, 2).Value = Array("Резерв '+', МВт", res.ReservePlus)
    ws.Cells(4, 7).Resize(1, 2).Value = Array("Резерв '-', МВт", res.ReserveMinus)
    ws.Cells(5, 7).Resize(1, 2).Value = Array("Pтреб исх., МВт", res.PReqOrig)
    ws.Cells(6, 7).Resize(1, 2).Value = Array("Ограничение", IIf(res.Limited, "Да (" & res.LimitType & ")", "нет"))
    ws.Cells(7, 7).Resize(1, 2).Value = Array("Kд (колич.), факт", res.KdUsedQuant)
    ws.Cells(8, 7).Resize(1, 2).Value = Array("Профиль Kд(t)", res.KdProfile)
    ws.Cells(9, 7).Resize(1, 2).Value = Array("Реакция 15 с (п. 4.2.11)", res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт)")
    ws.Cells(10, 7).Resize(1, 2).Value = Array("Дефицит (п. 4.2.14.1)", _
        IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.00") & " %Pном, " & Format(res.DeficitSec, "0") & " с", "нет"))

    ws.Range("A11:V11").Value = Array( _
        "Время", "Частота, Гц", "P, МВт", "dPфакт, МВт", "Pтреб_накоп, МВт", "dFr, Гц", _
        "", "Уровень +допуск", "Уровень -допуск", _
        "Маркер t5", "Маркер t10", "Маркер выхода за fнч", _
        "dPmax", "dPmin", _
        "Pтреб_абс, МВт", "Уровень +допуск_абс", "Уровень -допуск_абс", _
        "Pmax, МВт", "Pmin, МВт", _
        "Маркер t5_абс", "Маркер t10_абс", "Маркер fнч_абс" _
    )

    If st.PreBufferSec > 0 Then
        displayStartRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, -st.PreBufferSec)
        If displayStartRow < 2 Then displayStartRow = 2
        If displayStartRow > res.StartRow Then displayStartRow = res.StartRow
    Else
        displayStartRow = res.StartRow
    End If
    endRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, MaxD(st.QuantIntervalSec, g.T10Sec))
    outR = 12

    ' Уровни для маркеров. Целевой уровень = Pтреб (масштабируется к событию).
    signReq = SgnNZ(res.PReq)
    targetPreq = res.PReq
    tolPreq = g.PNom * 0.01
    ' Для вертикальных маркеров подбираем достаточный размах, чтобы линия была видна.
    ' Берём не меньше размаха Pmax/Pmin, чтобы маркер пересекал обе границы.
    Dim markerSpan As Double, markerSpanAbs As Double
    Dim dPmaxRel As Double, dPminRel As Double
    Dim markerHiAbs As Double, markerLoAbs As Double
    dPmaxRel = res.PMaxEff - res.P0
    dPminRel = res.PMinEff - res.P0
    markerSpan = MaxD(Abs(res.PReq) * 1.5, g.PNom * g.Dp10Pct / 100#)
    markerSpan = MaxD(markerSpan, MaxD(Abs(dPmaxRel), Abs(dPminRel)))
    If markerSpan <= 0 Then markerSpan = g.PNom * 0.1
    markerSpanAbs = markerSpan
    markerHiAbs = res.P0 + markerSpanAbs
    markerLoAbs = res.P0 - markerSpanAbs
    target10Val = targetPreq

    Dim preqStep As Double
    For r = displayStartRow To endRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        dFr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, g.Fnch)
        preqStep = -100# / g.SPct * g.PNom / st.FNom * DynamicKdByTime(g.EquipType, _
                   SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)) * dFr

        ws.Cells(outR, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(outR, 2).Value = wsRaw.Cells(r, fCol).Value
        ws.Cells(outR, 3).Value = wsRaw.Cells(r, pCol).Value
        ws.Cells(outR, 4).Value = dP
        ws.Cells(outR, 5).Value = preqStep
        ws.Cells(outR, 6).Value = dFr
        ws.Cells(outR, 8).Value = targetPreq + tolPreq
        ws.Cells(outR, 9).Value = targetPreq - tolPreq
        ws.Cells(outR, 13).Value = dPmaxRel
        ws.Cells(outR, 14).Value = dPminRel
        ws.Cells(outR, 15).Value = res.P0 + preqStep
        ws.Cells(outR, 16).Value = res.P0 + targetPreq + tolPreq
        ws.Cells(outR, 17).Value = res.P0 + targetPreq - tolPreq
        ws.Cells(outR, 18).Value = res.PMaxEff
        ws.Cells(outR, 19).Value = res.PMinEff
        outR = outR + 1
    Next r

    ' Маркеры t5 / t10 / выхода за fнч - две точки на линии (нижний и верхний уровень)
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 10, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 11, markerSpan, -markerSpan
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 12, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 20, markerHiAbs, markerLoAbs
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 21, markerHiAbs, markerLoAbs
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 22, markerHiAbs, markerLoAbs

    chartEndRow = outR - 1

    ws.Columns("A:V").AutoFit
    ApplyGeneratorSheetFormats ws
    ' Вывод записываем после AutoFit, чтобы длинный текст не растягивал колонку B
    ws.Cells(10, 1).Value = "Вывод"
    ws.Cells(10, 2).Value = res.Verdict
    ws.Cells(10, 1).Font.Bold = True
End Sub

Private Sub FillMarkerColumn(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
                             ByVal wsRaw As Worksheet, ByVal timeCol As Long, _
                             ByVal startRow As Long, ByVal offsetSec As Double, _
                             ByVal outCol As Long, ByVal vHi As Double, ByVal vLo As Double)
    Dim mkRow As Long
    mkRow = RowByTimeOffset(wsRaw, timeCol, startRow, offsetSec)
    If mkRow < 2 Then Exit Sub
    FillMarkerColumnAtTime ws, rr1, rr2, wsRaw, timeCol, wsRaw.Cells(mkRow, timeCol).Value, outCol, vHi, vLo
End Sub

Private Sub FillMarkerColumnAtTime(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
                                   ByVal wsRaw As Worksheet, ByVal timeCol As Long, _
                                   ByVal tMark As Variant, ByVal outCol As Long, _
                                   ByVal vHi As Double, ByVal vLo As Double)
    Dim r As Long, rowIdx As Long, tM As Double
    Dim tPrev As Double, tCur As Double, match As Boolean
    If Not IsDate(tMark) Then Exit Sub
    tM = CDbl(CDate(tMark))
    rowIdx = 12
    tPrev = 0
    match = False
    For r = rr1 To rr2
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            tCur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            If (tPrev <= tM And tCur >= tM) Or Abs(tCur - tM) < 0.5 / 86400# Then
                ws.Cells(rowIdx, outCol).Value = vHi
                ws.Cells(rowIdx + 1, outCol).Value = vLo
                match = True
                Exit For
            End If
            tPrev = tCur
        End If
        rowIdx = rowIdx + 1
    Next r
    If Not match Then Exit Sub
End Sub

Private Sub ApplyGeneratorSheetFormats(ByVal ws As Worksheet)
    ws.Range("A1:A9").NumberFormat = "@"
    ws.Range("B1:B3").NumberFormat = "@"
    ws.Range("B4:B5").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B6:B8").NumberFormat = "@"
    ws.Range("B9").NumberFormat = "0.00"
    ws.Range("D1:D9").NumberFormat = "@"
    ws.Range("E1:E9").NumberFormat = "0.000"
    ws.Range("A10:B10").NumberFormat = "@"
    ws.Range("G1:G10").NumberFormat = "@"
    ws.Range("H1:H5").NumberFormat = "0.000"
    ws.Range("H6").NumberFormat = "@"
    ws.Range("H7").NumberFormat = "0.000"
    ws.Range("H8:H10").NumberFormat = "@"
    ws.Range("A11:V11").NumberFormat = "@"
    ws.Range("A12:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B12:V100000").NumberFormat = "0.000"
End Sub

' ==========================================================
' Лист с графиком
' ==========================================================

Private Sub WriteGeneratorChartSheet(ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataSheet As String, chartSheet As String
    Dim lastDataRow As Long
    Dim chartObj As ChartObject
    Dim startDataRow As Long, endChartDataRow As Long
    Dim ySpan As Double
    Dim yMin As Double, yMax As Double, yPad As Double
    Dim farPos As Double, farNeg As Double, farAbs As Double

    dataSheet = GeneratorSheetName(g)
    chartSheet = GeneratorChartSheetName(g)
    Set wsData = ThisWorkbook.Worksheets(dataSheet)
    Set wsChart = EnsureSheet(chartSheet)

    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    ' Заголовок
    wsChart.Range("A1").Value = "График ОПРЧ: " & g.Station & " / " & g.Generator
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    ' Блок "Вывод"
    WriteChartVerdictBlock wsChart, g, res, st

    lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 12 Then Exit Sub

    startDataRow = 12
    endChartDataRow = FindChartEndRow(wsData, 12, lastDataRow, res.StartTime, st.ChartIntervalSec)
    If endChartDataRow < startDataRow Then endChartDataRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 170, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ: " & g.Station & " / " & g.Generator
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    ' Основные ряды
    If ChartOpt("REL_PFACT", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 4, "Pфакт, МВт", False
    If ChartOpt("REL_PREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 5, "Pтреб, МВт", False
    If ChartOpt("REL_TOL_P", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 8, "+Допуск уст.", False
    If ChartOpt("REL_TOL_M", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 9, "-Допуск уст.", False
    If ChartOpt("REL_LIMITS", True) Then
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 13, "Pmax (dPmax)"
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 14, "Pmin (dPmin)"
    End If
    If ChartOpt("REL_FREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True

    ' Вертикальные маркеры t5 / t10 / выхода за fнч - через отдельные "точечные" ряды
    If ChartOpt("REL_MARKERS", True) Then
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 10, "t5"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 11, "t10"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 12, "Выход за fнч"
    End If

    On Error Resume Next
    GetRangeMinMaxByCols wsData, startDataRow, endChartDataRow, Array(4, 5, 8, 9, 13, 14), yMin, yMax
    farPos = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(4, 5), True)
    farNeg = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(4, 5), False)
    farAbs = MaxD(farPos, farNeg)
    If farAbs > 0 Then
        If farPos >= farNeg Then
            yMax = MaxD(yMax, farPos * 1.05)
        Else
            yMin = -MaxD(Abs(yMin), farNeg * 1.05)
        End If
    End If
    yPad = MaxD((yMax - yMin) * 0.1, g.PNom * 0.01)
    If yPad <= 0 Then yPad = 0.5
    chartObj.Chart.Axes(xlValue, xlPrimary).MinimumScale = yMin - yPad
    chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = yMax + yPad
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(res.Df) * 1.2, 2# * g.Fnch)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0

    ' Второй график: абсолютные единицы (P, МВт и f, Гц)
    Set chartObj = wsChart.ChartObjects.Add(10, 550, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ (абсолютные): " & g.Station & " / " & g.Generator
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    If ChartOpt("ABS_PFACT", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 3, "Pфакт, МВт", False
    If ChartOpt("ABS_PREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 15, "Pтреб, МВт (абс)", False
    If ChartOpt("ABS_TOL_P", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 16, "+Допуск уст., МВт (абс)", False
    If ChartOpt("ABS_TOL_M", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 17, "-Допуск уст., МВт (абс)", False
    If ChartOpt("ABS_LIMITS", True) Then
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 18, "Pmax, МВт"
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 19, "Pmin, МВт"
    End If
    If ChartOpt("ABS_FREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True
    If ChartOpt("ABS_MARKERS", False) Then
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 20, "t5 abs"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 21, "t10 abs"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 22, "fнч abs"
    End If

    On Error Resume Next
    GetRangeMinMaxByCols wsData, startDataRow, endChartDataRow, Array(3, 15, 16, 17, 18, 19), yMin, yMax
    farPos = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(3, 15), True)
    farNeg = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(3, 15), False)
    farAbs = MaxD(farPos, farNeg)
    If farAbs > 0 Then
        If farPos >= farNeg Then
            yMax = MaxD(yMax, farPos * 1.05)
        Else
            yMin = -MaxD(Abs(yMin), farNeg * 1.05)
        End If
    End If
    yPad = MaxD((yMax - yMin) * 0.08, g.PNom * 0.005)
    If yPad <= 0 Then yPad = 1
    chartObj.Chart.Axes(xlValue, xlPrimary).MinimumScale = yMin - yPad
    chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = yMax + yPad
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(res.Df) * 1.2, 2# * g.Fnch)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0
End Sub

Private Sub GetRangeMinMaxByCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal cols As Variant, ByRef outMin As Double, ByRef outMax As Double)
    Dim i As Long, r As Long, c As Long, v As Variant
    Dim inited As Boolean
    inited = False
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        For r = r1 To r2
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                If Not inited Then
                    outMin = CDbl(v)
                    outMax = CDbl(v)
                    inited = True
                Else
                    If CDbl(v) < outMin Then outMin = CDbl(v)
                    If CDbl(v) > outMax Then outMax = CDbl(v)
                End If
            End If
        Next r
    Next i
    If Not inited Then
        outMin = 0
        outMax = 1
    End If
    If outMax < outMin Then
        outMax = outMin + 1
    ElseIf Abs(outMax - outMin) < 0.000001 Then
        outMax = outMax + 1
        outMin = outMin - 1
    End If
End Sub

Private Function MaxAbsInCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                              ByVal cols As Variant, ByVal positiveSide As Boolean) As Double
    Dim i As Long, r As Long, c As Long
    Dim v As Variant, vv As Double
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        For r = r1 To r2
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                vv = CDbl(v)
                If positiveSide Then
                    If vv > MaxAbsInCols Then MaxAbsInCols = vv
                Else
                    If -vv > MaxAbsInCols Then MaxAbsInCols = -vv
                End If
            End If
        Next r
    Next i
End Function

Private Sub WriteChartVerdictBlock(ByVal wsChart As Worksheet, ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings)
    Dim quantCell As String
    If res.Limited Then
        quantCell = "Ограничен " & res.LimitType & " (" & Format(res.QuantPct, "0") & " % от доступного)"
    ElseIf res.QuantPass Then
        quantCell = "ОК (" & Format(res.QuantPct, "0") & " %)"
    Else
        quantCell = "Нарушение (" & Format(res.QuantPct, "0") & " %)"
    End If

    wsChart.Range("A3").Value = "Количественный:"
    wsChart.Range("B3").Value = quantCell
    wsChart.Range("D3").Value = "Амплитуда:"
    wsChart.Range("E3").Value = Format(res.AmplPctPnom, "0.0") & " %Pном" & IIf(Len(res.AmplitudeTag) > 0, " / " & res.AmplitudeTag, "")

    wsChart.Range("A4").Value = "Качественный:"
    wsChart.Range("B4").Value = IIf(res.QualPass, "ОК", "Нарушение")
    wsChart.Range("D4").Value = "t5 / t10:"
    wsChart.Range("E4").Value = FormatSecOrNA(res.T5FactSec) & " / " & FormatSecOrNA(res.T10FactSec) & " с"

    wsChart.Range("A5").Value = "Характер:"
    wsChart.Range("B5").Value = res.TransientType & " (экстремумов " & res.NumExtrema & ")"
    wsChart.Range("D5").Value = "Перерегулирование:"
    wsChart.Range("E5").Value = IIf(res.Overshoot, "Да", "нет")

    wsChart.Range("A6").Value = "Уст_сред, МВт:"
    wsChart.Range("B6").Value = Format(res.PsteadyAvg, "0.000") & " (цель " & Format(res.P0 + res.PReqSteady, "0.000") & ")"
    wsChart.Range("D6").Value = "Проваленные подп.:"
    wsChart.Range("E6").Value = IIf(Len(res.QualFailedList) > 0, res.QualFailedList, "-")

    wsChart.Range("A7").Value = "Pmax / Pmin, МВт:"
    wsChart.Range("B7").Value = Format(res.PMaxEff, "0.0") & " / " & Format(res.PMinEff, "0.0") & _
                                " (резерв +" & Format(res.ReservePlus, "0.0") & " / -" & _
                                Format(res.ReserveMinus, "0.0") & ")"
    wsChart.Range("D7").Value = "Pтреб (исх / прим.):"
    wsChart.Range("E7").Value = Format(res.PReqOrig, "0.000") & " / " & Format(res.PReq, "0.000") & _
                                IIf(res.Limited, " (ограничен " & res.LimitType & ")", "")

    wsChart.Range("A8").Value = "Kд(t):"
    wsChart.Range("B8").Value = Format(res.KdUsedQuant, "0.000") & " / " & res.KdProfile

    wsChart.Range("A9").Value = "Реакция 15 с:"
    wsChart.Range("B9").Value = res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт; п. 4.2.11)"
    wsChart.Range("D9").Value = "Дефицит:"
    wsChart.Range("E9").Value = IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.00") & " %Pном, " & _
                                Format(res.DeficitSec, "0") & " с (п. 4.2.14.1)", "нет")

    wsChart.Range("A10").Value = "Вывод:"
    wsChart.Range("B10").Value = res.Verdict

    wsChart.Range("A3:A10").Font.Bold = True
    wsChart.Range("D3:D9").Font.Bold = True
    wsChart.Range("A3:E10").NumberFormat = "@"
    If res.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
        wsChart.Range("B3").Interior.Color = RGB(255, 235, 156)
        wsChart.Range("E7").Font.Color = RGB(156, 87, 0)
    ElseIf Not res.QuantPass Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If Not res.QualPass Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If res.TransientType = "Колебательный" Then wsChart.Range("B5").Font.Color = RGB(192, 0, 0)
    If res.React15Status = "Нарушение" Then wsChart.Range("B9").Font.Color = RGB(192, 0, 0)
    If Left$(res.Verdict, 9) = "НАРУШЕНИЕ" Then
        wsChart.Range("B10").Font.Color = RGB(192, 0, 0)
    Else
        wsChart.Range("B10").Font.Color = RGB(0, 97, 0)
    End If
End Sub

Private Function FindChartEndRow(ByVal wsData As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal t0 As Variant, ByVal intervalSec As Double) As Long
    Dim r As Long, t0d As Double, tCur As Double, tTarget As Double
    Dim bestRow As Long, bestDiff As Double, curDiff As Double

    If Not IsDate(t0) Then
        FindChartEndRow = r2
        Exit Function
    End If
    t0d = CDbl(CDate(t0))
    tTarget = t0d + intervalSec / 86400#
    bestRow = r1
    bestDiff = 1E+99
    For r = r1 To r2
        If IsDate(wsData.Cells(r, 1).Value) Then
            tCur = CDbl(CDate(wsData.Cells(r, 1).Value))
            If tCur >= tTarget Then
                FindChartEndRow = r
                Exit Function
            End If
            curDiff = Abs(tCur - tTarget)
            If curDiff < bestDiff Then
                bestDiff = curDiff
                bestRow = r
            End If
        End If
    Next r
    FindChartEndRow = bestRow
End Function

' ==========================================================
' Станционные суммы
' ==========================================================

Private Sub BuildStationAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                   ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long
    Dim g As TGenCfg
    Dim stations As Collection, key As Variant
    Dim done As Object

    cfgLast = LastUsedRow(wsCfg)
    Set stations = New Collection
    Set done = CreateObject("Scripting.Dictionary")

    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfg(wsCfg, r)
        If g.Enabled And g.InStationSum And Len(g.Station) > 0 Then
            If Not done.Exists(g.Station) Then
                done(g.Station) = 1
                stations.Add g.Station
            End If
        End If
NX:
    Next r

    For Each key In stations
        BuildOneStationAggregate wsRaw, wsCfg, wsSummary, CStr(key), "", st, outRow
        BuildParoprovodSubAggregates wsRaw, wsCfg, wsSummary, CStr(key), st, outRow
    Next key
End Sub

Private Sub BuildParoprovodSubAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                         ByVal stationName As String, ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long
    Dim g As TGenCfg
    Dim paropipes As Object, p As Variant

    Set paropipes = CreateObject("Scripting.Dictionary")
    cfgLast = LastUsedRow(wsCfg)
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfg(wsCfg, r)
        If g.Enabled And g.InStationSum And StationMatch(g.Station, stationName) Then
            If Len(g.Paroprovod) > 0 Then
                If paropipes.Exists(g.Paroprovod) Then
                    paropipes(g.Paroprovod) = paropipes(g.Paroprovod) + 1
                Else
                    paropipes(g.Paroprovod) = 1
                End If
            End If
        End If
NX:
    Next r

    For Each p In paropipes.Keys
        If paropipes(p) >= 1 Then
            BuildOneStationAggregate wsRaw, wsCfg, wsSummary, stationName, CStr(p), st, outRow
        End If
    Next p
End Sub

' Суммарная мощность включённых генераторов в строке r (учитываются только работающие, > порога)
Private Function SumPowerAtRow(ByVal wsRaw As Worksheet, ByRef pCols() As Long, ByVal cnt As Long, _
                               ByVal r As Long, ByVal thresholdMW As Double) As Double
    Dim i As Long, v As Double, total As Double
    total = 0
    For i = 1 To cnt
        v = NzD(wsRaw.Cells(r, pCols(i)).Value, 0)
        If v > thresholdMW Then total = total + v
    Next i
    SumPowerAtRow = total
End Function

Private Sub BuildOneStationAggregate(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long, cnt As Long
    Dim g As TGenCfg
    Dim pCols() As Long, fnchArr() As Double, sArr() As Double, pnomArr() As Double, etArr() As String
    Dim pmaxArr() As Double, pminArr() As Double
    Dim freqCol As Long, timeCol As Long
    Dim shName As String, ws As Worksheet
    Dim startRow As Long, endRow As Long, rowQ As Long, firstExceedRow As Long
    Dim p0 As Double, pNow As Double, preq As Double, preqOrig As Double
    Dim pfact As Double, dF As Double
    Dim i As Long, dfr As Double
    Dim suffix As String
    Dim pMaxSum As Double, pMinSum As Double, pNomSum As Double
    Dim reservePlus As Double, reserveMinus As Double
    Dim limited As Boolean, limitType As String
    Dim sr As TStationResult
    Dim srcLabel As String

    cfgLast = LastUsedRow(wsCfg)
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Exit Sub

    cnt = 0
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextCfg
        g = ReadGenCfg(wsCfg, r)
        If g.Enabled And g.InStationSum Then
            If StationMatch(g.Station, stationName) Then
                If Len(paropipeFilter) = 0 Or StrComp(g.Paroprovod, paropipeFilter, vbTextCompare) = 0 Then
                    cnt = cnt + 1
                    ReDim Preserve pCols(1 To cnt)
                    ReDim Preserve fnchArr(1 To cnt)
                    ReDim Preserve sArr(1 To cnt)
                    ReDim Preserve pnomArr(1 To cnt)
                    ReDim Preserve etArr(1 To cnt)
                    ReDim Preserve pmaxArr(1 To cnt)
                    ReDim Preserve pminArr(1 To cnt)
                    pCols(cnt) = FindHeaderCol(wsRaw, g.PowerHeader)
                    If pCols(cnt) = 0 Then cnt = cnt - 1: GoTo NextCfg
                    fnchArr(cnt) = g.Fnch
                    sArr(cnt) = g.SPct
                    pnomArr(cnt) = g.PNom
                    etArr(cnt) = g.EquipType
                    pmaxArr(cnt) = g.PMax
                    pminArr(cnt) = g.PMin
                    If freqCol = 0 Then freqCol = FindHeaderCol(wsRaw, g.FreqHeader)
                End If
            End If
        End If
NextCfg:
    Next r
    If cnt = 0 Then Exit Sub
    If freqCol = 0 Then Exit Sub

    srcLabel = stationName & IIf(Len(paropipeFilter) > 0, "/" & paropipeFilter, "") & " (сумма)"

    startRow = ResolveStartRow(wsRaw, timeCol, freqCol, st, MinArray(fnchArr), firstExceedRow)
    endRow = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)

    ' П. 4.2.10: сумма по включённому в работу (P > порога) и готовому к ОПРЧ оборудованию.
    ' Pном_сум / Pmax_сум / Pmin_сум - только по генераторам, находившимся в работе на старте.
    p0 = 0
    pMaxSum = 0
    pMinSum = 0
    pNomSum = 0
    Dim inWork() As Boolean
    ReDim inWork(1 To cnt)
    For i = 1 To cnt
        inWork(i) = (NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0) > st.WorkThresholdMW)
        If inWork(i) Then
            p0 = p0 + NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0)
            pMaxSum = pMaxSum + pmaxArr(i)
            pMinSum = pMinSum + pminArr(i)
            pNomSum = pNomSum + pnomArr(i)
        End If
    Next i
    If st.AvgWindowSec > 0 Then
        ' п. 4.2.6: усреднение суммарной мощности на окне перед стартом
        Dim rA As Long, r1 As Long, sumA As Double, cntA As Long
        r1 = RowByTimeOffset(wsRaw, timeCol, startRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To startRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then p0 = sumA / cntA
        sumA = 0: cntA = 0
        r1 = RowByTimeOffset(wsRaw, timeCol, endRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To endRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then pNow = sumA / cntA Else pNow = SumPowerAtRow(wsRaw, pCols, cnt, endRow, st.WorkThresholdMW)
    Else
        pNow = SumPowerAtRow(wsRaw, pCols, cnt, endRow, st.WorkThresholdMW)
    End If

    dF = MaxAbsDeviationInWindow(wsRaw, freqCol, startRow, endRow, st.FNom)
    preqOrig = 0
    Dim tSecAgg As Double
    tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(endRow, timeCol).Value)
    For i = 1 To cnt
        If inWork(i) Then
            dfr = DeadbandDeviation(dF, fnchArr(i))
            preqOrig = preqOrig + (-100# / sArr(i) * pnomArr(i) / st.FNom * DynamicKdByTime(etArr(i), tSecAgg) * dfr)
        End If
    Next i

    reservePlus = pMaxSum - p0
    reserveMinus = p0 - pMinSum
    If reservePlus < 0 Then reservePlus = 0
    If reserveMinus < 0 Then reserveMinus = 0
    preq = preqOrig
    limited = False
    limitType = ""
    If preq > reservePlus Then
        preq = reservePlus: limited = True: limitType = "Pmax"
    ElseIf preq < -reserveMinus Then
        preq = -reserveMinus: limited = True: limitType = "Pmin"
    End If
    pfact = pNow - p0

    If limited Then
        AppendLog "INFO", srcLabel, _
                  "Pтреб_сум ограничен " & limitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

    ' Количественный критерий по сумме (как для генератора)
    Dim qpct As Double, qpass As Boolean, dfrMin As Double
    dfrMin = DeadbandDeviation(dF, MinArray(fnchArr))
    If dfrMin = 0 Then
        qpct = 100: qpass = True
    ElseIf Abs(preq) < 0.000001 Then
        qpct = 100: qpass = True
    ElseIf SgnNZ(pfact) <> SgnNZ(preq) Then
        qpct = 0: qpass = False
    Else
        qpct = 100# * SafeDiv(Abs(pfact), Abs(preq), 0)
        qpass = (qpct >= (100# - st.QuantTolPct))
    End If

    ' п. 4.2.11 / 4.4.2: наличие реакции РЧВ через 15 с по суммарной мощности
    Dim baseRow As Long, row15 As Long, dp15 As Double
    Dim r15Status As String, r15Pass As Boolean, r15Note As String
    Dim preq15Sum As Double, dfr15 As Double, anyOutside As Boolean
    baseRow = IIf(firstExceedRow > 0, firstExceedRow, startRow)
    row15 = RowByTimeOffset(wsRaw, timeCol, baseRow, 15#)
    dp15 = SumPowerAtRow(wsRaw, pCols, cnt, row15, st.WorkThresholdMW) - p0
    preq15Sum = 0
    anyOutside = False
    For i = 1 To cnt
        If inWork(i) Then
            dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, freqCol).Value, st.FNom) - st.FNom, fnchArr(i))
            If dfr15 <> 0 Then
                anyOutside = True
                preq15Sum = preq15Sum + (-100# / sArr(i) * pnomArr(i) / st.FNom * DynamicKdByTime(etArr(i), 15#) * dfr15)
            End If
        End If
    Next i
    If preq15Sum > reservePlus Then preq15Sum = reservePlus
    If preq15Sum < -reserveMinus Then preq15Sum = -reserveMinus
    ClassifyReaction15 st, dp15, pNomSum, SgnNZ(preqOrig), reservePlus, reserveMinus, _
                       React15ThresholdMW(st, pNomSum, preq15Sum), (Not anyOutside), r15Status, r15Pass, r15Note
    If r15Status = "Нарушение" Then AppendLog "WARN", srcLabel, "п. 4.2.11: " & r15Note

    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    shName = MakeSheetName(stationName & suffix)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear

    ws.Range("A1:B1").Value = Array("Станция", stationName)
    If Len(paropipeFilter) > 0 Then
        ws.Range("A2:B2").Value = Array("Паропровод", paropipeFilter)
    Else
        ws.Range("A2:B2").Value = Array("Режим", "Суммарная нагрузка включённых генераторов (п. 4.2.10 МУ)")
    End If
    ws.Range("A3:B3").Value = Array("Старт (расч.)", wsRaw.Cells(startRow, timeCol).Value)
    If firstExceedRow > 0 Then
        ws.Range("A4:B4").Value = Array("Выход за fнч", wsRaw.Cells(firstExceedRow, timeCol).Value)
    Else
        ws.Range("A4:B4").Value = Array("Выход за fнч", "")
    End If
    ws.Range("A5:B5").Value = Array("P0, МВт", p0)
    ws.Range("A6:B6").Value = Array("Pтек, МВт", pNow)
    ws.Range("A7:B7").Value = Array("Pтреб, МВт", preq)
    ws.Range("A8:B8").Value = Array("Pфакт, МВт", pfact)

    ws.Cells(1, 4).Resize(1, 2).Value = Array("Pmax_сум, МВт", pMaxSum)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pmin_сум, МВт", pMinSum)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Резерв '+', МВт", reservePlus)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("Резерв '-', МВт", reserveMinus)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("Pтреб исх., МВт", preqOrig)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Ограничение", IIf(limited, "Да (" & limitType & ")", "нет"))
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Генераторов в сумме", cnt)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Pном_сум (в работе), МВт", pNomSum)

    ws.Range("A10:H10").Value = Array("Время", "Частота, Гц", "Pсум, МВт", "dPсум, МВт", "Pтреб_сум, МВт", _
                                       "dPmax_сум", "dPmin_сум", "Дефицит_сум, МВт")
    rowQ = 11
    Dim dPmaxRel As Double, dPminRel As Double
    Dim defMaxPct As Double, defSec As Double, tolSum As Double
    Dim preqStep As Double, preqCap As Double, allowed As Double, have As Double, deficit As Double, dt As Double
    dPmaxRel = pMaxSum - p0
    dPminRel = pMinSum - p0
    tolSum = 0.01 * pNomSum
    defMaxPct = 0
    defSec = 0
    For r = startRow To endRow
        ws.Cells(rowQ, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(rowQ, 2).Value = wsRaw.Cells(r, freqCol).Value

        pNow = SumPowerAtRow(wsRaw, pCols, cnt, r, st.WorkThresholdMW)
        ws.Cells(rowQ, 3).Value = pNow
        ws.Cells(rowQ, 4).Value = pNow - p0

        preqStep = 0
        tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        For i = 1 To cnt
            If inWork(i) Then
                dfr = DeadbandDeviation(NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom, fnchArr(i))
                preqStep = preqStep + (-100# / sArr(i) * pnomArr(i) / st.FNom * DynamicKdByTime(etArr(i), tSecAgg) * dfr)
            End If
        Next i
        ws.Cells(rowQ, 5).Value = preqStep
        ws.Cells(rowQ, 6).Value = dPmaxRel
        ws.Cells(rowQ, 7).Value = dPminRel

        ' п. 4.2.14.1: дефицит суммарной первичной мощности по точке (допуск 1 %Pном_сум);
        ' первые 15 с (переходный участок) не оцениваются - см. реакция РЧВ (п. 4.2.11)
        deficit = 0
        preqCap = preqStep
        If preqCap > reservePlus Then preqCap = reservePlus
        If preqCap < -reserveMinus Then preqCap = -reserveMinus
        allowed = Abs(preqCap) - tolSum
        If allowed > 0 And tSecAgg >= 15 Then
            have = SgnNZ(preqCap) * (pNow - p0)
            deficit = allowed - have
            If deficit < 0 Then deficit = 0
        End If
        ws.Cells(rowQ, 8).Value = deficit
        If deficit > 0 And pNomSum > 0 Then
            If 100# * deficit / pNomSum > defMaxPct Then defMaxPct = 100# * deficit / pNomSum
            If r < endRow Then
                dt = SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(r + 1, timeCol).Value)
            Else
                dt = 1
            End If
            If dt <= 0 Or dt > 60 Then dt = 1
            defSec = defSec + dt
        End If
        rowQ = rowQ + 1
    Next r

    ' Результат по сумме станции
    sr.Station = stationName
    sr.Paroprovod = paropipeFilter
    sr.Count = cnt
    sr.StartTime = wsRaw.Cells(startRow, timeCol).Value
    If firstExceedRow > 0 Then sr.FirstExceedTime = wsRaw.Cells(firstExceedRow, timeCol).Value Else sr.FirstExceedTime = ""
    sr.P0 = p0
    sr.PTek = SumPowerAtRow(wsRaw, pCols, cnt, endRow, st.WorkThresholdMW)
    sr.Df = dF
    sr.PReq = preq
    sr.PReqOrig = preqOrig
    sr.PFact = pfact
    sr.QuantPct = qpct
    sr.QuantPass = qpass
    sr.Limited = limited
    sr.LimitType = limitType
    sr.PMaxSum = pMaxSum
    sr.PMinSum = pMinSum
    sr.ReservePlus = reservePlus
    sr.ReserveMinus = reserveMinus
    sr.PNomSum = pNomSum
    sr.React15Status = r15Status
    sr.React15Pass = r15Pass
    sr.DP15 = dp15
    sr.React15Note = r15Note
    sr.DeficitMaxPct = defMaxPct
    sr.DeficitSec = defSec
    sr.SheetName = shName
    sr.ChartSheetName = StationChartSheetName(stationName, paropipeFilter)
    sr.Verdict = BuildStationVerdict(sr, st)

    ' Блок вердикта на листе суммы (G1:H6)
    ws.Cells(1, 7).Resize(1, 2).Value = Array("Колич. %", qpct)
    ws.Cells(2, 7).Resize(1, 2).Value = Array("Колич. статус", IIf(limited, "Ограничен " & limitType, IIf(qpass, "ОК", "Нарушение")))
    ws.Cells(3, 7).Resize(1, 2).Value = Array("Реакция 15 с (п. 4.2.11)", r15Status & " (dP15 = " & Format(dp15, "0.000") & " МВт)")
    ws.Cells(4, 7).Resize(1, 2).Value = Array("Дефицит (п. 4.2.14.1)", _
        IIf(defSec > 0, "до " & Format(defMaxPct, "0.00") & " %Pном_сум, " & Format(defSec, "0") & " с", "нет"))
    ws.Cells(5, 7).Resize(1, 2).Value = Array("Kд(t) по типам", "см. профили Config AA:AI (по каждому генератору)")

    ws.Columns("A:H").AutoFit
    ws.Cells(6, 7).Value = "Вывод"
    ws.Cells(6, 8).Value = sr.Verdict
    ws.Cells(6, 7).Font.Bold = True
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then ws.Cells(6, 8).Font.Color = RGB(192, 0, 0) Else ws.Cells(6, 8).Font.Color = RGB(0, 97, 0)

    ws.Range("A1:A8").NumberFormat = "@"
    ws.Range("B1:B2").NumberFormat = "@"
    ws.Range("B3:B4").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B5:B8").NumberFormat = "0.000"
    ws.Range("D1:D8").NumberFormat = "@"
    ws.Range("E1:E5").NumberFormat = "0.000"
    ws.Range("E6").NumberFormat = "@"
    ws.Range("E7").NumberFormat = "0"
    ws.Range("E8").NumberFormat = "0.000"
    ws.Range("G1:G6").NumberFormat = "@"
    ws.Range("H1").NumberFormat = "0.0"
    ws.Range("H2:H6").NumberFormat = "@"
    ws.Range("A10:H10").NumberFormat = "@"
    ws.Range("A11:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B11:H100000").NumberFormat = "0.000"

    AppendLog "INFO", srcLabel, sr.Verdict

    WriteStationChartSheet stationName, paropipeFilter, rowQ - 1, st, MinArray(fnchArr), dF, sr

    ' Строка в Summary
    WriteSummaryStationRow wsSummary, outRow, sr
    outRow = outRow + 1

    ' п. 4.2.10: пометка в строках генераторов станции о статусе суммы
    If Len(paropipeFilter) = 0 Then
        AnnotateGeneratorRowsWithStationVerdict wsSummary, outRow - 1, stationName, (Left$(sr.Verdict, 2) = "ОК")
    End If
End Sub

Private Sub AnnotateGeneratorRowsWithStationVerdict(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                                    ByVal stationName As String, ByVal stationOk As Boolean)
    Dim r As Long, addText As String, cur As String
    If stationOk Then
        addText = "Станция в целом: ОК (п. 4.2.10 МУ - индивидуальная оценка приведена справочно)"
    Else
        addText = "Станция в целом: не соответствует (п. 4.2.10 МУ - индивидуальная оценка обязательна)"
    End If
    For r = 2 To lastRow
        If StationMatch(CStr(ws.Cells(r, 1).Value), stationName) Then
            If Left$(CStr(ws.Cells(r, 2).Value), 5) <> "СУММА" Then
                cur = CStr(ws.Cells(r, 40).Value)
                If InStr(cur, "Станция в целом:") = 0 Then
                    ws.Cells(r, 40).Value = cur & IIf(Len(cur) > 0, " | ", "") & addText
                End If
            End If
        End If
    Next r
End Sub

' Вывод по сумме станции с учётом п. 4.2.10 МУ (индивидуальная оценка генераторов
' требуется только при нарушении по сумме)
Private Function BuildStationVerdict(ByRef sr As TStationResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String, objName As String
    objName = "по сумме " & IIf(Len(sr.Paroprovod) > 0, "паропровода " & sr.Paroprovod & " ", "станции ")
    bad = ""
    remark = ""
    If sr.Limited Then
        remark = remark & "Pтреб_сум ограничен " & sr.LimitType & " (оценка по располагаемому резерву); "
    ElseIf Not sr.QuantPass Then
        bad = bad & "колич. критерий (" & Format(sr.QuantPct, "0") & " % от Pтреб_сум); "
    End If
    If sr.React15Status = "Нарушение" Then
        bad = bad & "нет реакции РЧВ через 15 с (п. 4.2.11); "
    End If
    If sr.DeficitSec > 0 Then
        remark = remark & "дефицит до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум в течение " & _
                 Format(sr.DeficitSec, "0") & " с (п. 4.2.14.1); "
    End If

    If Len(bad) > 0 Then
        txt = "НАРУШЕНИЕ " & objName & ": " & bad & _
              "По п. 4.2.10 МУ требуется индивидуальная оценка каждого генератора - см. строки генераторов станции. "
    Else
        txt = "ОК " & objName & "(" & sr.Count & " ген.): количественно и по реакции 15 с соответствует. " & _
              "По п. 4.2.10 МУ индивидуальная оценка генераторов не требуется (строки генераторов приведены справочно). "
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    txt = Trim$(txt)
    If Right$(txt, 1) = ";" Then txt = Left$(txt, Len(txt) - 1)
    BuildStationVerdict = txt
End Function

Private Function StationChartSheetName(ByVal stationName As String, ByVal paropipeFilter As String) As String
    Dim suffix As String, dataName As String, chartName As String
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    dataName = MakeSheetName(stationName & suffix)
    chartName = MakeSheetName(stationName & suffix & CHART_SUFFIX)
    If StrComp(chartName, dataName, vbTextCompare) = 0 Then
        chartName = Left$(dataName, 28) & "_Гр"
    End If
    StationChartSheetName = chartName
End Function

Private Sub WriteStationChartSheet(ByVal stationName As String, ByVal paropipeFilter As String, _
                                   ByVal lastDataRow As Long, ByRef st As TSettings, _
                                   ByVal fnchMin As Double, ByVal dFmax As Double, ByRef sr As TStationResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataName As String, chartName As String
    Dim chartObj As ChartObject
    Dim endRow As Long, startRow As Long
    Dim suffix As String, ySpan As Double

    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    dataName = MakeSheetName(stationName & suffix)
    chartName = StationChartSheetName(stationName, paropipeFilter)

    Set wsData = ThisWorkbook.Worksheets(dataName)
    Set wsChart = EnsureSheet(chartName)
    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    wsChart.Range("A1").Value = "График ОПРЧ (сумма): " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "")
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    ' Блок вердикта по сумме (A3:E6)
    wsChart.Range("A3").Value = "Количественный:"
    wsChart.Range("B3").Value = IIf(sr.Limited, "Ограничен " & sr.LimitType & " (" & Format(sr.QuantPct, "0") & " % от доступного)", _
                                    IIf(sr.QuantPass, "ОК (" & Format(sr.QuantPct, "0") & " %)", "Нарушение (" & Format(sr.QuantPct, "0") & " %)"))
    wsChart.Range("D3").Value = "Pтреб / Pфакт, МВт:"
    wsChart.Range("E3").Value = Format(sr.PReq, "0.000") & " / " & Format(sr.PFact, "0.000")
    wsChart.Range("A4").Value = "Реакция 15 с:"
    wsChart.Range("B4").Value = sr.React15Status & " (dP15 = " & Format(sr.DP15, "0.000") & " МВт; п. 4.2.11)"
    wsChart.Range("D4").Value = "Дефицит:"
    wsChart.Range("E4").Value = IIf(sr.DeficitSec > 0, "до " & Format(sr.DeficitMaxPct, "0.00") & " %Pном_сум, " & _
                                Format(sr.DeficitSec, "0") & " с", "нет")
    wsChart.Range("A5").Value = "Pmax / Pmin сум, МВт:"
    wsChart.Range("B5").Value = Format(sr.PMaxSum, "0.0") & " / " & Format(sr.PMinSum, "0.0") & _
                                " (резерв +" & Format(sr.ReservePlus, "0.0") & " / -" & Format(sr.ReserveMinus, "0.0") & ")"
    wsChart.Range("D5").Value = "Генераторов:"
    wsChart.Range("E5").Value = CStr(sr.Count)
    wsChart.Range("A6").Value = "Вывод:"
    wsChart.Range("B6").Value = sr.Verdict
    wsChart.Range("A3:A6").Font.Bold = True
    wsChart.Range("D3:D5").Font.Bold = True
    wsChart.Range("A3:E6").NumberFormat = "@"
    If sr.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
    ElseIf Not sr.QuantPass Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If sr.React15Status = "Нарушение" Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then wsChart.Range("B6").Font.Color = RGB(192, 0, 0) Else wsChart.Range("B6").Font.Color = RGB(0, 97, 0)

    startRow = 11
    If lastDataRow < startRow Then Exit Sub
    endRow = FindChartEndRow(wsData, startRow, lastDataRow, wsData.Cells(startRow, 1).Value, st.ChartIntervalSec)
    If endRow < startRow Then endRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 110, 1020, 520)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Суммарный мониторинг ОПРЧ: " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "")
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 4, "dPсум, МВт", False
    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 5, "Pтреб_сум, МВт", False
    AddLimitLineSeries chartObj.Chart, wsData, startRow, endRow, 1, 6, "Pmax_сум (dPmax)"
    AddLimitLineSeries chartObj.Chart, wsData, startRow, endRow, 1, 7, "Pmin_сум (dPmin)"
    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 2, "Частота, Гц", True

    On Error Resume Next
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(dFmax) * 1.2, 2# * fnchMin)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0
End Sub

' ==========================================================
' Summary / оформление / валидация / лог
' ==========================================================

Private Sub WriteSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim note As String
    Dim overshootStr As String
    Dim quantStatusStr As String, limitStr As String
    note = ""
    If res.Limited Then
        quantStatusStr = "Ограничен " & res.LimitType
        limitStr = res.LimitType
    ElseIf res.QuantPass Then
        quantStatusStr = "ОК"
        limitStr = ""
    Else
        quantStatusStr = "Нарушение"
        limitStr = ""
    End If

    If Not res.QuantPass And Not res.Limited Then note = "Колич. критерий не выполнен"
    If res.Limited Then note = Trim$(note & "; Pтреб ограничен " & res.LimitType)
    If Abs(res.P0) < 0.001 And Abs(res.PFact) < 0.001 Then note = "Нет первичного отклика; проверьте генератор/датчик"
    If Len(res.AmplitudeTag) > 0 And res.AmplitudeTag = "Слабое" Then _
        note = Trim$(note & "; Слабое возмущение (< 3 %Pном)")
    If Len(res.AmplitudeTag) > 0 And res.AmplitudeTag = "Избыточное" Then _
        note = Trim$(note & "; Возмущение > 10 %Pном (вне нормативного диапазона)")
    If Len(note) > 0 And Left$(note, 2) = "; " Then note = Mid$(note, 3)
    overshootStr = IIf(res.Overshoot, "Да", "")

    If res.React15Status = "Нарушение" Then note = Trim$(note & "; " & res.React15Note)
    If Len(note) > 0 And Left$(note, 2) = "; " Then note = Mid$(note, 3)

    ws.Cells(r, 1).Resize(1, 40).Value = Array( _
        g.Station, g.Generator, g.EquipType, _
        res.StartTime, res.FirstExceedTime, _
        res.P0, res.PTek, res.PsteadyAvg, _
        res.Df, res.Dfr, _
        res.PReq, res.PFact, _
        Round(res.AmplPctPnom, 2), res.AmplitudeTag, _
        res.QuantPct, quantStatusStr, overshootStr, _
        res.TransientType, res.NumExtrema, _
        IIf(res.QualPass, "ОК", "Нарушение"), _
        IIf(res.QualT5Pass, "ОК", "Нарушение"), _
        IIf(res.QualT10Pass, "ОК", "Нарушение"), _
        IIf(res.QualSteadyPass, "ОК", "Нарушение"), _
        res.QualFailedList, _
        FormatSecOrNA(res.T5FactSec), FormatSecOrNA(res.T10FactSec), _
        GeneratorSheetName(g), GeneratorChartSheetName(g), _
        res.PMaxEff, res.PMinEff, res.ReservePlus, res.ReserveMinus, _
        res.PReqOrig, limitStr, _
        note, _
        res.React15Status, Round(res.DP15, 3), Round(res.DeficitMaxPct, 2), Round(res.DeficitSec, 0), _
        res.Verdict _
    )
End Sub

Private Sub WriteSummaryInvalid(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg)
    ws.Cells(r, 1).Resize(1, 40).Value = Array( _
        g.Station, g.Generator, g.EquipType, "", "", "", "", "", "", "", "", "", "", "", "", _
        "Нарушение", "", "", 0, "Н/Д", "Н/Д", "Н/Д", "Н/Д", "", "", "", "Config", "", _
        "", "", "", "", "", "", _
        "Не заполнен обязательный параметр config", _
        "Н/Д", "", "", "", "НАРУШЕНИЕ: не заполнены обязательные параметры Config" _
    )
End Sub

' Строка Summary по сумме станции (п. 4.2.10)
Private Sub WriteSummaryStationRow(ByVal ws As Worksheet, ByVal r As Long, ByRef sr As TStationResult)
    Dim quantStatusStr As String, limitStr As String, genLabel As String
    If sr.Limited Then
        quantStatusStr = "Ограничен " & sr.LimitType
        limitStr = sr.LimitType
    ElseIf sr.QuantPass Then
        quantStatusStr = "ОК"
        limitStr = ""
    Else
        quantStatusStr = "Нарушение"
        limitStr = ""
    End If
    genLabel = "СУММА" & IIf(Len(sr.Paroprovod) > 0, " (" & sr.Paroprovod & ")", "")
    Dim amplVal As Variant
    If sr.PNomSum > 0 Then amplVal = Round(100# * Abs(sr.PReqOrig) / sr.PNomSum, 2) Else amplVal = ""

    ws.Cells(r, 1).Resize(1, 40).Value = Array( _
        sr.Station, genLabel, "Сумма станции (" & sr.Count & " ген.)", _
        sr.StartTime, sr.FirstExceedTime, _
        sr.P0, sr.PTek, "", _
        sr.Df, "", _
        sr.PReq, sr.PFact, _
        amplVal, "", _
        sr.QuantPct, quantStatusStr, "", _
        "", "", _
        "Н/П", "Н/П", "Н/П", "Н/П", _
        "", "", "", _
        sr.SheetName, sr.ChartSheetName, _
        sr.PMaxSum, sr.PMinSum, sr.ReservePlus, sr.ReserveMinus, _
        sr.PReqOrig, limitStr, _
        IIf(sr.React15Status = "Нарушение", sr.React15Note, "Оценка по суммарной мощности станции (п. 4.2.10 МУ)"), _
        sr.React15Status, Round(sr.DP15, 3), Round(sr.DeficitMaxPct, 2), Round(sr.DeficitSec, 0), _
        sr.Verdict _
    )
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 40)).Font.Bold = True
End Sub

Private Sub ApplySummaryConditionalFormat(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ws.Range("A2:AN" & lastRow).Interior.Pattern = xlNone

    ' Строки сумм станций (B начинается с "СУММА") - подсветка A:C
    Dim rr As Long
    For rr = 2 To lastRow
        If Left$(CStr(ws.Cells(rr, 2).Value), 5) = "СУММА" Then
            ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 3)).Interior.Color = RGB(221, 235, 247)
        End If
    Next rr

    ' Колонки текстовых статусов: P, Q, T, U, V, W (колич/перерег/кач/t5/t10/уст)
    ApplyStatusCF ws, ws.Range("P2:P" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("T2:T" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("U2:U" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("V2:V" & lastRow), "ОК", "Нарушение"
    ' Реакция 15 с (AJ) и Вывод (AN)
    ApplyStatusCF ws, ws.Range("AJ2:AJ" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("AN2:AN" & lastRow), "ОК", "Нарушение"
    ' Дефицит (AL): любое значение > 0 - оранжевым (замечание)
    Set rng = ws.Range("AL2:AL" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    
    ' Кач.уст (W): "Нарушение" подсвечиваем как замечание (оранжевым), не красным.
    Set rng = ws.Range("W2:W" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Нарушение", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="ОК", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ' Количественный статус P: 'Ограничен' -> жёлтым
    Set rng = ws.Range("P2:P" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Ограничен", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Ограничение AH (34): любой непустой текст -> жёлтым
    Set rng = ws.Range("AH2:AH" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Pmax", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Pmin", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Перерегулирование (Q) - желтый при "Да"
    Set rng = ws.Range("Q2:Q" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Да", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Характер процесса (R)
    Set rng = ws.Range("R2:R" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Колебательный", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Апериодический", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Монотонный", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
    End With

    ' Масштаб события (N)
    Set rng = ws.Range("N2:N" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Слабое", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 225, 242)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Избыточное", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 217, 102)
    End With

    ' Колич. % (O) - цветовая шкала 0..100..200
    Set rng = ws.Range("O2:O" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueNumber
        .ColorScaleCriteria(1).Value = 0
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).Type = xlConditionValueNumber
        .ColorScaleCriteria(2).Value = 100
        .ColorScaleCriteria(2).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(3).Type = xlConditionValueNumber
        .ColorScaleCriteria(3).Value = 200
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
    End With

    ws.Range("A1:AN1").Font.Bold = True
    ws.Range("A1:AN1").Interior.Color = RGB(217, 225, 242)
End Sub

Private Sub ApplyStatusCF(ByVal ws As Worksheet, ByVal rng As Range, ByVal okText As String, ByVal badText As String)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:=badText, TextOperator:=xlContains)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:=okText, TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
End Sub

Private Sub WriteVersionStamp(ByVal wsSummary As Worksheet, ByVal wsRaw As Worksheet, ByVal t0Run As Double)
    Dim col As Long
    col = 42   ' колонка AP - правее данных Summary (после AN)
    wsSummary.Cells(1, col).Value = "Мониторинг ОПРЧ"
    wsSummary.Cells(2, col).Value = "Версия"
    wsSummary.Cells(2, col + 1).Value = OPRCH_VERSION
    wsSummary.Cells(3, col).Value = "Запуск"
    wsSummary.Cells(3, col + 1).Value = Now
    wsSummary.Cells(3, col + 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsSummary.Cells(4, col).Value = "Длительность, с"
    wsSummary.Cells(4, col + 1).Value = Round(Timer - t0Run, 1)
    wsSummary.Cells(5, col).Value = "Файл книги"
    wsSummary.Cells(5, col + 1).Value = ThisWorkbook.Name
    wsSummary.Cells(6, col).Value = "Строк RawData"
    wsSummary.Cells(6, col + 1).Value = LastUsedRow(wsRaw) - 1
    wsSummary.Range(wsSummary.Cells(1, col), wsSummary.Cells(6, col)).Font.Bold = True
    wsSummary.Range(wsSummary.Cells(1, col), wsSummary.Cells(6, col + 1)).Interior.Color = RGB(240, 240, 240)
End Sub

' ==========================================================
' Валидация и лог
' ==========================================================

Private Sub InitLog()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Уровень", "Источник", "Сообщение", "Время")
    ws.Range("A1:D1").Font.Bold = True
    m_LogRow = 2
End Sub

Private Sub AppendLog(ByVal level As String, ByVal source As String, ByVal message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then InitLog: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If m_LogRow < 2 Then m_LogRow = 2
    ws.Cells(m_LogRow, 1).Value = level
    ws.Cells(m_LogRow, 2).Value = source
    ws.Cells(m_LogRow, 3).Value = message
    ws.Cells(m_LogRow, 4).Value = Now
    ws.Cells(m_LogRow, 4).NumberFormat = "hh:mm:ss"
    m_LogRow = m_LogRow + 1
End Sub

Private Sub FinalizeLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Columns("A:D").AutoFit
    ws.Range("A1:D1").Interior.Color = RGB(217, 225, 242)
    If m_LogRow <= 2 Then
        ws.Range("A2").Value = "INFO"
        ws.Range("B2").Value = "-"
        ws.Range("C2").Value = "Замечаний не обнаружено"
        ws.Range("D2").Value = Now
        ws.Range("D2").NumberFormat = "hh:mm:ss"
    End If
End Sub

Private Sub ValidateInputs(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByRef st As TSettings, ByVal timeCol As Long)
    Dim lastR As Long, r As Long, prev As Double, cur As Double, gap As Double, maxGap As Double
    Dim cfgLast As Long, g As TGenCfg, headerRow As Long
    Dim paropipes As Object, stationParopipe As Object, st2 As String, kk As Variant

    ' RawData: шаг
    lastR = LastUsedRow(wsRaw)
    If lastR < 10 Then
        AppendLog "WARN", "RawData", "Слишком мало строк данных: " & (lastR - 1)
        Exit Sub
    End If
    prev = 0
    maxGap = 0
    For r = 2 To WorksheetFunction.Min(lastR, 500)
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            cur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            If prev > 0 Then
                gap = (cur - prev) * 86400#
                If gap > maxGap Then maxGap = gap
                If gap > 6 Then AppendLog "WARN", "RawData", "Разрыв по времени " & Format(gap, "0.0") & " c в строке " & r & " (норматив <= 5 c)"
            End If
            prev = cur
        End If
    Next r
    If maxGap > 0 Then AppendLog "INFO", "RawData", "Максимальный шаг по времени в первых строках: " & Format(maxGap, "0.0") & " c"

    ' Config: валидация параметров
    cfgLast = LastUsedRow(wsCfg)
    Set stationParopipe = CreateObject("Scripting.Dictionary")
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfg(wsCfg, r)
        If Not g.Enabled Then GoTo NX
        If g.PNom <= 0 Then AppendLog "WARN", g.Station & "/" & g.Generator, "Pном <= 0"
        If g.SPct <= 1 Or g.SPct > 15 Then AppendLog "WARN", g.Station & "/" & g.Generator, "S вне [1..15] %: " & g.SPct
        If g.Fnch < 0 Or g.Fnch > 0.5 Then AppendLog "WARN", g.Station & "/" & g.Generator, "fнч вне [0..0.5] Гц: " & g.Fnch
        If g.T10Sec < g.T5Sec Then AppendLog "WARN", g.Station & "/" & g.Generator, "t10 < t5 (" & g.T10Sec & " < " & g.T5Sec & ")"
        ' Pmax/Pmin
        If g.PMax > 0 And g.PMax < g.PMin Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax < Pmin (" & g.PMax & " < " & g.PMin & ")"
        If g.PMax > 0 And g.PNom > 0 And g.PMax > 1.3 * g.PNom Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax > 1.3*Pном (проверьте ед. измерения)"
        If g.PMin < 0 Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmin < 0: " & g.PMin
        ' Паропровод: все или никто в рамках станции
        If g.InStationSum Then
            If stationParopipe.Exists(g.Station) Then
                st2 = CStr(stationParopipe(g.Station))
                If (st2 = "YES" And Len(g.Paroprovod) = 0) Or (st2 = "NO" And Len(g.Paroprovod) > 0) Then
                    AppendLog "WARN", g.Station, "Колонка Паропровод заполнена частично - проверьте согласованность"
                End If
            Else
                stationParopipe(g.Station) = IIf(Len(g.Paroprovod) > 0, "YES", "NO")
            End If
        End If
NX:
    Next r

    If st.QuantIntervalSec < 60 Then AppendLog "WARN", "Settings", "Колич. интервал < 60 с: " & st.QuantIntervalSec
    If st.SteadyWindowSec < 10 Then AppendLog "WARN", "Settings", "Окно установившегося < 10 с: " & st.SteadyWindowSec
    If st.React15ThrPct > 2 Then AppendLog "WARN", "Settings", "Порог реакции 15 с > 2 %Pном (" & st.React15ThrPct & ") - п. 4.2.11 требует лишь наличия реакции"
    If st.AvgWindowSec > 30 Then AppendLog "WARN", "Settings", "Окно усреднения P0/Pтек > 30 с: " & st.AvgWindowSec

    ' Справка по профилям Kд(t): значения в контрольных точках 15 с и колич. интервала
    If Not m_KdProfiles Is Nothing Then
        For Each kk In m_KdProfiles.Keys
            On Error Resume Next
            AppendLog "INFO", "Kд(t) " & CStr(kk), "Kд(15 с) = " & Format(DynamicKdByTime(CStr(kk), 15#), "0.00") & _
                      "; Kд(" & Format(st.QuantIntervalSec, "0") & " с) = " & Format(DynamicKdByTime(CStr(kk), st.QuantIntervalSec), "0.00") & _
                      " (МУ: ТЭС 0.50 до 15 с -> 1.0 к 300/360/420 с; ПГУ 0.25 до 15 с -> 0.5 к 30 с -> 1.0 к 120 с)"
            On Error GoTo 0
        Next kk
    End If
End Sub

' ==========================================================
' Параметры / пресеты / чтение конфига
' ==========================================================

Private Function ReadSettings(ByVal wsCfg As Worksheet) As TSettings
    Dim st As TSettings
    Dim valCol As Long
    ' В 1.4.0 глобальные настройки перенесены в колонки W/X (23/24),
    ' чтобы освободить T/U для Pmax/Pmin у генераторов.
    ' Старые книги (до 1.3.x) продолжаем читать по адресу T/U (20/21),
    ' если W пустой.
    valCol = 24
    If Trim$(CStr(wsCfg.Cells(2, 23).Value)) = "" Then valCol = 21
    st.FNom = NzD(wsCfg.Cells(2, valCol).Value, 50#)
    st.AutoStart = (NzD(wsCfg.Cells(4, valCol).Value, 1) <> 0)
    st.QuantIntervalSec = NzD(wsCfg.Cells(5, valCol).Value, 82#)
    st.QuantTolPct = NzD(wsCfg.Cells(6, valCol).Value, 10#)
    st.WorkThresholdMW = NzD(wsCfg.Cells(7, valCol).Value, 1#)
    st.PreBufferSec = NzD(wsCfg.Cells(8, valCol).Value, 5#)
    st.ChartIntervalSec = NzD(wsCfg.Cells(9, valCol).Value, 120#)
    st.SteadyWindowSec = NzD(wsCfg.Cells(10, valCol).Value, 30#)
    st.React15ThrPct = NzD(wsCfg.Cells(11, valCol).Value, 0.5)
    st.AvgWindowSec = NzD(wsCfg.Cells(12, valCol).Value, 0#)
    If st.React15ThrPct < 0 Then st.React15ThrPct = 0
    If st.AvgWindowSec < 0 Then st.AvgWindowSec = 0
    If IsDate(wsCfg.Cells(3, valCol).Value) Then
        st.EventStart = CDate(wsCfg.Cells(3, valCol).Value)
    Else
        st.EventStart = 0
    End If
    ReadSettings = st
End Function

Private Function ReadGenCfg(ByVal ws As Worksheet, ByVal r As Long) As TGenCfg
    Dim g As TGenCfg, pr As TGenCfg
    g.Station = Trim$(CStr(ws.Cells(r, 1).Value))
    g.Generator = Trim$(CStr(ws.Cells(r, 2).Value))
    g.PowerHeader = Trim$(CStr(ws.Cells(r, 3).Value))
    g.FreqHeader = Trim$(CStr(ws.Cells(r, 4).Value))
    g.EquipType = Trim$(CStr(ws.Cells(r, 5).Value))
    g.PNom = NzD(ws.Cells(r, 6).Value, 0)
    g.SPct = NzD(ws.Cells(r, 7).Value, 0)
    g.Fnch = NzD(ws.Cells(r, 8).Value, -1)
    g.Enabled = (NzD(ws.Cells(r, 9).Value, 1) <> 0)
    g.QualEnabled = (NzD(ws.Cells(r, 10).Value, 1) <> 0)
    g.T5Sec = NzD(ws.Cells(r, 11).Value, 0)
    g.Dp5Pct = NzD(ws.Cells(r, 12).Value, 0)
    g.T10Sec = NzD(ws.Cells(r, 13).Value, 0)
    g.Dp10Pct = NzD(ws.Cells(r, 14).Value, 0)
    g.SteadyTolPct = NzD(ws.Cells(r, 15).Value, 0)
    g.InStationSum = (NzD(ws.Cells(r, 16).Value, 0) <> 0)
    g.CheckSteady = (NzD(ws.Cells(r, 17).Value, 1) <> 0)
    g.Paroprovod = Trim$(CStr(ws.Cells(r, 18).Value))
    ' Pmax/Pmin: пусто = значения по умолчанию (Pmax=Pном, Pmin=0)
    If Trim$(CStr(ws.Cells(r, 19).Value)) = "" Then
        g.PMax = g.PNom
    Else
        g.PMax = NzD(ws.Cells(r, 19).Value, g.PNom)
    End If
    If Trim$(CStr(ws.Cells(r, 20).Value)) = "" Then
        g.PMin = 0
    Else
        g.PMin = NzD(ws.Cells(r, 20).Value, 0)
    End If
    If g.PMax <= 0 Then g.PMax = g.PNom
    If g.PMin < 0 Then g.PMin = 0
    If g.PMax < g.PMin Then g.PMax = g.PMin

    pr = GetPreset(g.EquipType)
    If g.T5Sec <= 0 Then g.T5Sec = pr.T5Sec
    If g.Dp5Pct <= 0 Then g.Dp5Pct = pr.Dp5Pct
    If g.T10Sec <= 0 Then g.T10Sec = pr.T10Sec
    If g.Dp10Pct <= 0 Then g.Dp10Pct = pr.Dp10Pct
    If g.SteadyTolPct <= 0 Then g.SteadyTolPct = pr.SteadyTolPct
    If g.Fnch < 0 Then g.Fnch = pr.Fnch

    If Len(g.PowerHeader) = 0 Then g.PowerHeader = g.Generator
    If Len(g.FreqHeader) = 0 Then g.FreqHeader = "Частота"
    ReadGenCfg = g
End Function

Private Function GetPreset(ByVal equipType As String) As TGenCfg
    Dim g As TGenCfg, et As String
    et = UCase$(Trim$(equipType))
    ' Нормативные ступени динамики по МУ проверки готовности (Приложение 1 к
    ' Техтребованиям, табл. Е.1-Е.6, бланки-графики Ж.1-Ж.3) и параметрам Kд(t)
    ' Приложения 14:
    '   ПСУ блочные: 5 %Pном за 15 с, 10 %Pном за 5 (газомазутные) / 6 (пылеугольные) мин;
    '   ПСУ неблочные (общий паропровод): 5 % за 15 с, 10 % за 7 мин (по сумме ТГ);
    '   ГТУ / ПГУ утилизационные: 2.5 % за 15 с, 5 % за 30 с, 10 % за 2 мин;
    '   ПГУ сбросные: как ПСУ (5 % за 15 с, 10 % за 5-6 мин);
    '   ГПА: 10 % за 1 мин.
    ' Установившееся отклонение: не более ±1 %Pном (п. 6.2.2.11 Приложения 1, п. 4.2.7 Приложения 14).
    g.T5Sec = 15
    g.Dp5Pct = 5
    g.T10Sec = 420
    g.Dp10Pct = 10
    g.SteadyTolPct = 1
    g.Fnch = 0.075

    If InStr(et, "ГПА") > 0 Then
        g.T10Sec = 60
    ElseIf InStr(et, "ГТУ") > 0 Then
        g.Dp5Pct = 2.5
        g.T10Sec = 120
    ElseIf InStr(et, "ПГУ_СБРОСН") > 0 Or InStr(et, "СБРОСН") > 0 Then
        g.T10Sec = 300
    ElseIf InStr(et, "ПГУ_УТИЛ") > 0 Or InStr(et, "УТИЛИЗ") > 0 Or InStr(et, "ПГУ") > 0 Then
        g.Dp5Pct = 2.5
        g.T10Sec = 120
    ElseIf InStr(et, "УГОЛЬ") > 0 Or InStr(et, "ПЫЛЕУГ") > 0 Then
        g.T10Sec = 360
    ElseIf InStr(et, "ПТУ_БЛОК") > 0 Then
        g.T10Sec = 300
    ElseIf InStr(et, "ПТУ_НЕБЛОК") > 0 Or InStr(et, "НЕБЛОК") > 0 Then
        g.T10Sec = 420
    ElseIf InStr(et, "ПТУ") > 0 Then
        g.T10Sec = 420
    End If
    GetPreset = g
End Function

Private Function ValidateGenCfg(ByRef g As TGenCfg) As Boolean
    ValidateGenCfg = (Len(g.Station) > 0 And Len(g.Generator) > 0 And Len(g.PowerHeader) > 0 And Len(g.FreqHeader) > 0 _
                      And g.PNom > 0 And g.SPct > 0 And g.Fnch >= 0)
End Function

' ==========================================================
' Алгоритмические помощники
' ==========================================================

Private Function ResolveStartRow(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal freqCol As Long, _
                                 ByRef st As TSettings, ByVal fnch As Double, _
                                 ByRef firstExceedRow As Long) As Long
    Dim lastR As Long, r As Long
    Dim prevAbs As Double, curAbs As Double
    Dim bestRow As Long, bestAbs As Double, val As Double, t As Double, dt As Double

    firstExceedRow = 0
    lastR = LastUsedRow(wsRaw)
    If lastR < 3 Then ResolveStartRow = 2: Exit Function

    If Not st.AutoStart And st.EventStart > 0 Then
        bestRow = 2: bestAbs = 1E+99
        For r = 2 To lastR
            If IsDate(wsRaw.Cells(r, timeCol).Value) Then
                t = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
                dt = Abs(CDbl(t - st.EventStart))
                If dt < bestAbs Then bestAbs = dt: bestRow = r
            End If
        Next r
        ResolveStartRow = bestRow
        firstExceedRow = bestRow
        Exit Function
    End If

    prevAbs = Abs(NzD(wsRaw.Cells(2, freqCol).Value, st.FNom) - st.FNom)
    For r = 3 To lastR
        curAbs = Abs(NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom)
        If prevAbs <= fnch And curAbs > fnch Then
            firstExceedRow = r
            ResolveStartRow = r - 1
            Exit Function
        End If
        prevAbs = curAbs
    Next r

    bestRow = 2: bestAbs = -1
    For r = 2 To lastR
        val = NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom
        If Abs(val) > bestAbs Then bestAbs = Abs(val): bestRow = r
    Next r
    ResolveStartRow = bestRow
    firstExceedRow = bestRow
End Function

Private Function RowByTimeOffset(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal startRow As Long, ByVal offsetSec As Double) As Long
    Dim t0 As Double, tTarget As Double
    Dim r As Long, lastR As Long, bestRow As Long
    Dim curT As Double, bestDiff As Double, curDiff As Double
    Dim scanFrom As Long, scanTo As Long

    If Not IsDate(wsRaw.Cells(startRow, timeCol).Value) Then RowByTimeOffset = startRow: Exit Function
    t0 = CDbl(CDate(wsRaw.Cells(startRow, timeCol).Value))
    tTarget = t0 + offsetSec / 86400#
    lastR = LastUsedRow(wsRaw)
    bestRow = startRow: bestDiff = 1E+99

    If offsetSec >= 0 Then scanFrom = startRow: scanTo = lastR Else scanFrom = 2: scanTo = startRow

    For r = scanFrom To scanTo
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            curT = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            curDiff = Abs(CDbl(curT - tTarget))
            If curDiff < bestDiff Then bestDiff = curDiff: bestRow = r
        End If
    Next r
    RowByTimeOffset = bestRow
End Function

Private Function MaxAbsDeviationInWindow(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal fNom As Double) As Double
    Dim r As Long, d As Double, best As Double
    best = 0
    For r = r1 To r2
        d = NzD(wsRaw.Cells(r, freqCol).Value, fNom) - fNom
        If Abs(d) > Abs(best) Then best = d
    Next r
    MaxAbsDeviationInWindow = best
End Function

Private Function DeadbandDeviation(ByVal dF As Double, ByVal fnch As Double) As Double
    If Abs(dF) <= fnch Then DeadbandDeviation = 0 Else DeadbandDeviation = Sgn(dF) * (Abs(dF) - fnch)
End Function

Private Function SecBetween(ByVal t1 As Variant, ByVal t2 As Variant) As Double
    If IsDate(t1) And IsDate(t2) Then
        SecBetween = Abs(CDbl(CDate(t2) - CDate(t1))) * 86400#
    Else
        SecBetween = 0
    End If
End Function

Private Sub AddSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                      ByVal xCol As Long, ByVal yCol As Long, ByVal titleText As String, ByVal secondaryAxis As Boolean)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = titleText
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    If secondaryAxis Then s.AxisGroup = xlSecondary
End Sub

Private Sub AddMarkerSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                            ByVal xCol As Long, ByVal yCol As Long, ByVal nm As String)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    On Error Resume Next
    s.ChartType = xlLine
    s.Format.Line.Weight = 1.25
    s.Format.Line.DashStyle = msoLineDash
    s.MarkerStyle = xlMarkerStyleNone
    On Error GoTo 0
End Sub

Private Sub AddLimitLineSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                               ByVal xCol As Long, ByVal yCol As Long, ByVal nm As String)
    ' Горизонтальные уровни Pmax / Pmin: красная жирная пунктирная линия без маркеров.
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    On Error Resume Next
    s.ChartType = xlLine
    s.Format.Line.Weight = 1.75
    s.Format.Line.DashStyle = msoLineLongDash
    s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    s.MarkerStyle = xlMarkerStyleNone
    On Error GoTo 0
End Sub

Private Function AmplitudeTag(ByVal amplPct As Double, ByVal isEvent As Boolean) As String
    If Not isEvent Then AmplitudeTag = "": Exit Function
    If amplPct < 3# Then
        AmplitudeTag = "Слабое"
    ElseIf amplPct > 10# Then
        AmplitudeTag = "Избыточное"
    Else
        AmplitudeTag = "Норма"
    End If
End Function

Private Sub LoadKdProfiles(ByVal wsCfg As Worksheet)
    Dim r As Long, lastR As Long
    Dim et As String
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    Dim kd0 As Double, kd1 As Double, kd2 As Double, kd3 As Double
    Set m_KdProfiles = CreateObject("Scripting.Dictionary")
    lastR = LastUsedRow(wsCfg)
    For r = 3 To lastR
        et = UCase$(Trim$(CStr(wsCfg.Cells(r, 27).Value)))
        If Len(et) = 0 Then GoTo NX
        t0 = NzD(wsCfg.Cells(r, 28).Value, 0)
        kd0 = NzD(wsCfg.Cells(r, 29).Value, -1)
        t1 = NzD(wsCfg.Cells(r, 30).Value, 0)
        kd1 = NzD(wsCfg.Cells(r, 31).Value, -1)
        t2 = NzD(wsCfg.Cells(r, 32).Value, 0)
        kd2 = NzD(wsCfg.Cells(r, 33).Value, -1)
        t3 = NzD(wsCfg.Cells(r, 34).Value, 0)
        kd3 = NzD(wsCfg.Cells(r, 35).Value, -1)
        m_KdProfiles(et) = Array(t0, kd0, t1, kd1, t2, kd2, t3, kd3)
NX:
    Next r
End Sub

Private Function EvalKdAbs(ByVal tSec As Double, ByVal prof As Variant) As Double
    Dim t(0 To 3) As Double, k(0 To 3) As Double
    Dim i As Long, n As Long, j As Long
    n = 0
    For i = 0 To 3
        If CDbl(prof(2 * i + 1)) >= 0 Then
            t(n) = CDbl(prof(2 * i))
            k(n) = CDbl(prof(2 * i + 1))
            n = n + 1
        End If
    Next i
    If n < 2 Then
        Err.Raise vbObjectError + 2403, , "Профиль Kд(t) должен содержать минимум 2 точки (AA:AI)."
    End If
    For j = 1 To n - 1
        If tSec <= t(j) Then
            EvalKdAbs = k(j - 1) + (k(j) - k(j - 1)) * SafeDiv((tSec - t(j - 1)), (t(j) - t(j - 1)), 0)
            Exit Function
        End If
    Next j
    EvalKdAbs = k(n - 1)
End Function

Private Function ClampKd(ByVal kdVal As Double) As Double
    ClampKd = kdVal
    If ClampKd < 0.1 Then ClampKd = 0.1
    If ClampKd > 1# Then ClampKd = 1#
End Function

Private Function DynamicKdByTime(ByVal equipType As String, ByVal tSec As Double) As Double
    Dim prof As Variant, key As String
    key = UCase$(Trim$(equipType))
    If m_KdProfiles Is Nothing Then Err.Raise vbObjectError + 2401, , "Профили Kд(t) не загружены."
    If Not m_KdProfiles.Exists(key) Then
        Err.Raise vbObjectError + 2402, , "Для типа '" & equipType & "' не задан профиль Kд(t) в Config (AA:AI)."
    End If
    prof = m_KdProfiles(key)
    DynamicKdByTime = ClampKd(EvalKdAbs(tSec, prof))
End Function

Private Function KdProfileText(ByVal equipType As String) As String
    Dim prof As Variant, key As String, src As String
    key = UCase$(Trim$(equipType))
    src = "Config"
    If m_KdProfiles Is Nothing Then
        KdProfileText = "Config: профили Kд(t) не загружены"
        Exit Function
    End If
    If Not m_KdProfiles.Exists(key) Then
        KdProfileText = "Config: профиль для типа '" & equipType & "' не задан"
        Exit Function
    End If
    prof = m_KdProfiles(key)
    KdProfileText = src & ": t0=" & Format(prof(0), "0") & "с Kд0=" & Format(prof(1), "0.00") & _
                    "; t1=" & Format(prof(2), "0") & "с Kд1=" & Format(prof(3), "0.00") & _
                    "; t2=" & Format(prof(4), "0") & "с Kд2=" & Format(prof(5), "0.00") & _
                    "; t3=" & Format(prof(6), "0") & "с Kд3=" & Format(prof(7), "0.00")
End Function

Private Sub LoadChartDisplayOptions(ByVal wsCfg As Worksheet)
    Dim r As Long
    Set m_ChartOpts = CreateObject("Scripting.Dictionary")
    For r = 3 To 10
        m_ChartOpts("REL_" & CStr(r)) = (NzD(wsCfg.Cells(r, 37).Value, 1) <> 0) ' AK
        m_ChartOpts("ABS_" & CStr(r)) = (NzD(wsCfg.Cells(r, 38).Value, 1) <> 0) ' AL
    Next r
    ' Row mapping: 3 Pfact, 4 Preq, 5 Tol+, 6 Tol-, 7 Limits, 8 Freq, 9 Markers
    m_ChartOpts("REL_PFACT") = m_ChartOpts("REL_3")
    m_ChartOpts("REL_PREQ") = m_ChartOpts("REL_4")
    m_ChartOpts("REL_TOL_P") = m_ChartOpts("REL_5")
    m_ChartOpts("REL_TOL_M") = m_ChartOpts("REL_6")
    m_ChartOpts("REL_LIMITS") = m_ChartOpts("REL_7")
    m_ChartOpts("REL_FREQ") = m_ChartOpts("REL_8")
    m_ChartOpts("REL_MARKERS") = m_ChartOpts("REL_9")
    m_ChartOpts("ABS_PFACT") = m_ChartOpts("ABS_3")
    m_ChartOpts("ABS_PREQ") = m_ChartOpts("ABS_4")
    m_ChartOpts("ABS_TOL_P") = m_ChartOpts("ABS_5")
    m_ChartOpts("ABS_TOL_M") = m_ChartOpts("ABS_6")
    m_ChartOpts("ABS_LIMITS") = m_ChartOpts("ABS_7")
    m_ChartOpts("ABS_FREQ") = m_ChartOpts("ABS_8")
    m_ChartOpts("ABS_MARKERS") = m_ChartOpts("ABS_9")
End Sub

Private Function ChartOpt(ByVal key As String, ByVal fallback As Boolean) As Boolean
    If m_ChartOpts Is Nothing Then
        ChartOpt = fallback
    ElseIf m_ChartOpts.Exists(key) Then
        ChartOpt = CBool(m_ChartOpts(key))
    Else
        ChartOpt = fallback
    End If
End Function

Private Sub EnsureChartOptionCheckboxes(ByVal ws As Worksheet)
    Dim r As Long
    Dim cb As CheckBox
    Dim nm As String, c As Long
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If Left$(cb.Name, 7) = "chkOpt_" Then cb.Delete
    Next cb
    On Error GoTo 0

    For r = 3 To 9
        For c = 37 To 38 ' AK, AL
            nm = "chkOpt_" & r & "_" & c
            Set cb = ws.CheckBoxes.Add(ws.Cells(r, c).Left + 2, ws.Cells(r, c).Top + 2, ws.Cells(r, c).Width - 4, ws.Cells(r, c).Height - 4)
            cb.Name = nm
            cb.Caption = ""
            cb.LinkedCell = ws.Cells(r, c).Address(True, True, xlA1, True)
            cb.Value = IIf(NzD(ws.Cells(r, c).Value, 1) <> 0, xlOn, xlOff)
            cb.Placement = xlMoveAndSize
        Next c
    Next r
End Sub

' ==========================================================
' Служебные
' ==========================================================

Private Sub CollectOldOutputSheets(ByRef names As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_RAW And ws.Name <> SH_CFG And ws.Name <> SH_SUM And ws.Name <> SH_LOG Then
            names.Add ws.Name
        End If
    Next ws
End Sub

Private Sub DeleteOutputSheets(ByVal names As Collection)
    Dim i As Long
    Application.DisplayAlerts = False
    For i = names.Count To 1 Step -1
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(names(i))).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
End Sub

Private Function StationMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim na As String, nb As String
    na = UCase$(Trim$(a))
    nb = UCase$(Trim$(b))
    If InStr(na, "ВОРКУТ") > 0 And InStr(nb, "ВОРКУТ") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СОСНОГОР") > 0 And InStr(nb, "СОСНОГОР") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СЛПК") > 0 And InStr(nb, "СЛПК") > 0 Then StationMatch = True: Exit Function
    StationMatch = (na = nb)
End Function

Private Function MinArray(ByRef arr() As Double) As Double
    Dim i As Long, v As Double
    v = 1E+99
    For i = LBound(arr) To UBound(arr)
        If arr(i) > 0 And arr(i) < v Then v = arr(i)
    Next i
    If v = 1E+99 Then v = 0.05
    MinArray = v
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxD = a Else MaxD = b
End Function

Private Function FormatSecOrNA(ByVal secVal As Double) As String
    If secVal < 0 Then FormatSecOrNA = "н/д" Else FormatSecOrNA = Format(secVal, "0.0")
End Function

Private Function GeneratorSheetName(ByRef g As TGenCfg) As String
    GeneratorSheetName = MakeSheetName(g.Station & "_" & g.Generator)
End Function

Private Function GeneratorChartSheetName(ByRef g As TGenCfg) As String
    Dim base As String, full As String
    base = g.Station & "_" & g.Generator
    full = base & CHART_SUFFIX
    If Len(full) > 31 Then
        base = Left$(base, 31 - Len(CHART_SUFFIX))
        full = base & CHART_SUFFIX
    End If
    GeneratorChartSheetName = MakeSheetName(full)
End Function

Private Function MakeSheetName(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, ":", "_")
    out = Replace(out, "\", "_")
    out = Replace(out, "/", "_")
    out = Replace(out, "?", "_")
    out = Replace(out, "*", "_")
    out = Replace(out, "[", "_")
    out = Replace(out, "]", "_")
    If Len(out) > 31 Then out = Left$(out, 31)
    If Len(out) = 0 Then out = "Sheet_OPRCH"
    MakeSheetName = out
End Function

Private Function GetRequiredSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetRequiredSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetRequiredSheet Is Nothing Then
        Err.Raise vbObjectError + 2999, , "Не найден лист '" & name & "'. Запустите SetupOPRCHTemplate."
    End If
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, v As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(v, headerName, vbTextCompare) = 0 Then FindHeaderCol = c: Exit Function
    Next c
    FindHeaderCol = 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Private Function NzD(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        NzD = fallback
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = fallback
    End If
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double, ByVal fallback As Double) As Double
    If Abs(b) < 0.0000001 Then SafeDiv = fallback Else SafeDiv = a / b
End Function

Private Function SgnNZ(ByVal v As Double) As Integer
    If v > 0.000001 Then
        SgnNZ = 1
    ElseIf v < -0.000001 Then
        SgnNZ = -1
    Else
        SgnNZ = 0
    End If
End Function

' ==========================================================
' Кнопки управления
' ==========================================================

Private Sub EnsureControlButtons(ByVal ws As Worksheet)
    AddOrReplaceButton ws, "btnRunOPRCH", ws.Range("E30"), 260, 32, _
        "Запустить мониторинг ОПРЧ", "AnalyzeOPRCH", RGB(40, 120, 220), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnClearOPRCH", ws.Range("I30"), 200, 32, _
        "Очистить результаты", "ClearOPRCHResults", RGB(150, 150, 150), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnPresets", ws.Range("M30"), 220, 32, _
        "Применить пресеты типов", "ApplyPresetsToConfig", RGB(80, 160, 100), RGB(255, 255, 255)
End Sub

Private Sub AddOrReplaceButton(ByVal ws As Worksheet, ByVal nm As String, ByVal anchor As Range, _
                               ByVal w As Double, ByVal h As Double, _
                               ByVal caption As String, ByVal onAction As String, _
                               ByVal fillColor As Long, ByVal textColor As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, w, h)
    shp.Name = nm
    shp.TextFrame2.TextRange.Text = caption
    shp.OnAction = onAction
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(70, 70, 70)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
    shp.TextFrame2.TextRange.Font.Size = 11
End Sub
