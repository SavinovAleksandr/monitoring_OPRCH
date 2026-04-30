Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' ==========================================================
' OPRCH monitoring  (quantitative + qualitative)
' Per-generator:  data sheet <Ст>_<Ген>        +  chart sheet <Ст>_<Ген>_Граф
' Per-station:    data sheet <Ст>_Сумма[_<Пп>] +  chart sheet <...>_Граф
' Extra sheets:   Summary, Log
'
' Версия: 1.4.0
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
' ==========================================================

Public Const OPRCH_VERSION As String = "1.4.0"

Private Const SH_RAW As String = "RawData"
Private Const SH_CFG As String = "Config"
Private Const SH_SUM As String = "Summary"
Private Const SH_LOG As String = "Log"
Private Const CHART_SUFFIX As String = "_Граф"

Private m_LogRow As Long
Private m_KdProfiles As Object   ' key=EQUIPTYPE, value=Array(t0,kd0,t1,kd1,t2,kd2,t3,kd3)

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

    wsCfg.Range("AA1").Value = "Профили Kд(t) (абсолютные значения Kд)"
    wsCfg.Range("AA2:AI2").Value = Array("Тип_оборудования", "t0, с", "Kд0", "t1, с", "Kд1", "t2, с", "Kд2", "t3, с", "Kд3")
    wsCfg.Cells(3, 27).Resize(1, 9).Value = Array("ПТУ_блок", 0, 1, 30, 0.5, 240, 0.8, 600, 1)
    wsCfg.Cells(4, 27).Resize(1, 9).Value = Array("ПТУ_неблок", 0, 1, 30, 0.5, 240, 0.8, 600, 1)
    wsCfg.Cells(5, 27).Resize(1, 9).Value = Array("ГТУ", 0, 1, 30, 0.9, 240, 0.8, 600, 1)
    wsCfg.Cells(6, 27).Resize(1, 9).Value = Array("ПГУ_утил", 0, 1, 30, 0.75, 240, 0.7, 900, 1)
    wsCfg.Cells(7, 27).Resize(1, 9).Value = Array("ПГУ_сбросн", 0, 1, 30, 0.7, 240, 0.65, 1200, 1)
    wsCfg.Cells(8, 27).Resize(1, 9).Value = Array("ГПА", 0, 1, 30, 0.9, 240, 0.85, 600, 1)

    wsCfg.Columns("A:AI").AutoFit
    wsRaw.Columns("A:E").AutoFit
    EnsureControlButtons wsCfg

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
    wsSummary.Range("A1:AI1").Value = Array( _
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
        "Примечание" _
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
    BuildStationAggregates wsRaw, wsCfg, wsSummary, st

    stepName = "Оформление Summary"
    wsSummary.Columns("A:AI").AutoFit
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
    p0 = NzD(wsRaw.Cells(startRow, pCol).Value, 0)
    ptek = NzD(wsRaw.Cells(endQ, pCol).Value, 0)
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

    calcStep = "Характер переходного процесса"
    EvaluateTransient wsRaw, st, g, res, pCol, timeCol

    calcStep = "Качественный критерий"
    EvaluateQualitative wsRaw, st, g, res, pCol, fCol, timeCol

    AnalyzeOneGenerator = res
    Exit Function

EH:
    Err.Raise vbObjectError + 2199, , "AnalyzeOneGenerator (" & g.Generator & " / " & calcStep & "): " & Err.Description
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
    ' Корректная методика для мониторинга реальных событий:
    ' цели масштабируются к фактическому |Pтреб|, а dP5/dP10 задают пропорции
    ' (например, 5%/10% = 0.5 = половина Pтреб к моменту t5, полный Pтреб к t10).
    ' Для контрольных испытаний (ступень 10 % Pном) это даёт тот же результат,
    ' т. к. Pтреб = Pном * dp10% / 100.
    Dim pReqAbs As Double, ratio5 As Double
    pReqAbs = Abs(res.PReq)
    If g.Dp10Pct > 0 Then
        ratio5 = g.Dp5Pct / g.Dp10Pct
    Else
        ratio5 = 0.5
    End If
    target5 = signReq * pReqAbs * ratio5
    target10 = signReq * pReqAbs

    row5 = 0
    row10 = 0
    hit5 = False
    hit10 = False

    qStep = "Поиск t5/t10"
    For r = res.StartRow To res.EndQualRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        If Not hit5 Then
            If signReq * dP >= signReq * target5 Then
                hit5 = True
                row5 = r
            End If
        End If
        If Not hit10 Then
            If signReq * dP >= signReq * target10 Then
                hit10 = True
                row10 = r
            End If
        End If
    Next r

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

    res.QualT5Pass = (hit5 And t5 <= g.T5Sec)
    res.QualT10Pass = (hit10 And t10 <= g.T10Sec)
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

    For r = displayStartRow To endRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        dFr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, g.Fnch)

        ws.Cells(outR, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(outR, 2).Value = wsRaw.Cells(r, fCol).Value
        ws.Cells(outR, 3).Value = wsRaw.Cells(r, pCol).Value
        ws.Cells(outR, 4).Value = dP
        ws.Cells(outR, 5).Value = -100# / g.SPct * g.PNom / st.FNom * DynamicKdByTime(g.EquipType, _
                              SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)) * dFr
        ws.Cells(outR, 6).Value = dFr
        ws.Cells(outR, 8).Value = targetPreq + tolPreq
        ws.Cells(outR, 9).Value = targetPreq - tolPreq
        ws.Cells(outR, 13).Value = dPmaxRel
        ws.Cells(outR, 14).Value = dPminRel
        ws.Cells(outR, 15).Value = res.P0 + targetPreq
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
    ws.Range("G1:G8").NumberFormat = "@"
    ws.Range("H1:H5").NumberFormat = "0.000"
    ws.Range("H6").NumberFormat = "@"
    ws.Range("H7").NumberFormat = "0.000"
    ws.Range("H8").NumberFormat = "@"
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

    Set chartObj = wsChart.ChartObjects.Add(10, 140, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ: " & g.Station & " / " & g.Generator
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlTimeScale
        On Error GoTo 0
    End With

    ' Основные ряды
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 4, "Pфакт, МВт", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 5, "Pтреб, МВт", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 8, "+Допуск уст.", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 9, "-Допуск уст.", False
    AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 13, "Pmax (dPmax)"
    AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 14, "Pmin (dPmin)"
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True

    ' Вертикальные маркеры t5 / t10 / выхода за fнч - через отдельные "точечные" ряды
    AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 10, "t5"
    AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 11, "t10"
    AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 12, "Выход за fнч"

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
    Set chartObj = wsChart.ChartObjects.Add(10, 520, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ (абсолютные): " & g.Station & " / " & g.Generator
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlTimeScale
        On Error GoTo 0
    End With

    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 3, "Pфакт, МВт", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 15, "Pтреб, МВт (абс)", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 16, "+Допуск уст., МВт (абс)", False
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 17, "-Допуск уст., МВт (абс)", False
    AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 18, "Pmax, МВт"
    AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 19, "Pmin, МВт"
    AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True
    ' Маркеры для абсолютного графика отключены, чтобы избежать проблем компиляции в части локалей VBA.

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

    wsChart.Range("A3:A8").Font.Bold = True
    wsChart.Range("D3:D7").Font.Bold = True
    wsChart.Range("A3:E8").NumberFormat = "@"
    If res.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
        wsChart.Range("B3").Interior.Color = RGB(255, 235, 156)
        wsChart.Range("E7").Font.Color = RGB(156, 87, 0)
    ElseIf Not res.QuantPass Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If Not res.QualPass Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If res.TransientType = "Колебательный" Then wsChart.Range("B5").Font.Color = RGB(192, 0, 0)
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

Private Sub BuildStationAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, ByRef st As TSettings)
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
        BuildOneStationAggregate wsRaw, wsCfg, CStr(key), "", st
        BuildParoprovodSubAggregates wsRaw, wsCfg, CStr(key), st
    Next key
End Sub

Private Sub BuildParoprovodSubAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal stationName As String, ByRef st As TSettings)
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
            BuildOneStationAggregate wsRaw, wsCfg, stationName, CStr(p), st
        End If
    Next p
End Sub

Private Sub BuildOneStationAggregate(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings)
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
    Dim pMaxSum As Double, pMinSum As Double
    Dim reservePlus As Double, reserveMinus As Double
    Dim limited As Boolean, limitType As String

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

    startRow = ResolveStartRow(wsRaw, timeCol, freqCol, st, MinArray(fnchArr), firstExceedRow)
    endRow = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)

    p0 = 0
    pMaxSum = 0
    pMinSum = 0
    For i = 1 To cnt
        If NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0) > st.WorkThresholdMW Then
            p0 = p0 + NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0)
            pMaxSum = pMaxSum + pmaxArr(i)
            pMinSum = pMinSum + pminArr(i)
        End If
    Next i

    pNow = 0
    For i = 1 To cnt
        If NzD(wsRaw.Cells(endRow, pCols(i)).Value, 0) > st.WorkThresholdMW Then
            pNow = pNow + NzD(wsRaw.Cells(endRow, pCols(i)).Value, 0)
        End If
    Next i

    dF = MaxAbsDeviationInWindow(wsRaw, freqCol, startRow, endRow, st.FNom)
    preqOrig = 0
    Dim tSecAgg As Double
    For i = 1 To cnt
        dfr = DeadbandDeviation(dF, fnchArr(i))
        tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(endRow, timeCol).Value)
        preqOrig = preqOrig + (-100# / sArr(i) * pnomArr(i) / st.FNom * DynamicKdByTime(etArr(i), tSecAgg) * dfr)
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
        AppendLog "INFO", stationName & IIf(Len(paropipeFilter) > 0, "/" & paropipeFilter, ""), _
                  "Pтреб_сум ограничен " & limitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

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
        ws.Range("A2:B2").Value = Array("Режим", "Суммарная нагрузка включённых генераторов")
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

    ws.Range("A10:G10").Value = Array("Время", "Частота, Гц", "Pсум, МВт", "dPсум, МВт", "Pтреб_сум, МВт", _
                                       "dPmax_сум", "dPmin_сум")
    rowQ = 11
    Dim dPmaxRel As Double, dPminRel As Double
    dPmaxRel = pMaxSum - p0
    dPminRel = pMinSum - p0
    For r = startRow To endRow
        ws.Cells(rowQ, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(rowQ, 2).Value = wsRaw.Cells(r, freqCol).Value

        pNow = 0
        For i = 1 To cnt
            If NzD(wsRaw.Cells(r, pCols(i)).Value, 0) > st.WorkThresholdMW Then
                pNow = pNow + NzD(wsRaw.Cells(r, pCols(i)).Value, 0)
            End If
        Next i

        ws.Cells(rowQ, 3).Value = pNow
        ws.Cells(rowQ, 4).Value = pNow - p0

        Dim preqStep As Double
        preqStep = 0
        For i = 1 To cnt
            dfr = DeadbandDeviation(NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom, fnchArr(i))
            tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
            preqStep = preqStep + (-100# / sArr(i) * pnomArr(i) / st.FNom * DynamicKdByTime(etArr(i), tSecAgg) * dfr)
        Next i
        ws.Cells(rowQ, 5).Value = preqStep
        ws.Cells(rowQ, 6).Value = dPmaxRel
        ws.Cells(rowQ, 7).Value = dPminRel
        rowQ = rowQ + 1
    Next r

    ws.Columns("A:G").AutoFit
    ws.Range("A1:A8").NumberFormat = "@"
    ws.Range("B1:B2").NumberFormat = "@"
    ws.Range("B3:B4").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B5:B8").NumberFormat = "0.000"
    ws.Range("D1:D7").NumberFormat = "@"
    ws.Range("E1:E5").NumberFormat = "0.000"
    ws.Range("E6").NumberFormat = "@"
    ws.Range("E7").NumberFormat = "0"
    ws.Range("A10:G10").NumberFormat = "@"
    ws.Range("A11:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B11:G100000").NumberFormat = "0.000"

    WriteStationChartSheet stationName, paropipeFilter, rowQ - 1, st, MinArray(fnchArr), dF
End Sub

Private Sub WriteStationChartSheet(ByVal stationName As String, ByVal paropipeFilter As String, _
                                   ByVal lastDataRow As Long, ByRef st As TSettings, _
                                   ByVal fnchMin As Double, ByVal dFmax As Double)
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
    chartName = MakeSheetName(stationName & suffix & CHART_SUFFIX)
    If StrComp(chartName, dataName, vbTextCompare) = 0 Then
        chartName = Left$(dataName, 28) & "_Гр"
    End If

    Set wsData = ThisWorkbook.Worksheets(dataName)
    Set wsChart = EnsureSheet(chartName)
    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    wsChart.Range("A1").Value = "График ОПРЧ (сумма): " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "")
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    startRow = 11
    If lastDataRow < startRow Then Exit Sub
    endRow = FindChartEndRow(wsData, startRow, lastDataRow, wsData.Cells(startRow, 1).Value, st.ChartIntervalSec)
    If endRow < startRow Then endRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 40, 1020, 520)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Суммарный мониторинг ОПРЧ: " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "")
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlTimeScale
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

    ws.Cells(r, 1).Resize(1, 35).Value = Array( _
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
        note _
    )
End Sub

Private Sub WriteSummaryInvalid(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg)
    ws.Cells(r, 1).Resize(1, 35).Value = Array( _
        g.Station, g.Generator, g.EquipType, "", "", "", "", "", "", "", "", "", "", "", "", _
        "Нарушение", "", "", 0, "Н/Д", "Н/Д", "Н/Д", "Н/Д", "", "", "", "Config", "", _
        "", "", "", "", "", "", _
        "Не заполнен обязательный параметр config" _
    )
End Sub

Private Sub ApplySummaryConditionalFormat(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ws.Range("A2:AI" & lastRow).Interior.Pattern = xlNone

    ' Колонки текстовых статусов: P, Q, T, U, V, W (колич/перерег/кач/t5/t10/уст)
    ApplyStatusCF ws, ws.Range("P2:P" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("T2:T" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("U2:U" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("V2:V" & lastRow), "ОК", "Нарушение"
    
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

    ws.Range("A1:AI1").Font.Bold = True
    ws.Range("A1:AI1").Interior.Color = RGB(217, 225, 242)
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
    col = 37   ' колонка AK - правее данных Summary (после AI)
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
    g.T5Sec = 15
    g.Dp5Pct = 5
    g.T10Sec = 420
    g.Dp10Pct = 10
    g.SteadyTolPct = 1
    g.Fnch = 0.075

    If InStr(et, "ГПА") > 0 Then
        g.T10Sec = 120
    ElseIf InStr(et, "ГТУ") > 0 Then
        g.T10Sec = 900
    ElseIf InStr(et, "ПГУ_СБРОСН") > 0 Or InStr(et, "СБРОСН") > 0 Then
        g.T10Sec = 2100
    ElseIf InStr(et, "ПГУ_УТИЛ") > 0 Or InStr(et, "УТИЛИЗ") > 0 Then
        g.T10Sec = 900
    ElseIf InStr(et, "ПТУ_БЛОК") > 0 Then
        g.T10Sec = 360
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
