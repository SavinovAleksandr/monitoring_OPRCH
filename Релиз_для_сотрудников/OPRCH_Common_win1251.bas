Attribute VB_Name = "OPRCH_Common"
Option Explicit



Public Const OPRCH_VERSION As String = "1.8.2"

Public Const SH_RAW As String = "RawData"
Public Const SH_CFG As String = "Config"
Public Const SH_SUM As String = "Summary"
Public Const SH_LOG As String = "Log"
Public Const SH_PROT As String = "Протокол"
Public Const SH_INSTR As String = "Инструкция"
Public Const CHART_SUFFIX As String = "_Граф"

Public Const QUANT_DF_HZ As Double = 0.2          ' п. 4.1.2: колич. оценка при |dF| >= 0.2 Гц
Public Const QUANT_MIN_SEC_TES As Double = 30#    ' п. 4.1.2: длительность >= 30 с (ТЭС)
Public Const QUANT_MIN_SEC_GPA As Double = 60#    ' п. 4.1.2: длительность >= 1 мин (ГПА/ДГУ/ГЭС)
Public Const QUAL_JUMP_SEC As Double = 10#        ' п. 4.1.3: резкое отклонение - за время <= 10 с
Public Const QUAL_MIN_SEC As Double = 20#         ' п. 4.1.3: продолжительность >= 20 с
Public Const QUAL_RM_DF_HZ As Double = 0.15       ' п. 4.4.4: с РМ - выход за 50±0.15 Гц
Public Const QUAL_NORM_DF_MAX_HZ As Double = 0.4  ' п. 4.4.3: без РМ - скачок ±0.2..0.4 Гц
Public Const NO_RM_S_PCT As Double = 6#           ' без РМ в режиме «Авария»: S = 6 %
Public Const NO_RM_FNCH_HZ As Double = 0.15       ' без РМ в режиме «Авария»: fнч = 0.15 Гц
Public Const QS_HALF_WINDOW_SEC As Double = 15#   ' Приложение: квазиуст. режим - отрезок ±15 с
Public Const QS_TOL_HZ As Double = 0.01           ' Приложение: |f - fср| <= 10 мГц
Public Const OSC_MIN_PERIODS As Long = 5          ' п. 4.2.6/4.2.14.4: колебания > 5 периодов
Public Const OSC_MIN_AMPL_PCT As Double = 1#      ' п. 4.2.14.4: амплитуда > 1 %Pном
Public Const TELEM_MAX_REPEAT As Long = 11        ' Прил. 1 п. 6.1.8.3: не более 11 одинаковых подряд
Public Const H15_REF As Double = 0.3              ' ориентир для h(15 с) (Приложение п. 6.2, справочно)
Public Const EVENT_MIN_OUTSIDE_SEC As Double = 5# ' фильтр коротких выбросов частоты при поиске событий

Public m_LogRow As Long
Public m_KdProfiles As Object   ' key=EQUIPTYPE, value=Array(t0,kd0,t1,kd1,t2,kd2,t3,kd3)
Public m_ChartOpts As Object    ' key=REL_*/ABS_*, value=Boolean
Public m_Protocol As Collection ' строки текстового протокола (лист «Протокол»)

Public Type TSettings
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
    Mode As String            ' «Авария» / «Испытание»
    IsEmergency As Boolean    ' True = «Авария»
    MaxEventSec As Double     ' макс. длительность окна одного события, с
    MaxEvents As Long         ' макс. число событий в выборке (1 = только первое)
    MaxReqPct As Double       ' ограничение Pтреб, %Pном (Приложение: 10 %; 0 = без ограничения)
    StrictDeficit As Boolean  ' True: дефицит «хотя бы в одном периоде» = НАРУШЕНИЕ (п. 4.2.14.1)
End Type

Public Type TEvent
    StartRow As Long          ' последняя строка до выхода за fнч (t0)
    FirstExceedRow As Long    ' первая строка за fнч
    EndRow As Long            ' конец окна события
End Type

Public Type TGenCfg
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
    HasRM As Boolean        ' оснащён регулятором мощности (Config U; пусто = 1)
    Layout As String        ' «блок» / «неблок» (Config V; пусто = по типу оборудования)
    SPctCfg As Double       ' S из Config (до подстановки 6 % для оборудования без РМ)
    FnchCfg As Double       ' fнч из Config (до подстановки 0.15 Гц)
    ForcedNoRM As Boolean   ' True, если в режиме «Авария» применены S = 6 %, fнч = 0.15 Гц
End Type

Public Type TGenResult
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
    EventIdx As Long        ' номер события в выборке (1..N)
    EventEndRow As Long     ' конец окна события
    QuantApplicable As Boolean  ' п. 4.1.2
    QualApplicable As Boolean   ' п. 4.1.3 / 4.4.3 / 4.4.4
    Applicability As String     ' текст статуса применимости
    DfMaxEvent As Double        ' макс. отклонение частоты за событие (со знаком), Гц
    DurAbove02Sec As Double     ' длительность |dF| >= 0.2 Гц (непрерывная макс.), с
    JumpSec As Double           ' время нарастания отклонения до максимума, с
    DurOutsideSec As Double     ' длительность |dF| > fнч, с
    TrRow As Long               ' строка tр (п. 4.2.5: худшая квазиустановившаяся точка)
    TrTime As Variant
    TrSec As Double             ' tр - t0, с
    TrQuasiSteady As Boolean    ' tр в квазиустановившемся режиме по частоте
    PReqRef As Double           ' Pтреб на контрольном интервале (Config W5), МВт
    PFactRef As Double          ' Pфакт на контрольном интервале, МВт
    OscViolation As Boolean     ' п. 4.2.14.4
    OscPeriods As Long
    OscAmplPct As Double
    OscNote As String
    TelemetryOK As Boolean
    TelemetryNote As String
    ReactionType As String      ' п. 4.4.1: «1 - адекватная» ... «8 - телеизмерение»
    ReactionCode As Long
    QualClauseVerdict As String ' вывод по п. 4.4.3/4.4.4
    H15 As Double               ' оценка переходной функции h(15 с); -999 = не вычислено
    H15Note As String
    F0 As Double                ' опорная частота: fном (с РМ) или f(t0) (без РМ), п. 4.2.4.1
End Type

Public Type TStationResult
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
    EventIdx As Long
    IsNeblok As Boolean         ' в сумме есть генераторы неблочной компоновки (п. 4.2.10 применим)
    QuantApplicable As Boolean
    Applicability As String
    TrTime As Variant
    TrSec As Double
    TrQuasiSteady As Boolean
    F0 As Double
End Type


Public Function PReqCalc(ByVal sPct As Double, ByVal pNom As Double, ByVal fNom As Double, _
                          ByVal kd As Double, ByVal dfr As Double, ByVal maxReqPct As Double) As Double
    Dim v As Double, capMW As Double
    If dfr = 0 Or sPct <= 0 Or fNom <= 0 Then PReqCalc = 0: Exit Function
    v = -100# / sPct * pNom / fNom * kd * dfr
    If maxReqPct > 0 Then
        capMW = maxReqPct / 100# * pNom
        If v > capMW Then v = capMW
        If v < -capMW Then v = -capMW
    End If
    PReqCalc = v
End Function

Public Function ClampReserve(ByVal v As Double, ByVal reservePlus As Double, ByVal reserveMinus As Double) As Double
    If v > reservePlus Then v = reservePlus
    If v < -reserveMinus Then v = -reserveMinus
    ClampReserve = v
End Function


Public Function ReadGenCfgEff(ByVal ws As Worksheet, ByVal r As Long, ByRef st As TSettings) As TGenCfg
    Dim g As TGenCfg
    g = ReadGenCfg(ws, r)
    g.SPctCfg = g.SPct
    g.FnchCfg = g.Fnch
    g.ForcedNoRM = False
    If st.IsEmergency And Not g.HasRM Then
        g.SPct = NO_RM_S_PCT
        g.Fnch = NO_RM_FNCH_HZ
        g.ForcedNoRM = True
    End If
    ReadGenCfgEff = g
End Function


Public Sub WindowRowsAround(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal row As Long, _
                             ByVal halfSec As Double, ByRef r1 As Long, ByRef r2 As Long)
    Dim lastR As Long, t0 As Double, tv As Variant
    lastR = LastUsedRow(wsRaw)
    r1 = row: r2 = row
    If Not IsDate(wsRaw.Cells(row, timeCol).Value) Then Exit Sub
    t0 = CDbl(CDate(wsRaw.Cells(row, timeCol).Value)) * 86400#
    Do While r1 > 2
        tv = wsRaw.Cells(r1 - 1, timeCol).Value
        If Not IsDate(tv) Then Exit Do
        If t0 - CDbl(CDate(tv)) * 86400# > halfSec + 0.0001 Then Exit Do
        r1 = r1 - 1
    Loop
    Do While r2 < lastR
        tv = wsRaw.Cells(r2 + 1, timeCol).Value
        If Not IsDate(tv) Then Exit Do
        If CDbl(CDate(tv)) * 86400# - t0 > halfSec + 0.0001 Then Exit Do
        r2 = r2 + 1
    Loop
End Sub

Public Function MeanFreqAround(ByVal wsRaw As Worksheet, ByVal fCol As Long, ByVal timeCol As Long, _
                                ByVal row As Long, ByVal halfSec As Double, ByVal fNom As Double) As Double
    Dim r1 As Long, r2 As Long, r As Long, s As Double, n As Long
    WindowRowsAround wsRaw, timeCol, row, halfSec, r1, r2
    For r = r1 To r2
        If IsNumeric(wsRaw.Cells(r, fCol).Value) And Len(Trim$(CStr(wsRaw.Cells(r, fCol).Value))) > 0 Then
            s = s + CDbl(wsRaw.Cells(r, fCol).Value)
            n = n + 1
        End If
    Next r
    If n > 0 Then MeanFreqAround = s / n Else MeanFreqAround = fNom
End Function


Public Function IsQuasiSteady(ByVal wsRaw As Worksheet, ByVal fCol As Long, ByVal timeCol As Long, _
                               ByVal row As Long, ByVal fNom As Double, ByRef fMean As Double) As Boolean
    Dim r1 As Long, r2 As Long, r As Long, v As Variant
    fMean = MeanFreqAround(wsRaw, fCol, timeCol, row, QS_HALF_WINDOW_SEC, fNom)
    WindowRowsAround wsRaw, timeCol, row, QS_HALF_WINDOW_SEC, r1, r2
    For r = r1 To r2
        v = wsRaw.Cells(r, fCol).Value
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            If Abs(CDbl(v) - fMean) > QS_TOL_HZ + 0.0000001 Then IsQuasiSteady = False: Exit Function
        End If
    Next r
    IsQuasiSteady = True
End Function


Public Function IdentifyH15(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                             ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                             ByVal centerRow As Long, ByVal p0 As Double, ByVal f0 As Double, ByRef note As String) As Double
    Const NPTS As Long = 50      ' t-17 .. t+32 (46 точек + по 2 для сглаживания)
    Const NH As Long = 15
    Dim x() As Double, y() As Double, xs() As Double, ys() As Double, dx() As Double, dy() As Double
    Dim a() As Double, b() As Double, dh() As Double
    Dim i As Long, j As Long, k As Long, r As Long, lastR As Long
    Dim dfr As Double, s As Double, c As Long, sumAbsDx As Double, h As Double, tr As Double
    ReDim x(1 To NPTS): ReDim y(1 To NPTS): ReDim xs(1 To NPTS): ReDim ys(1 To NPTS)
    ReDim dx(1 To NPTS): ReDim dy(1 To NPTS)
    ReDim a(1 To NH, 1 To NH): ReDim b(1 To NH): ReDim dh(1 To NH)

    IdentifyH15 = -999
    note = ""
    If centerRow <= 0 Or g.PNom <= 0 Or g.SPct <= 0 Then note = "h(15 с): нет момента выхода за fнч": Exit Function
    lastR = LastUsedRow(wsRaw)
    If RowByTimeOffset(wsRaw, timeCol, centerRow, 32#) >= lastR Then note = "h(15 с): недостаточно данных после выхода за fнч": Exit Function

    For i = 1 To NPTS
        r = RowByTimeOffset(wsRaw, timeCol, centerRow, CDbl(i - 18))   ' i=18 -> t
        If r < 2 Then r = 2
        dfr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, f0) - f0, g.Fnch)
        x(i) = PReqCalc(g.SPct, g.PNom, st.FNom, 1#, dfr, st.MaxReqPct) / g.PNom * 100#
        y(i) = (NzD(wsRaw.Cells(r, pCol).Value, p0) - p0) / g.PNom * 100#
    Next i
    For i = 1 To NPTS
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= NPTS Then s = s + x(j): c = c + 1
        Next j
        xs(i) = s / c
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= NPTS Then s = s + y(j): c = c + 1
        Next j
        ys(i) = s / c
    Next i
    dx(3) = 0: dy(3) = 0
    sumAbsDx = 0
    For i = 4 To 48
        dx(i) = xs(i) - xs(i - 1)
        dy(i) = ys(i) - ys(i - 1)
        sumAbsDx = sumAbsDx + Abs(dx(i))
    Next i
    If sumAbsDx < 0.05 Then note = "h(15 с): входное воздействие (Pтреб) практически не менялось - идентификация не выполнена": Exit Function

    For i = 1 To NH
        b(i) = 0
        For j = 1 To NH
            a(i, j) = 0
        Next j
    Next i
    For k = 3 + NH To 48
        For i = 1 To NH
            b(i) = b(i) + dx(k - i) * dy(k)
            For j = 1 To NH
                a(i, j) = a(i, j) + dx(k - i) * dx(k - j)
            Next j
        Next i
    Next k
    tr = 0
    For i = 1 To NH
        tr = tr + a(i, i)
    Next i
    For i = 1 To NH
        a(i, i) = a(i, i) + 0.000001 * (tr / NH + 0.000001)
    Next i
    If Not SolveLinearSystem(a, b, dh, NH) Then note = "h(15 с): СЛАУ вырождена": Exit Function
    h = 0
    For i = 1 To NH
        h = h + dh(i)
    Next i
    IdentifyH15 = h
    note = "h(15 с) = " & Format(h, "0.00") & " (Прил. п. 6.2, справочно; ориентир " & Format(H15_REF, "0.0") & ")"
End Function


Public Function SolveLinearSystem(ByRef a() As Double, ByRef b() As Double, ByRef x() As Double, ByVal n As Long) As Boolean
    Dim i As Long, j As Long, k As Long, piv As Long, tmp As Double, factor As Double
    For k = 1 To n
        piv = k
        For i = k + 1 To n
            If Abs(a(i, k)) > Abs(a(piv, k)) Then piv = i
        Next i
        If Abs(a(piv, k)) < 1E-14 Then SolveLinearSystem = False: Exit Function
        If piv <> k Then
            For j = 1 To n
                tmp = a(k, j): a(k, j) = a(piv, j): a(piv, j) = tmp
            Next j
            tmp = b(k): b(k) = b(piv): b(piv) = tmp
        End If
        For i = k + 1 To n
            factor = a(i, k) / a(k, k)
            For j = k To n
                a(i, j) = a(i, j) - factor * a(k, j)
            Next j
            b(i) = b(i) - factor * b(k)
        Next i
    Next k
    For i = n To 1 Step -1
        tmp = b(i)
        For j = i + 1 To n
            tmp = tmp - a(i, j) * x(j)
        Next j
        x(i) = tmp / a(i, i)
    Next i
    SolveLinearSystem = True
End Function


Public Function FindHeaderColFuzzy(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, v As String, target As String, found As Long
    FindHeaderColFuzzy = FindHeaderCol(ws, headerName)
    If FindHeaderColFuzzy > 0 Then Exit Function
    target = NormalizeHeader(headerName)
    If Len(target) = 0 Then Exit Function
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = NormalizeHeader(CStr(ws.Cells(1, c).Value))
        If v = target Then FindHeaderColFuzzy = c: Exit Function
    Next c
    found = 0
    For c = 1 To lastCol
        v = NormalizeHeader(CStr(ws.Cells(1, c).Value))
        If Len(v) > 0 Then
            If InStr(v, target) > 0 Or InStr(target, v) > 0 Then
                If found = 0 Then found = c Else found = -1: Exit For
            End If
        End If
    Next c
    If found > 0 Then
        FindHeaderColFuzzy = found
        AppendLog "INFO", "RawData", "Колонка '" & headerName & "' сопоставлена с '" & CStr(ws.Cells(1, found).Value) & "' (нестрогий поиск)"
    End If
End Function

Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, " ", ""): t = Replace(t, "_", ""): t = Replace(t, "-", ""): t = Replace(t, ".", ""): t = Replace(t, ",", "")
    t = Replace(t, "P", "Р"): t = Replace(t, "C", "С"): t = Replace(t, "T", "Т"): t = Replace(t, "B", "В")
    t = Replace(t, "H", "Н"): t = Replace(t, "K", "К"): t = Replace(t, "M", "М"): t = Replace(t, "O", "О")
    t = Replace(t, "A", "А"): t = Replace(t, "E", "Е"): t = Replace(t, "X", "Х"): t = Replace(t, "Y", "У")
    NormalizeHeader = t
End Function


Public Function ReadSettings(ByVal wsCfg As Worksheet) As TSettings
    Dim st As TSettings
    Dim valCol As Long, modeTxt As String
    valCol = 24
    If Trim$(CStr(wsCfg.Cells(2, 23).Value)) = "" Then
        Err.Raise vbObjectError + 2003, , "В Config не найдены глобальные настройки (W2:X16). Запустите SetupOPRCHTemplate или заполните колонки W/X."
    End If
    st.FNom = NzD(wsCfg.Cells(2, valCol).Value, 50#)
    st.AutoStart = (NzD(wsCfg.Cells(4, valCol).Value, 1) <> 0)
    st.QuantIntervalSec = NzD(wsCfg.Cells(5, valCol).Value, 300#)
    st.QuantTolPct = NzD(wsCfg.Cells(6, valCol).Value, 10#)
    st.WorkThresholdMW = NzD(wsCfg.Cells(7, valCol).Value, 1#)
    st.PreBufferSec = NzD(wsCfg.Cells(8, valCol).Value, 5#)
    st.ChartIntervalSec = NzD(wsCfg.Cells(9, valCol).Value, 120#)
    st.SteadyWindowSec = NzD(wsCfg.Cells(10, valCol).Value, 30#)
    st.React15ThrPct = NzD(wsCfg.Cells(11, valCol).Value, 0.5)
    st.AvgWindowSec = NzD(wsCfg.Cells(12, valCol).Value, 0#)
    If st.React15ThrPct < 0 Then st.React15ThrPct = 0
    If st.AvgWindowSec < 0 Then st.AvgWindowSec = 0
    modeTxt = UCase$(Trim$(CStr(wsCfg.Cells(13, valCol).Value)))
    If Len(modeTxt) = 0 Then modeTxt = "АВАРИЯ"
    If Left$(modeTxt, 4) = "ИСПЫ" Or Left$(modeTxt, 4) = "ТЕСТ" Or Left$(modeTxt, 4) = "TEST" Then
        st.Mode = "Испытание"
        st.IsEmergency = False
    Else
        st.Mode = "Авария"
        st.IsEmergency = True
    End If
    st.MaxEventSec = NzD(wsCfg.Cells(14, valCol).Value, 600#)
    If st.MaxEventSec < 60 Then st.MaxEventSec = 60
    st.MaxEvents = CLng(NzD(wsCfg.Cells(15, valCol).Value, 3#))
    If st.MaxEvents < 1 Then st.MaxEvents = 1
    st.MaxReqPct = NzD(wsCfg.Cells(16, valCol).Value, 10#)
    If st.MaxReqPct < 0 Then st.MaxReqPct = 0
    If Len(Trim$(CStr(wsCfg.Cells(17, 23).Value))) > 0 Or Len(Trim$(CStr(wsCfg.Cells(17, valCol).Value))) > 0 Then
        st.StrictDeficit = (NzD(wsCfg.Cells(17, valCol).Value, IIf(st.IsEmergency, 1, 0)) <> 0)
    Else
        st.StrictDeficit = st.IsEmergency
    End If
    If IsDate(wsCfg.Cells(3, valCol).Value) Then
        st.EventStart = CDate(wsCfg.Cells(3, valCol).Value)
    Else
        st.EventStart = 0
    End If
    ReadSettings = st
End Function

Public Function ReadGenCfg(ByVal ws As Worksheet, ByVal r As Long) As TGenCfg
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
    g.HasRM = (NzD(ws.Cells(r, 21).Value, 1) <> 0)
    g.Layout = LCase$(Trim$(CStr(ws.Cells(r, 22).Value)))
    If Len(g.Layout) = 0 Then
        If InStr(UCase$(g.EquipType), "НЕБЛОК") > 0 Then g.Layout = "неблок" Else g.Layout = "блок"
    ElseIf InStr(g.Layout, "неблок") > 0 Or InStr(g.Layout, "общ") > 0 Or InStr(g.Layout, "попереч") > 0 Then
        g.Layout = "неблок"
    Else
        g.Layout = "блок"
    End If
    g.SPctCfg = g.SPct
    g.FnchCfg = g.Fnch
    g.ForcedNoRM = False

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

Public Function GetPreset(ByVal equipType As String) As TGenCfg
    Dim g As TGenCfg, et As String
    et = UCase$(Trim$(equipType))
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
    ElseIf InStr(et, "ПТУ") > 0 Or InStr(et, "ПСУ") > 0 Then
        g.T10Sec = 300
    End If
    GetPreset = g
End Function

Public Function ValidateGenCfg(ByRef g As TGenCfg) As Boolean
    ValidateGenCfg = (Len(g.Station) > 0 And Len(g.Generator) > 0 And Len(g.PowerHeader) > 0 And Len(g.FreqHeader) > 0 _
                      And g.PNom > 0 And g.SPct > 0 And g.Fnch >= 0)
End Function


Public Function RowByTimeOffset(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal startRow As Long, ByVal offsetSec As Double) As Long
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

Public Function MaxAbsDeviationInWindow(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal fNom As Double) As Double
    Dim r As Long, d As Double, best As Double
    best = 0
    For r = r1 To r2
        d = NzD(wsRaw.Cells(r, freqCol).Value, fNom) - fNom
        If Abs(d) > Abs(best) Then best = d
    Next r
    MaxAbsDeviationInWindow = best
End Function

Public Function DeadbandDeviation(ByVal dF As Double, ByVal fnch As Double) As Double
    If Abs(dF) <= fnch Then DeadbandDeviation = 0 Else DeadbandDeviation = Sgn(dF) * (Abs(dF) - fnch)
End Function

Public Function SecBetween(ByVal t1 As Variant, ByVal t2 As Variant) As Double
    If IsDate(t1) And IsDate(t2) Then
        SecBetween = Abs(CDbl(CDate(t2) - CDate(t1))) * 86400#
    Else
        SecBetween = 0
    End If
End Function

Public Function AmplitudeTag(ByVal amplPct As Double, ByVal isEvent As Boolean) As String
    If Not isEvent Then AmplitudeTag = "": Exit Function
    If amplPct < 3# Then
        AmplitudeTag = "Слабое"
    ElseIf amplPct > 10# Then
        AmplitudeTag = "Избыточное"
    Else
        AmplitudeTag = "Норма"
    End If
End Function

Public Sub LoadKdProfiles(ByVal wsCfg As Worksheet)
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

Public Function EvalKdAbs(ByVal tSec As Double, ByVal prof As Variant) As Double
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

Public Function ClampKd(ByVal kdVal As Double) As Double
    ClampKd = kdVal
    If ClampKd < 0.1 Then ClampKd = 0.1
    If ClampKd > 1# Then ClampKd = 1#
End Function


Public Function ResolveKdKey(ByVal equipType As String) As String
    Dim key As String
    key = UCase$(Trim$(equipType))
    ResolveKdKey = key
    If m_KdProfiles Is Nothing Then Exit Function
    If m_KdProfiles.Exists(key) Then Exit Function
    If InStr(key, "ГПА") > 0 And m_KdProfiles.Exists("ГПА") Then ResolveKdKey = "ГПА": Exit Function
    If InStr(key, "ГТУ") > 0 And m_KdProfiles.Exists("ГТУ") Then ResolveKdKey = "ГТУ": Exit Function
    If (InStr(key, "ПГУ") > 0 And (InStr(key, "СБРОС") > 0 Or InStr(key, "ПАРАЛ") > 0)) And m_KdProfiles.Exists("ПГУ_СБРОСН") Then
        ResolveKdKey = "ПГУ_СБРОСН": Exit Function
    End If
    If InStr(key, "ПГУ") > 0 And m_KdProfiles.Exists("ПГУ_УТИЛ") Then ResolveKdKey = "ПГУ_УТИЛ": Exit Function
    If (InStr(key, "НЕБЛОК") > 0 Or InStr(key, "ОБЩ") > 0) And m_KdProfiles.Exists("ПТУ_НЕБЛОК") Then
        ResolveKdKey = "ПТУ_НЕБЛОК": Exit Function
    End If
    If (InStr(key, "УГОЛЬ") > 0 Or InStr(key, "ПЫЛЕУГ") > 0) And m_KdProfiles.Exists("ПТУ_БЛОК_УГОЛЬ") Then
        ResolveKdKey = "ПТУ_БЛОК_УГОЛЬ": Exit Function
    End If
    If (InStr(key, "ПТУ") > 0 Or InStr(key, "ПСУ") > 0 Or InStr(key, "БЛОК") > 0) And m_KdProfiles.Exists("ПТУ_БЛОК") Then
        ResolveKdKey = "ПТУ_БЛОК": Exit Function
    End If
End Function

Public Function DynamicKdByTime(ByVal equipType As String, ByVal tSec As Double) As Double
    Dim prof As Variant, key As String
    If m_KdProfiles Is Nothing Then Err.Raise vbObjectError + 2401, , "Профили Kд(t) не загружены."
    key = ResolveKdKey(equipType)
    If Not m_KdProfiles.Exists(key) Then
        Err.Raise vbObjectError + 2402, , "Для типа '" & equipType & "' не задан профиль Kд(t) в Config (AA:AI)."
    End If
    prof = m_KdProfiles(key)
    DynamicKdByTime = ClampKd(EvalKdAbs(tSec, prof))
End Function

Public Function KdProfileText(ByVal equipType As String) As String
    Dim prof As Variant, key As String, src As String
    src = "Config"
    If m_KdProfiles Is Nothing Then
        KdProfileText = "Config: профили Kд(t) не загружены"
        Exit Function
    End If
    key = ResolveKdKey(equipType)
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

Public Function StationMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim na As String, nb As String
    na = UCase$(Trim$(a))
    nb = UCase$(Trim$(b))
    If InStr(na, "ВОРКУТ") > 0 And InStr(nb, "ВОРКУТ") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СОСНОГОР") > 0 And InStr(nb, "СОСНОГОР") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СЛПК") > 0 And InStr(nb, "СЛПК") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "ПЕЧОР") > 0 And InStr(nb, "ПЕЧОР") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "УСИНС") > 0 And InStr(nb, "УСИНС") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "ЯРЕГ") > 0 And InStr(nb, "ЯРЕГ") > 0 Then StationMatch = True: Exit Function
    StationMatch = (na = nb)
End Function

Public Function MinArray(ByRef arr() As Double) As Double
    Dim i As Long, v As Double
    v = 1E+99
    For i = LBound(arr) To UBound(arr)
        If arr(i) > 0 And arr(i) < v Then v = arr(i)
    Next i
    If v = 1E+99 Then v = 0.05
    MinArray = v
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxD = a Else MaxD = b
End Function

Public Function FormatSecOrNA(ByVal secVal As Double) As String
    If secVal < 0 Then FormatSecOrNA = "н/д" Else FormatSecOrNA = Format(secVal, "0.0")
End Function


Public Function EventSuffix(ByVal evIdx As Long) As String
    If evIdx > 1 Then EventSuffix = "_E" & evIdx Else EventSuffix = ""
End Function

Public Function MakeSheetName(ByVal s As String) As String
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

Public Function GetRequiredSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetRequiredSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetRequiredSheet Is Nothing Then
        Err.Raise vbObjectError + 2999, , "Не найден лист '" & name & "'. Запустите SetupOPRCHTemplate."
    End If
End Function

Public Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Public Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, v As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(v, headerName, vbTextCompare) = 0 Then FindHeaderCol = c: Exit Function
    Next c
    FindHeaderCol = 0
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Public Function NzD(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        NzD = fallback
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = fallback
    End If
End Function

Public Function SafeDiv(ByVal a As Double, ByVal b As Double, ByVal fallback As Double) As Double
    If Abs(b) < 0.0000001 Then SafeDiv = fallback Else SafeDiv = a / b
End Function

Public Function SgnNZ(ByVal v As Double) As Integer
    If v > 0.000001 Then
        SgnNZ = 1
    ElseIf v < -0.000001 Then
        SgnNZ = -1
    Else
        SgnNZ = 0
    End If
End Function
