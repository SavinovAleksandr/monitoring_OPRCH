Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' ==========================================================
' OPRCH monitoring (quantitative + qualitative)
' Per-generator sheets + station aggregate sheets.
' ==========================================================

Private Const SH_RAW As String = "RawData"
Private Const SH_CFG As String = "Config"
Private Const SH_SUM As String = "Summary"

Private Type TSettings
    FNom As Double
    EventStart As Date
    AutoStart As Boolean
    QuantIntervalSec As Double
    QuantTolPct As Double
    WorkThresholdMW As Double
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
    Kd As Double
    Enabled As Boolean
    QualEnabled As Boolean
    T5Sec As Double
    Dp5Pct As Double
    T10Sec As Double
    Dp10Pct As Double
    SteadyTolPct As Double
    InStationSum As Boolean
End Type

Private Type TGenResult
    StartRow As Long
    EndQuantRow As Long
    StartTime As Variant
    P0 As Double
    PTek As Double
    Df As Double
    Dfr As Double
    PReq As Double
    PFact As Double
    QuantPct As Double
    QuantPass As Boolean
    QualPass As Boolean
    QualReason As String
End Type

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

    wsCfg.Range("A1:Q1").Value = Array( _
        "Станция", "Генератор", "Колонка_мощности", "Колонка_частоты", "Тип_оборудования", _
        "Pном, МВт", "S, %", "fнч, Гц", "Kд", "Вкл (1/0)", "Кач_вкл (1/0)", _
        "t5, c", "dP5, %Pном", "t10, c", "dP10, %Pном", "Уст_допуск, %Pном", "В сумму станции (1/0)" _
    )

    wsCfg.Cells(2, 1).Resize(1, 17).Value = Array("Сосногорская ТЭЦ", "ТГ-5", "ТГ-5", "Частота", "ПТУ", 55, 4.2, 0.105, 0.5, 1, 1, 15, 5, 420, 10, 1, 1)
    wsCfg.Cells(3, 1).Resize(1, 17).Value = Array("Сосногорская ТЭЦ", "ТГ-7", "ТГ-7", "Частота", "ПТУ", 60, 4.5, 0.11, 0.5, 1, 1, 15, 5, 420, 10, 1, 1)
    wsCfg.Cells(4, 1).Resize(1, 17).Value = Array("СЛПК", "ТГ-2Э", "ТГ-2Э", "f СЛПК", "ПГУ_утилиз", 50, 4.5, 0.1, 0.5, 1, 1, 15, 5, 300, 10, 1, 1)

    wsCfg.Range("S1").Value = "Глобальные настройки"
    wsCfg.Cells(2, 18).Resize(1, 2).Value = Array("fном, Гц", 50)
    wsCfg.Cells(3, 18).Resize(1, 2).Value = Array("Время начала события", "")
    wsCfg.Cells(4, 18).Resize(1, 2).Value = Array("Автопоиск старта (1/0)", 1)
    wsCfg.Cells(5, 18).Resize(1, 2).Value = Array("Колич. интервал, с", 82)
    wsCfg.Cells(6, 18).Resize(1, 2).Value = Array("Допуск количеств., %", 10)
    wsCfg.Cells(7, 18).Resize(1, 2).Value = Array("Порог включения в работу, МВт", 1)

    wsCfg.Columns("A:T").AutoFit
    wsRaw.Columns("A:E").AutoFit
    EnsureRunButton wsCfg

    MsgBox "Шаблон создан. Заполните RawData/Config и нажмите кнопку 'Запустить мониторинг ОПРЧ' на листе Config.", vbInformation
End Sub

Public Sub AnalyzeOPRCH()
    On Error GoTo EH

    Dim wsRaw As Worksheet, wsCfg As Worksheet, wsSummary As Worksheet
    Dim st As TSettings
    Dim timeCol As Long, cfgLast As Long, r As Long, outRow As Long
    Dim g As TGenCfg, res As TGenResult
    Dim targetSheets As Collection

    Set wsRaw = GetRequiredSheet(SH_RAW)
    Set wsCfg = GetRequiredSheet(SH_CFG)
    Set wsSummary = EnsureSheet(SH_SUM)

    st = ReadSettings(wsCfg)
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Err.Raise vbObjectError + 2001, , "В RawData не найдена колонка 'Время'."

    cfgLast = LastUsedRow(wsCfg)
    If cfgLast < 2 Then Err.Raise vbObjectError + 2002, , "В Config нет строк генераторов."

    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    DeleteOutputSheets targetSheets

    wsSummary.Cells.Clear
    wsSummary.Range("A1:Q1").Value = Array( _
        "Станция", "Генератор", "Тип", "Старт", "P0, МВт", "Pтек, МВт", "dF, Гц", "dFr, Гц", _
        "Pтреб, МВт", "Pфакт, МВт", "Колич. %", "Колич. статус", "Кач. статус", "t5 факт, c", "t10 факт, c", "Лист", "Примечание" _
    )
    outRow = 2

    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextGen

        g = ReadGenCfg(wsCfg, r)
        If Not g.Enabled Then GoTo NextGen

        If Not ValidateGenCfg(g) Then
            wsSummary.Cells(outRow, 1).Resize(1, 17).Value = Array(g.Station, g.Generator, g.EquipType, "", "", "", "", "", "", "", "", "Нарушение", "Н/Д", "", "", "", "Не заполнен обязательный параметр config")
            outRow = outRow + 1
            GoTo NextGen
        End If

        res = AnalyzeOneGenerator(wsRaw, st, g)
        WriteGeneratorSheet wsRaw, st, g, res
        WriteSummaryRow wsSummary, outRow, g, res
        outRow = outRow + 1

NextGen:
    Next r

    BuildStationAggregates wsRaw, wsCfg, wsSummary, st

    wsSummary.Columns("A:Q").AutoFit
    wsSummary.Range("D:D").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    EnsureRunButton wsCfg
    MsgBox "Мониторинг ОПРЧ завершен.", vbInformation
    Exit Sub

EH:
    MsgBox "Ошибка AnalyzeOPRCH: " & Err.Description, vbCritical
End Sub

Private Function AnalyzeOneGenerator(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg) As TGenResult
    Dim res As TGenResult
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim startRow As Long, endQ As Long
    Dim p0 As Double, ptek As Double, df As Double, dfr As Double, preq As Double, pfact As Double
    Dim qpct As Double, qpass As Boolean

    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderCol(wsRaw, g.PowerHeader)
    fCol = FindHeaderCol(wsRaw, g.FreqHeader)
    If pCol = 0 Then Err.Raise vbObjectError + 2101, , "Не найдена колонка мощности '" & g.PowerHeader & "' для " & g.Generator
    If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator

    startRow = ResolveStartRow(wsRaw, timeCol, fCol, st, g.Fnch)
    endQ = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)

    p0 = NzD(wsRaw.Cells(startRow, pCol).Value, 0)
    ptek = NzD(wsRaw.Cells(endQ, pCol).Value, 0)
    df = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, endQ, st.FNom)
    dfr = DeadbandDeviation(df, g.Fnch)

    If dfr <> 0 Then
        preq = -100# / g.SPct * g.PNom / st.FNom * g.Kd * dfr
    Else
        preq = 0
    End If
    pfact = ptek - p0

    If dfr = 0 Then
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
    res.StartTime = wsRaw.Cells(startRow, timeCol).Value
    res.P0 = p0
    res.PTek = ptek
    res.Df = df
    res.Dfr = dfr
    res.PReq = preq
    res.PFact = pfact
    res.QuantPct = qpct
    res.QuantPass = qpass

    EvaluateQualitative wsRaw, st, g, res, pCol, fCol, timeCol
    AnalyzeOneGenerator = res
End Function

Private Sub EvaluateQualitative(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long)
    Dim signReq As Integer, t5 As Double, t10 As Double
    Dim r As Long, endRow As Long, dP As Double, target5 As Double, target10 As Double
    Dim hit5 As Boolean, hit10 As Boolean, row5 As Long, row10 As Long
    Dim endVal As Double, steadyTolMW As Double
    Dim reason As String

    If Not g.QualEnabled Then
        res.QualPass = True
        res.QualReason = "Качественная проверка отключена"
        Exit Sub
    End If

    signReq = SgnNZ(res.PReq)
    If signReq = 0 Then
        res.QualPass = True
        res.QualReason = "Качественная оценка: вне зоны отклонения"
        Exit Sub
    End If

    target5 = signReq * g.PNom * g.Dp5Pct / 100#
    target10 = signReq * g.PNom * g.Dp10Pct / 100#

    endRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, g.T10Sec)
    row5 = 0
    row10 = 0
    hit5 = False
    hit10 = False

    For r = res.StartRow To endRow
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

    t5 = IIf(hit5, SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row5, timeCol).Value), -1)
    t10 = IIf(hit10, SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row10, timeCol).Value), -1)

    endVal = NzD(wsRaw.Cells(endRow, pCol).Value, 0) - res.P0
    steadyTolMW = g.SteadyTolPct / 100# * g.PNom

    reason = ""
    If Not hit5 Or t5 > g.T5Sec Then reason = reason & "Не достигнут dP5 в t5; "
    If Not hit10 Or t10 > g.T10Sec Then reason = reason & "Не достигнут dP10 в t10; "
    If Abs(endVal - target10) > steadyTolMW Then reason = reason & "Отклонение от целевого установившегося > допуска; "

    If Len(reason) = 0 Then
        res.QualPass = True
        res.QualReason = "Качественные критерии выполнены; t5=" & Format(t5, "0.0") & "с; t10=" & Format(t10, "0.0") & "с"
    Else
        res.QualPass = False
        res.QualReason = reason & " t5=" & IIf(t5 >= 0, Format(t5, "0.0"), "н/д") & "с; t10=" & IIf(t10 >= 0, Format(t10, "0.0"), "н/д") & "с"
    End If
End Sub

Private Sub WriteGeneratorSheet(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim ws As Worksheet
    Dim shName As String
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim endRow As Long, r As Long, outR As Long
    Dim dP As Double, dFr As Double
    Dim chartObj As ChartObject

    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderCol(wsRaw, g.PowerHeader)
    fCol = FindHeaderCol(wsRaw, g.FreqHeader)

    shName = MakeSheetName(g.Station & "_" & g.Generator)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    ws.Range("A1:B1").Value = Array("Станция", g.Station)
    ws.Range("A2:B2").Value = Array("Генератор", g.Generator)
    ws.Range("A3:B3").Value = Array("Тип", g.EquipType)
    ws.Range("A4:B4").Value = Array("Старт", res.StartTime)
    ws.Range("A5:B5").Value = Array("Колич. статус", IIf(res.QuantPass, "ОК", "Нарушение"))
    ws.Range("A6:B6").Value = Array("Кач. статус", IIf(res.QualPass, "ОК", "Нарушение"))
    ws.Range("A7:B7").Value = Array("Кач. примечание", res.QualReason)

    ws.Range("D1:E8").Value = Array( _
        Array("P0, МВт", res.P0), _
        Array("Pтек, МВт", res.PTek), _
        Array("dF, Гц", res.Df), _
        Array("dFr, Гц", res.Dfr), _
        Array("Pтреб, МВт", res.PReq), _
        Array("Pфакт, МВт", res.PFact), _
        Array("Колич. %", res.QuantPct), _
        Array("Интервал, с", st.QuantIntervalSec) _
    )

    ws.Range("A10:F10").Value = Array("Время", "Частота, Гц", "P, МВт", "dPфакт, МВт", "Pтреб_накоп, МВт", "dFr, Гц")
    endRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, MaxD(st.QuantIntervalSec, g.T10Sec))
    outR = 11

    For r = res.StartRow To endRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        dFr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, g.Fnch)

        ws.Cells(outR, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(outR, 2).Value = wsRaw.Cells(r, fCol).Value
        ws.Cells(outR, 3).Value = wsRaw.Cells(r, pCol).Value
        ws.Cells(outR, 4).Value = dP
        ws.Cells(outR, 5).Value = -100# / g.SPct * g.PNom / st.FNom * g.Kd * dFr
        ws.Cells(outR, 6).Value = dFr
        outR = outR + 1
    Next r

    Set chartObj = ws.ChartObjects.Add(20, 320, 950, 280)
    chartObj.Chart.ChartType = xlLine
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Мониторинг ОПРЧ: " & g.Generator

    AddSeries chartObj.Chart, ws, 11, outR - 1, 1, 4, "dPфакт, МВт", False
    AddSeries chartObj.Chart, ws, 11, outR - 1, 1, 5, "Pтреб, МВт", False
    AddSeries chartObj.Chart, ws, 11, outR - 1, 1, 2, "Частота, Гц", True

    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"

    ws.Columns("A:F").AutoFit
    ws.Range("A:A").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B:F").NumberFormat = "0.000"
End Sub

Private Sub BuildStationAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, ByRef st As TSettings)
    Dim stations As Variant, s As Variant
    stations = Array("Сосногорская ТЭЦ", "Воркутинская ТЭЦ", "СЛПК")
    For Each s In stations
        BuildOneStationAggregate wsRaw, wsCfg, CStr(s), st
    Next s
End Sub

Private Sub BuildOneStationAggregate(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal stationName As String, ByRef st As TSettings)
    Dim cfgLast As Long, r As Long, cnt As Long
    Dim g As TGenCfg
    Dim pCols() As Long, fnchArr() As Double, sArr() As Double, kdArr() As Double, pnomArr() As Double
    Dim freqCol As Long, timeCol As Long
    Dim shName As String, ws As Worksheet
    Dim startRow As Long, endRow As Long, rowQ As Long
    Dim p0 As Double, pNow As Double, preq As Double, pfact As Double, dF As Double
    Dim i As Long, dfr As Double

    cfgLast = LastUsedRow(wsCfg)
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Exit Sub

    cnt = 0
    For r = 2 To cfgLast
        g = ReadGenCfg(wsCfg, r)
        If g.Enabled And g.InStationSum Then
            If StationMatch(g.Station, stationName) Then
                cnt = cnt + 1
                ReDim Preserve pCols(1 To cnt)
                ReDim Preserve fnchArr(1 To cnt)
                ReDim Preserve sArr(1 To cnt)
                ReDim Preserve kdArr(1 To cnt)
                ReDim Preserve pnomArr(1 To cnt)
                pCols(cnt) = FindHeaderCol(wsRaw, g.PowerHeader)
                If pCols(cnt) = 0 Then cnt = cnt - 1: GoTo NextCfg
                fnchArr(cnt) = g.Fnch
                sArr(cnt) = g.SPct
                kdArr(cnt) = g.Kd
                pnomArr(cnt) = g.PNom
                If freqCol = 0 Then freqCol = FindHeaderCol(wsRaw, g.FreqHeader)
            End If
        End If
NextCfg:
    Next r
    If cnt = 0 Then Exit Sub
    If freqCol = 0 Then Exit Sub

    startRow = ResolveStartRow(wsRaw, timeCol, freqCol, st, MinArray(fnchArr))
    endRow = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)

    p0 = 0
    For i = 1 To cnt
        If NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0) > st.WorkThresholdMW Then
            p0 = p0 + NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0)
        End If
    Next i

    pNow = 0
    For i = 1 To cnt
        If NzD(wsRaw.Cells(endRow, pCols(i)).Value, 0) > st.WorkThresholdMW Then
            pNow = pNow + NzD(wsRaw.Cells(endRow, pCols(i)).Value, 0)
        End If
    Next i

    dF = MaxAbsDeviationInWindow(wsRaw, freqCol, startRow, endRow, st.FNom)
    preq = 0
    For i = 1 To cnt
        dfr = DeadbandDeviation(dF, fnchArr(i))
        preq = preq + (-100# / sArr(i) * pnomArr(i) / st.FNom * kdArr(i) * dfr)
    Next i
    pfact = pNow - p0

    shName = MakeSheetName(stationName & "_Сумма")
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    ws.Range("A1:B1").Value = Array("Станция", stationName)
    ws.Range("A2:B2").Value = Array("Режим", "Суммарная нагрузка включенных генераторов")
    ws.Range("A3:B3").Value = Array("Старт", wsRaw.Cells(startRow, timeCol).Value)
    ws.Range("A4:B4").Value = Array("P0, МВт", p0)
    ws.Range("A5:B5").Value = Array("Pтек, МВт", pNow)
    ws.Range("A6:B6").Value = Array("Pтреб, МВт", preq)
    ws.Range("A7:B7").Value = Array("Pфакт, МВт", pfact)

    ws.Range("A10:E10").Value = Array("Время", "Частота, Гц", "Pсум, МВт", "dPсум, МВт", "Pтреб_сум, МВт")
    rowQ = 11
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

        preq = 0
        For i = 1 To cnt
            dfr = DeadbandDeviation(NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom, fnchArr(i))
            preq = preq + (-100# / sArr(i) * pnomArr(i) / st.FNom * kdArr(i) * dfr)
        Next i
        ws.Cells(rowQ, 5).Value = preq
        rowQ = rowQ + 1
    Next r

    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(20, 290, 930, 260)
    chartObj.Chart.ChartType = xlLine
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Суммарный мониторинг ОПРЧ: " & stationName
    AddSeries chartObj.Chart, ws, 11, rowQ - 1, 1, 4, "dPсум, МВт", False
    AddSeries chartObj.Chart, ws, 11, rowQ - 1, 1, 5, "Pтреб_сум, МВт", False
    AddSeries chartObj.Chart, ws, 11, rowQ - 1, 1, 2, "Частота, Гц", True
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"

    ws.Columns("A:E").AutoFit
    ws.Range("A:A").NumberFormat = "dd.mm.yyyy hh:mm:ss"
End Sub

Private Sub WriteSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim note As String
    note = ""
    If Not res.QuantPass Then note = "Колич. критерий не выполнен"
    ws.Cells(r, 1).Resize(1, 17).Value = Array( _
        g.Station, g.Generator, g.EquipType, res.StartTime, res.P0, res.PTek, res.Df, res.Dfr, _
        res.PReq, res.PFact, res.QuantPct, IIf(res.QuantPass, "ОК", "Нарушение"), _
        IIf(res.QualPass, "ОК", "Нарушение"), ExtractT(res.QualReason, "t5"), ExtractT(res.QualReason, "t10"), _
        MakeSheetName(g.Station & "_" & g.Generator), note _
    )
End Sub

Private Function ReadSettings(ByVal wsCfg As Worksheet) As TSettings
    Dim st As TSettings
    st.FNom = NzD(wsCfg.Range("T2").Value, 50#)
    st.AutoStart = (NzD(wsCfg.Range("T4").Value, 1) <> 0)
    st.QuantIntervalSec = NzD(wsCfg.Range("T5").Value, 82#)
    st.QuantTolPct = NzD(wsCfg.Range("T6").Value, 10#)
    st.WorkThresholdMW = NzD(wsCfg.Range("T7").Value, 1#)
    If IsDate(wsCfg.Range("T3").Value) Then
        st.EventStart = CDate(wsCfg.Range("T3").Value)
    Else
        st.EventStart = 0
    End If
    ReadSettings = st
End Function

Private Function ReadGenCfg(ByVal ws As Worksheet, ByVal r As Long) As TGenCfg
    Dim g As TGenCfg
    g.Station = Trim$(CStr(ws.Cells(r, 1).Value))
    g.Generator = Trim$(CStr(ws.Cells(r, 2).Value))
    g.PowerHeader = Trim$(CStr(ws.Cells(r, 3).Value))
    g.FreqHeader = Trim$(CStr(ws.Cells(r, 4).Value))
    g.EquipType = Trim$(CStr(ws.Cells(r, 5).Value))
    g.PNom = NzD(ws.Cells(r, 6).Value, 0)
    g.SPct = NzD(ws.Cells(r, 7).Value, 0)
    g.Fnch = NzD(ws.Cells(r, 8).Value, 0)
    g.Kd = NzD(ws.Cells(r, 9).Value, 0.5)
    g.Enabled = (NzD(ws.Cells(r, 10).Value, 1) <> 0)
    g.QualEnabled = (NzD(ws.Cells(r, 11).Value, 1) <> 0)
    g.T5Sec = NzD(ws.Cells(r, 12).Value, 15)
    g.Dp5Pct = NzD(ws.Cells(r, 13).Value, 5)
    g.T10Sec = NzD(ws.Cells(r, 14).Value, 420)
    g.Dp10Pct = NzD(ws.Cells(r, 15).Value, 10)
    g.SteadyTolPct = NzD(ws.Cells(r, 16).Value, 1)
    g.InStationSum = (NzD(ws.Cells(r, 17).Value, 1) <> 0)
    If Len(g.PowerHeader) = 0 Then g.PowerHeader = g.Generator
    If Len(g.FreqHeader) = 0 Then g.FreqHeader = "Частота"
    ReadGenCfg = g
End Function

Private Function ValidateGenCfg(ByRef g As TGenCfg) As Boolean
    ValidateGenCfg = (Len(g.Station) > 0 And Len(g.Generator) > 0 And Len(g.PowerHeader) > 0 And Len(g.FreqHeader) > 0 _
                      And g.PNom > 0 And g.SPct > 0 And g.Fnch >= 0 And g.Kd > 0)
End Function

Private Function ResolveStartRow(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal freqCol As Long, _
                                 ByRef st As TSettings, ByVal fnch As Double) As Long
    Dim lastR As Long, r As Long
    Dim prevAbs As Double, curAbs As Double
    Dim bestRow As Long, bestAbs As Double, val As Double, t As Date, dt As Double

    lastR = LastUsedRow(wsRaw)
    If lastR < 3 Then
        ResolveStartRow = 2
        Exit Function
    End If

    If Not st.AutoStart And st.EventStart > 0 Then
        bestRow = 2
        bestAbs = 1E+99
        For r = 2 To lastR
            If IsDate(wsRaw.Cells(r, timeCol).Value) Then
                t = CDate(wsRaw.Cells(r, timeCol).Value)
                dt = Abs(CDbl(t - st.EventStart))
                If dt < bestAbs Then
                    bestAbs = dt
                    bestRow = r
                End If
            End If
        Next r
        ResolveStartRow = bestRow
        Exit Function
    End If

    prevAbs = Abs(NzD(wsRaw.Cells(2, freqCol).Value, st.FNom) - st.FNom)
    For r = 3 To lastR
        curAbs = Abs(NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom)
        If prevAbs <= fnch And curAbs > fnch Then
            ResolveStartRow = r - 1
            Exit Function
        End If
        prevAbs = curAbs
    Next r

    bestRow = 2
    bestAbs = -1
    For r = 2 To lastR
        val = NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom
        If Abs(val) > bestAbs Then
            bestAbs = Abs(val)
            bestRow = r
        End If
    Next r
    ResolveStartRow = bestRow
End Function

Private Function RowByTimeOffset(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal startRow As Long, ByVal offsetSec As Double) As Long
    Dim t0 As Date, tTarget As Date
    Dim r As Long, lastR As Long, bestRow As Long
    Dim curT As Date, bestDiff As Double, curDiff As Double

    If Not IsDate(wsRaw.Cells(startRow, timeCol).Value) Then
        RowByTimeOffset = startRow
        Exit Function
    End If

    t0 = CDate(wsRaw.Cells(startRow, timeCol).Value)
    tTarget = DateAdd("s", CLng(offsetSec), t0)
    lastR = LastUsedRow(wsRaw)
    bestRow = startRow
    bestDiff = 1E+99

    For r = startRow To lastR
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            curT = CDate(wsRaw.Cells(r, timeCol).Value)
            curDiff = Abs(CDbl(curT - tTarget))
            If curDiff < bestDiff Then
                bestDiff = curDiff
                bestRow = r
            End If
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
    If Abs(dF) <= fnch Then
        DeadbandDeviation = 0
    Else
        DeadbandDeviation = Sgn(dF) * (Abs(dF) - fnch)
    End If
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

Private Sub CollectOldOutputSheets(ByRef names As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_RAW And ws.Name <> SH_CFG And ws.Name <> SH_SUM Then
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

Private Function ExtractT(ByVal txt As String, ByVal key As String) As String
    Dim p As Long, s As String
    p = InStr(1, txt, key & "=", vbTextCompare)
    If p = 0 Then
        ExtractT = ""
        Exit Function
    End If
    s = Mid$(txt, p)
    p = InStr(s, ";")
    If p > 0 Then s = Left$(s, p - 1)
    ExtractT = s
End Function

Private Function MakeSheetName(ByVal s As String) As String
    Dim out As String, i As Long, ch As String
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
        If StrComp(v, headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
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
    If Abs(b) < 0.0000001 Then
        SafeDiv = fallback
    Else
        SafeDiv = a / b
    End If
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

Private Sub EnsureRunButton(ByVal ws As Worksheet)
    Const BTN_NAME As String = "btnRunOPRCH"
    Dim shp As Shape

    On Error Resume Next
    ws.Shapes(BTN_NAME).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 870, 8, 250, 30)
    shp.Name = BTN_NAME
    shp.TextFrame2.TextRange.Text = "Запустить мониторинг ОПРЧ"
    shp.OnAction = "AnalyzeOPRCH"
    shp.Fill.ForeColor.RGB = RGB(40, 120, 220)
    shp.Line.ForeColor.RGB = RGB(20, 70, 140)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 11
End Sub
