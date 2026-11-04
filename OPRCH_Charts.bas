Attribute VB_Name = "OPRCH_Charts"
Option Explicit

' Графики генераторов и станции. Импортировать 4.
' OPRCH_Monitor v1.8.0 — часть раздельного модуля (лимит VBA ~64 КБ на модуль).

Public Sub FillMarkerColumn(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
                             ByVal wsRaw As Worksheet, ByVal timeCol As Long, _
                             ByVal startRow As Long, ByVal offsetSec As Double, _
                             ByVal outCol As Long, ByVal vHi As Double, ByVal vLo As Double)
    Dim mkRow As Long
    mkRow = RowByTimeOffset(wsRaw, timeCol, startRow, offsetSec)
    If mkRow < 2 Then Exit Sub
    FillMarkerColumnAtTime ws, rr1, rr2, wsRaw, timeCol, wsRaw.Cells(mkRow, timeCol).Value, outCol, vHi, vLo
End Sub

Public Sub FillMarkerColumnAtTime(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
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

Public Sub ApplyGeneratorSheetFormats(ByVal ws As Worksheet)
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
    ws.Range("A11:W11").NumberFormat = "@"
    ws.Range("A12:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B12:W100000").NumberFormat = "0.000"
    ws.Range("X1:X11").NumberFormat = "@"
    ws.Range("Y1").NumberFormat = "0"
    ws.Range("Y2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("Y3").NumberFormat = "0"
    ws.Range("Y4:Y11").NumberFormat = "@"
    ws.Range("Y4:Y11").WrapText = True
End Sub

' ==========================================================
' Лист с графиком
' ==========================================================

Public Sub WriteGeneratorChartSheet(ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataSheet As String, chartSheet As String
    Dim lastDataRow As Long
    Dim chartObj As ChartObject
    Dim startDataRow As Long, endChartDataRow As Long
    Dim ySpan As Double
    Dim yMin As Double, yMax As Double, yPad As Double
    Dim farPos As Double, farNeg As Double, farAbs As Double

    dataSheet = GeneratorSheetName(g, res.EventIdx)
    chartSheet = GeneratorChartSheetName(g, res.EventIdx)
    Set wsData = ThisWorkbook.Worksheets(dataSheet)
    Set wsChart = EnsureSheet(chartSheet)

    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    ' Заголовок
    wsChart.Range("A1").Value = "График ОПРЧ: " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    ' Блок "Вывод"
    WriteChartVerdictBlock wsChart, g, res, st

    lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 12 Then Exit Sub

    startDataRow = 12
    endChartDataRow = FindChartEndRow(wsData, 12, lastDataRow, res.StartTime, st.ChartIntervalSec)
    If endChartDataRow < startDataRow Then endChartDataRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 215, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ: " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
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
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 7, "tр"
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
    ApplyChartAxesFormat chartObj.Chart, wsData, startDataRow, endChartDataRow, "dP, МВт", "f, Гц"

    ' Второй график: абсолютные единицы (P, МВт и f, Гц)
    Set chartObj = wsChart.ChartObjects.Add(10, 595, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ (абсолютные): " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
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
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 23, "tр abs"
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
    ApplyChartAxesFormat chartObj.Chart, wsData, startDataRow, endChartDataRow, "P, МВт", "f, Гц"
End Sub

' Оформление осей по И.2 Приложения 1: подписи осей с единицами измерения,
' вертикальная сетка не реже чем через 15 с (по числу точек на 15 с), подписи времени.

Public Sub ApplyChartAxesFormat(ByVal ch As Chart, ByVal wsData As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal yTitle As String, ByVal y2Title As String)
    Dim n As Long, spanSec As Double, ptsPer15 As Long
    On Error Resume Next
    n = r2 - r1 + 1
    spanSec = SecBetween(wsData.Cells(r1, 1).Value, wsData.Cells(r2, 1).Value)
    If spanSec > 0 And n > 1 Then
        ptsPer15 = CLng(15# * (n - 1) / spanSec)
    Else
        ptsPer15 = 15
    End If
    If ptsPer15 < 1 Then ptsPer15 = 1
    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Время, чч:мм:сс"
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        .TickMarkSpacing = ptsPer15
        .TickLabelSpacing = ptsPer15
        .TickLabels.Orientation = 90
    End With
    With ch.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = yTitle
        .HasMajorGridlines = True
    End With
    With ch.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = y2Title
    End With
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    On Error GoTo 0
End Sub

Public Sub GetRangeMinMaxByCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
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

Public Function MaxAbsInCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
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

Public Sub WriteChartVerdictBlock(ByVal wsChart As Worksheet, ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings)
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

    wsChart.Range("A10").Value = "tр / применимость:"
    wsChart.Range("B10").Value = IIf(IsDate(res.TrTime), Format(res.TrTime, "hh:mm:ss"), "-") & " (" & Format(res.TrSec, "0") & " с, " & _
                                 IIf(res.TrQuasiSteady, "квазиуст.", "переходный") & "); " & res.Applicability
    wsChart.Range("A11").Value = "Тип реакции / ТИ:"
    wsChart.Range("B11").Value = IIf(Len(res.ReactionType) > 0, res.ReactionType & " (п. 4.4.1)", "-") & _
                                 IIf(Len(res.QualClauseVerdict) > 0, "; " & res.QualClauseVerdict, "") & _
                                 IIf(Len(res.TelemetryNote) > 0, "; " & res.TelemetryNote, "") & _
                                 IIf(Len(res.OscNote) > 0, "; " & res.OscNote, "") & _
                                 IIf(res.H15 > -900, "; h(15 с) = " & Format(res.H15, "0.00"), "")

    wsChart.Range("A12").Value = "Вывод:"
    wsChart.Range("B12").Value = res.Verdict

    wsChart.Range("A3:A12").Font.Bold = True
    wsChart.Range("D3:D9").Font.Bold = True
    wsChart.Range("A3:E12").NumberFormat = "@"
    If res.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
        wsChart.Range("B3").Interior.Color = RGB(255, 235, 156)
        wsChart.Range("E7").Font.Color = RGB(156, 87, 0)
    ElseIf Not res.QuantPass And res.QuantApplicable Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If Not res.QuantApplicable Then wsChart.Range("B3").Value = "Н/П (п. 4.1.2) - справочно: " & quantCell
    If Not res.QualApplicable Then wsChart.Range("B4").Value = "Н/П (п. 4.1.3)"
    If Not res.QualPass And res.QualApplicable Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If res.TransientType = "Колебательный" Or res.OscViolation Then wsChart.Range("B5").Font.Color = RGB(192, 0, 0)
    If res.React15Status = "Нарушение" Then wsChart.Range("B9").Font.Color = RGB(192, 0, 0)
    If Left$(res.QualClauseVerdict, 9) = "НАРУШЕНИЕ" Then wsChart.Range("B11").Font.Color = RGB(192, 0, 0)
    If Left$(res.Verdict, 9) = "НАРУШЕНИЕ" Then
        wsChart.Range("B12").Font.Color = RGB(192, 0, 0)
    ElseIf Left$(res.Verdict, 2) = "ОК" Then
        wsChart.Range("B12").Font.Color = RGB(0, 97, 0)
    Else
        wsChart.Range("B12").Font.Color = RGB(89, 89, 89)
    End If
End Sub

Public Function FindChartEndRow(ByVal wsData As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
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

Public Sub WriteStationChartSheet(ByVal stationName As String, ByVal paropipeFilter As String, _
                                   ByVal lastDataRow As Long, ByRef st As TSettings, _
                                   ByVal fnchMin As Double, ByVal dFmax As Double, ByRef sr As TStationResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataName As String, chartName As String
    Dim chartObj As ChartObject
    Dim endRow As Long, startRow As Long
    Dim ySpan As Double

    dataName = sr.SheetName
    chartName = sr.ChartSheetName

    Set wsData = ThisWorkbook.Worksheets(dataName)
    Set wsChart = EnsureSheet(chartName)
    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    wsChart.Range("A1").Value = "График ОПРЧ (сумма): " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "") & _
                                IIf(sr.EventIdx > 1, " (событие " & sr.EventIdx & ")", "")
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
    wsChart.Range("E5").Value = CStr(sr.Count) & IIf(sr.IsNeblok, " (неблочная часть)", " (блочная компоновка)")
    wsChart.Range("A6").Value = "tр / применимость:"
    wsChart.Range("B6").Value = IIf(IsDate(sr.TrTime), Format(sr.TrTime, "hh:mm:ss"), "-") & " (" & Format(sr.TrSec, "0") & " с, " & _
                                IIf(sr.TrQuasiSteady, "квазиуст.", "переходный") & "); " & sr.Applicability
    wsChart.Range("A7").Value = "Вывод:"
    wsChart.Range("B7").Value = sr.Verdict
    wsChart.Range("A3:A7").Font.Bold = True
    wsChart.Range("D3:D5").Font.Bold = True
    wsChart.Range("A3:E7").NumberFormat = "@"
    If sr.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
    ElseIf Not sr.QuantPass And sr.QuantApplicable Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If sr.React15Status = "Нарушение" Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then
        wsChart.Range("B7").Font.Color = RGB(192, 0, 0)
    ElseIf Left$(sr.Verdict, 2) = "ОК" Then
        wsChart.Range("B7").Font.Color = RGB(0, 97, 0)
    Else
        wsChart.Range("B7").Font.Color = RGB(89, 89, 89)
    End If

    startRow = 11
    If lastDataRow < startRow Then Exit Sub
    endRow = FindChartEndRow(wsData, startRow, lastDataRow, wsData.Cells(startRow, 1).Value, st.ChartIntervalSec)
    If endRow < startRow Then endRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 130, 1020, 520)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Суммарный мониторинг ОПРЧ: " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "") & _
                           IIf(sr.EventIdx > 1, " (событие " & sr.EventIdx & ")", "")
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
    ApplyChartAxesFormat chartObj.Chart, wsData, startRow, endRow, "dPсум, МВт", "f, Гц"
End Sub

' ==========================================================
' Summary / оформление / валидация / лог
' ==========================================================

Public Sub AddSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                      ByVal xCol As Long, ByVal yCol As Long, ByVal titleText As String, ByVal secondaryAxis As Boolean)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = titleText
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    If secondaryAxis Then s.AxisGroup = xlSecondary
End Sub

Public Sub AddMarkerSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
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

Public Sub AddLimitLineSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
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

Public Sub LoadChartDisplayOptions(ByVal wsCfg As Worksheet)
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

Public Function ChartOpt(ByVal key As String, ByVal fallback As Boolean) As Boolean
    If m_ChartOpts Is Nothing Then
        ChartOpt = fallback
    ElseIf m_ChartOpts.Exists(key) Then
        ChartOpt = CBool(m_ChartOpts(key))
    Else
        ChartOpt = fallback
    End If
End Function

Public Sub EnsureChartOptionCheckboxes(ByVal ws As Worksheet)
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
