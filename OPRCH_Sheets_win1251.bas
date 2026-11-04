Attribute VB_Name = "OPRCH_Sheets"
Option Explicit


Public Sub AddProtocolLine(ByVal s As String)
    If m_Protocol Is Nothing Then Set m_Protocol = New Collection
    m_Protocol.Add s
End Sub

Public Sub AddProtocolGenerator(ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim hdr As String
    hdr = g.Station & " / " & g.Generator & " (" & g.EquipType & ", Pном " & Format(g.PNom, "0.0") & " МВт, S " & Format(g.SPct, "0.0") & _
          " %, fнч " & Format(g.Fnch, "0.000") & " Гц, " & IIf(g.HasRM, "с РМ", "без РМ") & ", " & g.Layout & ")"
    If res.EventIdx > 1 Then hdr = hdr & " - событие " & res.EventIdx
    AddProtocolLine hdr
    AddProtocolLine "  Старт t0: " & Format(res.StartTime, "dd.mm.yyyy hh:mm:ss") & "; выход за fнч: " & _
        IIf(IsDate(res.FirstExceedTime), Format(res.FirstExceedTime, "hh:mm:ss"), "-") & "; tр: " & _
        IIf(IsDate(res.TrTime), Format(res.TrTime, "hh:mm:ss"), "-") & " (" & Format(res.TrSec, "0") & " с, " & _
        IIf(res.TrQuasiSteady, "квазиустановившийся режим", "переходный режим") & ")"
    AddProtocolLine "  Применимость: " & res.Applicability
    AddProtocolLine "  dF(tр) = " & Format(res.Dfr + IIf(res.Dfr > 0, g.Fnch, IIf(res.Dfr < 0, -g.Fnch, 0)), "0.000") & " Гц; dFр = " & Format(res.Dfr, "0.000") & _
        " Гц; Kд(tр) = " & Format(res.KdUsedQuant, "0.00") & "; P0 = " & Format(res.P0, "0.000") & " МВт; Pтек(tр) = " & Format(res.PTek, "0.000") & _
        " МВт; Pтреб = " & Format(res.PReq, "0.000") & " МВт" & IIf(res.Limited, " (ограничен " & res.LimitType & ")", "") & _
        "; Pфакт = " & Format(res.PFact, "0.000") & " МВт; выполнение " & Format(res.QuantPct, "0") & " %"
    AddProtocolLine "  Реакция РЧВ 15 с (п. 4.2.11): " & res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт)" & _
        "; дефицит (п. 4.2.14.1): " & IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.0") & " %Pном, " & Format(res.DeficitSec, "0") & " с", "нет") & _
        "; тип реакции (п. 4.4.1): " & IIf(Len(res.ReactionType) > 0, res.ReactionType, "-")
    AddProtocolLine "  ВЫВОД: " & res.Verdict
    AddProtocolLine ""
End Sub

Public Sub AddProtocolStation(ByRef sr As TStationResult)
    AddProtocolLine sr.Station & IIf(Len(sr.Paroprovod) > 0, " / паропровод " & sr.Paroprovod, "") & " - СУММА (" & sr.Count & " ген., " & _
        IIf(sr.IsNeblok, "неблочная часть, п. 4.2.10", "блочная компоновка - справочно") & ")" & IIf(sr.EventIdx > 1, " - событие " & sr.EventIdx, "")
    AddProtocolLine "  Старт t0: " & Format(sr.StartTime, "dd.mm.yyyy hh:mm:ss") & "; tр: " & IIf(IsDate(sr.TrTime), Format(sr.TrTime, "hh:mm:ss"), "-") & _
        " (" & Format(sr.TrSec, "0") & " с); P0 = " & Format(sr.P0, "0.000") & " МВт; Pтреб_сум = " & Format(sr.PReq, "0.000") & _
        " МВт; Pфакт_сум = " & Format(sr.PFact, "0.000") & " МВт; выполнение " & Format(sr.QuantPct, "0") & " %"
    AddProtocolLine "  Применимость: " & sr.Applicability
    AddProtocolLine "  Реакция РЧВ 15 с: " & sr.React15Status & " (dP15 = " & Format(sr.DP15, "0.000") & " МВт); дефицит: " & _
        IIf(sr.DeficitSec > 0, "до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум, " & Format(sr.DeficitSec, "0") & " с", "нет")
    AddProtocolLine "  ВЫВОД: " & sr.Verdict
    AddProtocolLine ""
End Sub

Public Sub WriteProtocolSheet(ByRef st As TSettings, ByVal wsRaw As Worksheet)
    Dim ws As Worksheet, r As Long, i As Long, lastR As Long
    Dim tFirst As Variant, tLast As Variant
    Set ws = EnsureSheet(SH_PROT)
    ws.Cells.Clear
    lastR = LastUsedRow(wsRaw)
    tFirst = wsRaw.Cells(2, 1).Value
    tLast = wsRaw.Cells(lastR, 1).Value
    ws.Range("A1").Value = "ПРОТОКОЛ оценки участия генерирующего оборудования в ОПРЧ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Методика: Методические указания по оценке участия генерирующего оборудования в ОПРЧ (Приложение 14 к Техтребованиям); " & _
                           "ступени динамики - Приложение 1 (МУ проверки готовности)."
    ws.Range("A3").Value = "Программа: OPRCH_Monitor v" & OPRCH_VERSION & "; режим: " & st.Mode & "; дата расчёта: " & Format(Now, "dd.mm.yyyy hh:mm")
    ws.Range("A4").Value = "Выборка RawData: " & IIf(IsDate(tFirst), Format(tFirst, "dd.mm.yyyy hh:mm:ss"), "-") & " - " & _
                           IIf(IsDate(tLast), Format(tLast, "dd.mm.yyyy hh:mm:ss"), "-") & " (" & (lastR - 1) & " строк); fном = " & st.FNom & " Гц; " & _
                           "контрольный интервал " & st.QuantIntervalSec & " с; допуск колич. " & st.QuantTolPct & " %; ограничение Pтреб " & st.MaxReqPct & " %Pном"
    ws.Range("A5").Value = IIf(st.IsEmergency, _
        "Режим «Авария»: для оборудования без РМ приняты S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц; применимость оценок - по п. 4.1.2/4.1.3 МУ.", _
        "Режим «Испытание»: использованы фактические S/fнч из Config; применимость по п. 4.1.2/4.1.3 не проверялась.")
    ws.Range("A6").Value = "Нарушения формулируются со ссылкой на пункты МУ: 4.2.11 (реакция РЧВ 15 с), 4.2.14.1 (дефицит первичной мощности), " & _
                           "4.2.14.4 (колебания), 4.4.3/4.4.4 (отсутствие «адекватной» реакции без РМ / с РМ), 4.2.10 (оценка неблочной части по сумме)."
    r = 8
    If Not m_Protocol Is Nothing Then
        For i = 1 To m_Protocol.Count
            ws.Cells(r, 1).Value = CStr(m_Protocol(i))
            If Left$(CStr(m_Protocol(i)), 2) <> "  " And Len(CStr(m_Protocol(i))) > 0 Then ws.Cells(r, 1).Font.Bold = True
            If InStr(CStr(m_Protocol(i)), "ВЫВОД: НАРУШЕНИЕ") > 0 Then ws.Cells(r, 1).Font.Color = RGB(192, 0, 0)
            If InStr(CStr(m_Protocol(i)), "ВЫВОД: ОК") > 0 Then ws.Cells(r, 1).Font.Color = RGB(0, 97, 0)
            r = r + 1
        Next i
    End If
    ws.Cells(r + 1, 1).Value = "Не автоматизировано (требует ручной проверки): срабатывание технологических защит / вмешательство персонала (п. 4.2.14.1, абз. 2), " & _
                               "влияние команд диспетчера и АРЧМ (п. 4.2.9), правило трёх фактов за год для оборудования без РМ (п. 4.4.3), " & _
                               "заявки на вывод из ОПРЧ / пуски и остановы (п. 4.1.6)."
    ws.Columns("A").ColumnWidth = 160
    ws.Range("A1:A" & (r + 1)).WrapText = True
    ws.Range("A1:A" & (r + 1)).VerticalAlignment = xlTop
End Sub


Public Sub WriteGeneratorSheet(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
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
    pCol = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
    fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)

    shName = GeneratorSheetName(g, res.EventIdx)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

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
    ws.Range("A6:B6").Value = Array("Колич. статус", IIf(res.QuantApplicable, quantStatusStr, "Н/П (п. 4.1.2)"))
    ws.Range("A7:B7").Value = Array("Кач. статус", IIf(res.QualApplicable, IIf(res.QualPass, "ОК", "Нарушение"), "Н/П (п. 4.1.3)"))
    ws.Range("A8:B8").Value = Array("Характер", res.TransientType)
    ws.Range("A9:B9").Value = Array("Амплитуда, %Pном", Round(res.AmplPctPnom, 2))
    ws.Cells(1, 4).Resize(1, 2).Value = Array("P0, МВт", res.P0)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pтек(tр), МВт", res.PTek)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Pуст_сред, МВт", res.PsteadyAvg)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("dF макс, Гц", res.Df)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("dFr(tр), Гц", res.Dfr)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Pтреб(tр), МВт", res.PReq)
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Pфакт(tр), МВт", res.PFact)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Колич. %", res.QuantPct)
    ws.Cells(9, 4).Resize(1, 2).Value = Array("Экстремумов", res.NumExtrema)

    ws.Cells(1, 24).Resize(1, 2).Value = Array("Событие №", res.EventIdx)
    ws.Cells(2, 24).Resize(1, 2).Value = Array("tр (п. 4.2.5)", res.TrTime)
    ws.Cells(3, 24).Resize(1, 2).Value = Array("tр, с от старта", res.TrSec)
    ws.Cells(4, 24).Resize(1, 2).Value = Array("Режим в tр", IIf(res.TrQuasiSteady, "квазиустановившийся (±15 с, 10 мГц)", "переходный"))
    ws.Cells(5, 24).Resize(1, 2).Value = Array("Применимость", res.Applicability)
    ws.Cells(6, 24).Resize(1, 2).Value = Array("Тип реакции (п. 4.4.1)", IIf(Len(res.ReactionType) > 0, res.ReactionType, "-"))
    ws.Cells(7, 24).Resize(1, 2).Value = Array("Колебания (п. 4.2.14.4)", IIf(Len(res.OscNote) > 0, res.OscNote, "нет"))
    ws.Cells(8, 24).Resize(1, 2).Value = Array("Качество ТИ", IIf(Len(res.TelemetryNote) > 0, res.TelemetryNote, "ОК"))
    ws.Cells(9, 24).Resize(1, 2).Value = Array("h(15 с), Прил. п. 6.2", IIf(res.H15 > -900, Round(res.H15, 3), res.H15Note))
    ws.Cells(10, 24).Resize(1, 2).Value = Array("Pтреб/Pфакт (Tинт), МВт", Format(res.PReqRef, "0.000") & " / " & Format(res.PFactRef, "0.000"))
    ws.Cells(11, 24).Resize(1, 2).Value = Array("Режим / РМ", st.Mode & " / " & IIf(g.HasRM, "с РМ", "без РМ") & IIf(g.ForcedNoRM, " (S=6 %, fнч=0.15 Гц)", "") & " / " & g.Layout)

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

    ws.Range("A11:W11").Value = Array( _
        "Время", "Частота, Гц", "P, МВт", "dPфакт, МВт", "Pтреб_накоп, МВт", "dFr, Гц", _
        "Маркер tр", "Уровень +допуск", "Уровень -допуск", _
        "Маркер t5", "Маркер t10", "Маркер выхода за fнч", _
        "dPmax", "dPmin", _
        "Pтреб_абс, МВт", "Уровень +допуск_абс", "Уровень -допуск_абс", _
        "Pmax, МВт", "Pmin, МВт", _
        "Маркер t5_абс", "Маркер t10_абс", "Маркер fнч_абс", "Маркер tр_абс" _
    )

    If st.PreBufferSec > 0 Then
        displayStartRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, -st.PreBufferSec)
        If displayStartRow < 2 Then displayStartRow = 2
        If displayStartRow > res.StartRow Then displayStartRow = res.StartRow
    Else
        displayStartRow = res.StartRow
    End If
    endRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, MaxD(st.QuantIntervalSec, g.T10Sec))
    If res.EventEndRow > endRow Then endRow = res.EventEndRow
    If res.TrRow > endRow Then endRow = res.TrRow
    outR = 12

    signReq = SgnNZ(res.PReq)
    targetPreq = res.PReq
    tolPreq = g.PNom * 0.01
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

    Dim preqStep As Double, f0g As Double
    f0g = res.F0
    If f0g <= 0 Then f0g = st.FNom
    For r = displayStartRow To endRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        dFr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, f0g) - f0g, g.Fnch)
        preqStep = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, _
                   SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)), dFr, st.MaxReqPct), _
                   res.ReservePlus, res.ReserveMinus)

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

    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 10, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 11, markerSpan, -markerSpan
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 12, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 20, markerHiAbs, markerLoAbs
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 21, markerHiAbs, markerLoAbs
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 22, markerHiAbs, markerLoAbs
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.TrTime, 7, markerSpan, -markerSpan
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.TrTime, 23, markerHiAbs, markerLoAbs

    chartEndRow = outR - 1

    ws.Columns("A:W").AutoFit
    ApplyGeneratorSheetFormats ws
    ws.Columns("X").ColumnWidth = 24
    ws.Columns("Y").ColumnWidth = 60
    ws.Cells(10, 1).Value = "Вывод"
    ws.Cells(10, 2).Value = res.Verdict
    ws.Cells(10, 1).Font.Bold = True
End Sub

Public Sub BuildStationAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
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
        g = ReadGenCfgEff(wsCfg, r, st)
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

Public Sub BuildParoprovodSubAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                         ByVal stationName As String, ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long
    Dim g As TGenCfg
    Dim paropipes As Object, p As Variant

    Set paropipes = CreateObject("Scripting.Dictionary")
    cfgLast = LastUsedRow(wsCfg)
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfgEff(wsCfg, r, st)
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


Public Function SumPowerAtRow(ByVal wsRaw As Worksheet, ByRef pCols() As Long, ByVal cnt As Long, _
                               ByVal r As Long, ByVal thresholdMW As Double) As Double
    Dim i As Long, v As Double, total As Double
    total = 0
    For i = 1 To cnt
        v = NzD(wsRaw.Cells(r, pCols(i)).Value, 0)
        If v > thresholdMW Then total = total + v
    Next i
    SumPowerAtRow = total
End Function

Public Sub BuildOneStationAggregate(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long, cnt As Long
    Dim g As TGenCfg
    Dim pCols() As Long, fnchArr() As Double, sArr() As Double, pnomArr() As Double, etArr() As String
    Dim pmaxArr() As Double, pminArr() As Double
    Dim freqCol As Long, timeCol As Long
    Dim isNeblok As Boolean, anyHasRM As Boolean
    Dim evList() As TEvent, nEv As Long, k As Long
    Dim fRef As Double

    cfgLast = LastUsedRow(wsCfg)
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Exit Sub

    cnt = 0
    isNeblok = False
    anyHasRM = False
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextCfg
        g = ReadGenCfgEff(wsCfg, r, st)
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
                    pCols(cnt) = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
                    If pCols(cnt) = 0 Then cnt = cnt - 1: GoTo NextCfg
                    fnchArr(cnt) = g.Fnch
                    sArr(cnt) = g.SPct
                    pnomArr(cnt) = g.PNom
                    etArr(cnt) = g.EquipType
                    pmaxArr(cnt) = g.PMax
                    pminArr(cnt) = g.PMin
                    If UCase$(g.Layout) = "НЕБЛОК" Then isNeblok = True
                    If g.HasRM Then anyHasRM = True
                    If freqCol = 0 Then freqCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
                End If
            End If
        End If
NextCfg:
    Next r
    If cnt = 0 Then Exit Sub
    If freqCol = 0 Then Exit Sub

    If anyHasRM Then
        fRef = st.FNom
    Else
        fRef = NzD(wsRaw.Cells(2, freqCol).Value, st.FNom)
    End If
    nEv = FindEvents(wsRaw, timeCol, freqCol, st, MinArray(fnchArr), fRef, evList)
    For k = 1 To nEv
        BuildOneStationAggregateEvent wsRaw, wsSummary, stationName, paropipeFilter, st, outRow, _
                                      pCols, fnchArr, sArr, pnomArr, etArr, pmaxArr, pminArr, cnt, _
                                      freqCol, timeCol, isNeblok, anyHasRM, evList(k), k
    Next k
End Sub

Public Sub BuildOneStationAggregateEvent(ByVal wsRaw As Worksheet, ByVal wsSummary As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings, ByRef outRow As Long, _
                                      ByRef pCols() As Long, ByRef fnchArr() As Double, ByRef sArr() As Double, _
                                      ByRef pnomArr() As Double, ByRef etArr() As String, _
                                      ByRef pmaxArr() As Double, ByRef pminArr() As Double, ByVal cnt As Long, _
                                      ByVal freqCol As Long, ByVal timeCol As Long, ByVal isNeblok As Boolean, _
                                      ByVal anyHasRM As Boolean, ByRef ev As TEvent, ByVal evIdx As Long)
    Dim r As Long
    Dim shName As String, ws As Worksheet
    Dim startRow As Long, endRow As Long, rowQ As Long, firstExceedRow As Long, evEnd As Long
    Dim p0 As Double, pNow As Double, preq As Double, preqOrig As Double
    Dim pfact As Double, dF As Double
    Dim i As Long, dfr As Double
    Dim suffix As String
    Dim pMaxSum As Double, pMinSum As Double, pNomSum As Double
    Dim reservePlus As Double, reserveMinus As Double
    Dim limited As Boolean, limitType As String
    Dim sr As TStationResult
    Dim srcLabel As String
    Dim f0 As Double

    srcLabel = stationName & IIf(Len(paropipeFilter) > 0, "/" & paropipeFilter, "") & " (сумма)" & IIf(evIdx > 1, " событие " & evIdx, "")

    startRow = ev.StartRow
    firstExceedRow = ev.FirstExceedRow
    evEnd = ev.EndRow
    If evEnd < startRow Then evEnd = startRow
    endRow = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)
    f0 = ResolveF0(wsRaw, freqCol, startRow, anyHasRM, st.FNom)
    sr.F0 = f0

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
    Dim rA As Long, r1 As Long, sumA As Double, cntA As Long
    If st.AvgWindowSec > 0 Then
        r1 = RowByTimeOffset(wsRaw, timeCol, startRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To startRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then p0 = sumA / cntA
    End If

    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    shName = StationSheetName(stationName, paropipeFilter, evIdx)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear

    ws.Range("A10:H10").Value = Array("Время", "Частота, Гц", "Pсум, МВт", "dPсум, МВт", "Pтреб_сум, МВт", _
                                       "dPmax_сум", "dPmin_сум", "Дефицит_сум, МВт")
    rowQ = 11
    Dim dPmaxRel As Double, dPminRel As Double
    Dim defMaxPct As Double, defSec As Double, tolSum As Double
    Dim preqStep As Double, preqCap As Double, allowed As Double, have As Double, deficit As Double, dt As Double
    Dim tSecAgg As Double, fMean As Double, isQS As Boolean
    Dim trRow As Long, trDef As Double, anyQS As Boolean, trDfrRow As Long, trDfrAbs As Double
    Dim run02 As Double, best02 As Double, dfAbs As Double
    reservePlus = pMaxSum - p0
    reserveMinus = p0 - pMinSum
    If reservePlus < 0 Then reservePlus = 0
    If reserveMinus < 0 Then reserveMinus = 0
    dPmaxRel = pMaxSum - p0
    dPminRel = pMinSum - p0
    tolSum = 0.01 * pNomSum
    defMaxPct = 0
    defSec = 0
    trRow = 0: trDef = -1E+99: anyQS = False: trDfrRow = 0: trDfrAbs = 0
    run02 = 0: best02 = 0
    dF = 0
    For r = startRow To evEnd
        ws.Cells(rowQ, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(rowQ, 2).Value = wsRaw.Cells(r, freqCol).Value

        pNow = SumPowerAtRow(wsRaw, pCols, cnt, r, st.WorkThresholdMW)
        ws.Cells(rowQ, 3).Value = pNow
        ws.Cells(rowQ, 4).Value = pNow - p0

        dfAbs = NzD(wsRaw.Cells(r, freqCol).Value, f0) - f0
        If Abs(dfAbs) > Abs(dF) Then dF = dfAbs
        If r < evEnd Then
            dt = SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(r + 1, timeCol).Value)
        Else
            dt = 1
        End If
        If dt <= 0 Or dt > 60 Then dt = 1
        If Abs(dfAbs) >= QUANT_DF_HZ - 0.0000001 Then
            run02 = run02 + dt
            If run02 > best02 Then best02 = run02
        Else
            run02 = 0
        End If

        preqStep = 0
        tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        For i = 1 To cnt
            If inWork(i) Then
                dfr = DeadbandDeviation(dfAbs, fnchArr(i))
                preqStep = preqStep + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), tSecAgg), dfr, st.MaxReqPct)
            End If
        Next i
        ws.Cells(rowQ, 5).Value = preqStep
        ws.Cells(rowQ, 6).Value = dPmaxRel
        ws.Cells(rowQ, 7).Value = dPminRel

        deficit = 0
        preqCap = ClampReserve(preqStep, reservePlus, reserveMinus)
        allowed = Abs(preqCap) - tolSum
        If allowed > 0 And tSecAgg >= 15 Then
            have = SgnNZ(preqCap) * (pNow - p0)
            deficit = allowed - have
            If deficit < 0 Then deficit = 0
        End If
        ws.Cells(rowQ, 8).Value = deficit
        If deficit > 0 And pNomSum > 0 Then
            If 100# * deficit / pNomSum > defMaxPct Then defMaxPct = 100# * deficit / pNomSum
            defSec = defSec + dt
        End If

        If tSecAgg >= 15 And preqCap <> 0 Then
            have = SgnNZ(preqCap) * (pNow - p0)
            If (Abs(preqCap) - tolSum) - have > trDef Then
                If Not anyQS Then
                    If (Abs(preqCap) - tolSum) - have > trDef Then trDef = (Abs(preqCap) - tolSum) - have: trRow = r
                End If
            End If
            isQS = IsQuasiSteady(wsRaw, freqCol, timeCol, r, f0, fMean)
            If isQS Then
                If DeadbandDeviation(fMean - f0, MinArray(fnchArr)) <> 0 Then
                    anyQS = True
                    have = SgnNZ(preqCap) * (pNow - p0)
                    If (Abs(preqCap) - tolSum) - have > trDef Then trDef = (Abs(preqCap) - tolSum) - have: trRow = r
                    If Abs(fMean - f0) > trDfrAbs Then trDfrAbs = Abs(fMean - f0): trDfrRow = r
                End If
            ElseIf Abs(dfAbs) > trDfrAbs Then
                trDfrAbs = Abs(dfAbs): trDfrRow = r
            End If
        End If
        rowQ = rowQ + 1
    Next r

    Dim trQS As Boolean
    If anyQS Then
        If trDef > 0 Then
        Else
            trRow = IIf(trDfrRow > 0, trDfrRow, trRow)
        End If
        trQS = True
    ElseIf trRow > 0 Then
        trQS = False
    Else
        trRow = endRow
        trQS = False
    End If
    If trRow < startRow Then trRow = startRow
    If trRow > evEnd And anyQS Then trRow = evEnd
    tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(trRow, timeCol).Value)

    If st.AvgWindowSec > 0 Then
        sumA = 0: cntA = 0
        r1 = RowByTimeOffset(wsRaw, timeCol, trRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To trRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then pNow = sumA / cntA Else pNow = SumPowerAtRow(wsRaw, pCols, cnt, trRow, st.WorkThresholdMW)
    Else
        pNow = SumPowerAtRow(wsRaw, pCols, cnt, trRow, st.WorkThresholdMW)
    End If
    pfact = pNow - p0

    Dim dfTr As Double
    If trQS Then
        dfTr = MeanFreqAround(wsRaw, freqCol, timeCol, trRow, QS_HALF_WINDOW_SEC, f0) - f0
    Else
        dfTr = NzD(wsRaw.Cells(trRow, freqCol).Value, f0) - f0
    End If
    preqOrig = 0
    For i = 1 To cnt
        If inWork(i) Then
            dfr = DeadbandDeviation(dfTr, fnchArr(i))
            preqOrig = preqOrig + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), tSecAgg), dfr, st.MaxReqPct)
        End If
    Next i
    preq = preqOrig
    limited = False
    limitType = ""
    If preq > reservePlus Then
        preq = reservePlus: limited = True: limitType = "Pmax"
    ElseIf preq < -reserveMinus Then
        preq = -reserveMinus: limited = True: limitType = "Pmin"
    End If
    If limited Then
        AppendLog "INFO", srcLabel, _
                  "Pтреб_сум ограничен " & limitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

    Dim qpct As Double, qpass As Boolean, dfrMin As Double
    dfrMin = DeadbandDeviation(dfTr, MinArray(fnchArr))
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

    If st.IsEmergency Then
        sr.QuantApplicable = (best02 >= QUANT_MIN_SEC_TES - 0.5)
        If sr.QuantApplicable Then
            sr.Applicability = "колич. (п. 4.1.2: |dF| >= 0.2 Гц в течение " & Format(best02, "0") & " с)"
        Else
            sr.Applicability = "колич. оценка не применяется (п. 4.1.2: |dF| >= 0.2 Гц " & Format(best02, "0") & " с < 30 с; |dF|max = " & Format(Abs(dF), "0.000") & " Гц)"
        End If
    Else
        sr.QuantApplicable = True
        sr.Applicability = "Испытание: применимость по п. 4.1.2 не проверяется"
    End If

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
            dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, freqCol).Value, f0) - f0, fnchArr(i))
            If dfr15 <> 0 Then
                anyOutside = True
                preq15Sum = preq15Sum + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), 15#), dfr15, st.MaxReqPct)
            End If
        End If
    Next i
    preq15Sum = ClampReserve(preq15Sum, reservePlus, reserveMinus)
    Dim needSignSum As Integer
    needSignSum = SgnNZ(preqOrig)
    If needSignSum = 0 Then needSignSum = -SgnNZ(DeadbandDeviation(dF, MinArray(fnchArr)))
    ClassifyReaction15 st, dp15, pNomSum, needSignSum, reservePlus, reserveMinus, _
                       React15ThresholdMW(st, pNomSum, preq15Sum), (Not anyOutside), r15Status, r15Pass, r15Note
    If r15Status = "Нарушение" Then AppendLog "WARN", srcLabel, "п. 4.2.11: " & r15Note

    ws.Range("A1:B1").Value = Array("Станция", stationName & IIf(evIdx > 1, " (событие " & evIdx & ")", ""))
    If Len(paropipeFilter) > 0 Then
        ws.Range("A2:B2").Value = Array("Паропровод", paropipeFilter)
    Else
        ws.Range("A2:B2").Value = Array("Режим", IIf(isNeblok, "Суммарная нагрузка неблочной части (п. 4.2.10 МУ)", "Суммарная нагрузка (блочная компоновка - справочно)"))
    End If
    ws.Range("A3:B3").Value = Array("Старт (расч.)", wsRaw.Cells(startRow, timeCol).Value)
    If firstExceedRow > 0 Then
        ws.Range("A4:B4").Value = Array("Выход за fнч", wsRaw.Cells(firstExceedRow, timeCol).Value)
    Else
        ws.Range("A4:B4").Value = Array("Выход за fнч", "")
    End If
    ws.Range("A5:B5").Value = Array("P0, МВт", p0)
    ws.Range("A6:B6").Value = Array("Pтек(tр), МВт", pNow)
    ws.Range("A7:B7").Value = Array("Pтреб(tр), МВт", preq)
    ws.Range("A8:B8").Value = Array("Pфакт(tр), МВт", pfact)
    ws.Range("A9:B9").Value = Array("tр", wsRaw.Cells(trRow, timeCol).Value)

    ws.Cells(1, 4).Resize(1, 2).Value = Array("Pmax_сум, МВт", pMaxSum)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pmin_сум, МВт", pMinSum)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Резерв '+', МВт", reservePlus)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("Резерв '-', МВт", reserveMinus)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("Pтреб исх., МВт", preqOrig)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Ограничение", IIf(limited, "Да (" & limitType & ")", "нет"))
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Генераторов в сумме", cnt)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Pном_сум (в работе), МВт", pNomSum)
    ws.Cells(9, 4).Resize(1, 2).Value = Array("tр, с / режим", Format(tSecAgg, "0") & " с / " & IIf(trQS, "квазиустановившийся", "переходный"))

    sr.Station = stationName
    sr.Paroprovod = paropipeFilter
    sr.Count = cnt
    sr.EventIdx = evIdx
    sr.IsNeblok = isNeblok
    sr.StartTime = wsRaw.Cells(startRow, timeCol).Value
    If firstExceedRow > 0 Then sr.FirstExceedTime = wsRaw.Cells(firstExceedRow, timeCol).Value Else sr.FirstExceedTime = ""
    sr.TrTime = wsRaw.Cells(trRow, timeCol).Value
    sr.TrSec = tSecAgg
    sr.TrQuasiSteady = trQS
    sr.P0 = p0
    sr.PTek = pNow
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
    sr.ChartSheetName = StationChartSheetName(stationName, paropipeFilter, evIdx)
    sr.Verdict = BuildStationVerdict(sr, st)

    ws.Cells(1, 7).Resize(1, 2).Value = Array("Колич. %", qpct)
    ws.Cells(2, 7).Resize(1, 2).Value = Array("Колич. статус", IIf(Not sr.QuantApplicable, "Н/П (п. 4.1.2)", IIf(limited, "Ограничен " & limitType, IIf(qpass, "ОК", "Нарушение"))))
    ws.Cells(3, 7).Resize(1, 2).Value = Array("Реакция 15 с (п. 4.2.11)", r15Status & " (dP15 = " & Format(dp15, "0.000") & " МВт)")
    ws.Cells(4, 7).Resize(1, 2).Value = Array("Дефицит (п. 4.2.14.1)", _
        IIf(defSec > 0, "до " & Format(defMaxPct, "0.00") & " %Pном_сум, " & Format(defSec, "0") & " с", "нет"))
    ws.Cells(5, 7).Resize(1, 2).Value = Array("Применимость", sr.Applicability)

    ws.Columns("A:H").AutoFit
    ws.Cells(6, 7).Value = "Вывод"
    ws.Cells(6, 8).Value = sr.Verdict
    ws.Cells(6, 7).Font.Bold = True
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then ws.Cells(6, 8).Font.Color = RGB(192, 0, 0) Else ws.Cells(6, 8).Font.Color = RGB(0, 97, 0)

    ws.Range("A1:A9").NumberFormat = "@"
    ws.Range("B1:B2").NumberFormat = "@"
    ws.Range("B3:B4").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B5:B8").NumberFormat = "0.000"
    ws.Range("B9").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("D1:D9").NumberFormat = "@"
    ws.Range("E1:E5").NumberFormat = "0.000"
    ws.Range("E6").NumberFormat = "@"
    ws.Range("E7").NumberFormat = "0"
    ws.Range("E8").NumberFormat = "0.000"
    ws.Range("E9").NumberFormat = "@"
    ws.Range("G1:G6").NumberFormat = "@"
    ws.Range("H1").NumberFormat = "0.0"
    ws.Range("H2:H6").NumberFormat = "@"
    ws.Range("A10:H10").NumberFormat = "@"
    ws.Range("A11:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B11:H100000").NumberFormat = "0.000"

    AppendLog "INFO", srcLabel, sr.Verdict

    WriteStationChartSheet stationName, paropipeFilter, rowQ - 1, st, MinArray(fnchArr), dF, sr

    WriteSummaryStationRow wsSummary, outRow, sr
    AddProtocolStation sr
    outRow = outRow + 1

    If Len(paropipeFilter) = 0 And isNeblok And evIdx = 1 Then
        AnnotateGeneratorRowsWithStationVerdict wsSummary, outRow - 1, stationName, (Left$(sr.Verdict, 2) = "ОК")
    End If
End Sub

Public Sub AnnotateGeneratorRowsWithStationVerdict(ByVal ws As Worksheet, ByVal lastRow As Long, _
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


Public Function StationSheetName(ByVal stationName As String, ByVal paropipeFilter As String, ByVal evIdx As Long) As String
    Dim suffix As String, base As String
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    suffix = suffix & EventSuffix(evIdx)
    base = stationName
    If Len(base & suffix) > 31 Then base = Left$(base, 31 - Len(suffix))
    StationSheetName = MakeSheetName(base & suffix)
End Function

Public Function StationChartSheetName(ByVal stationName As String, ByVal paropipeFilter As String, ByVal evIdx As Long) As String
    Dim suffix As String, dataName As String, chartName As String, base As String
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    suffix = suffix & EventSuffix(evIdx)
    dataName = StationSheetName(stationName, paropipeFilter, evIdx)
    base = stationName
    If Len(base & suffix & CHART_SUFFIX) > 31 Then base = Left$(base, 31 - Len(suffix) - Len(CHART_SUFFIX))
    chartName = MakeSheetName(base & suffix & CHART_SUFFIX)
    If StrComp(chartName, dataName, vbTextCompare) = 0 Then
        chartName = Left$(dataName, 28) & "_Гр"
    End If
    StationChartSheetName = chartName
End Function

Public Sub WriteSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg, ByRef res As TGenResult)
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
    If g.ForcedNoRM Then note = Trim$(note & "; без РМ: принято S = " & Format(g.SPct, "0.0") & " %, fнч = " & Format(g.Fnch, "0.000") & " Гц (режим «Авария»)")
    If Len(note) > 0 And Left$(note, 2) = "; " Then note = Mid$(note, 3)
    If Not res.QuantApplicable Then quantStatusStr = "Н/П"

    Dim qualStatusStr As String, qualT5Str As String, qualT10Str As String, qualStStr As String
    If res.QualApplicable Then
        qualStatusStr = IIf(res.QualPass, "ОК", "Нарушение")
        qualT5Str = IIf(res.QualT5Pass, "ОК", "Нарушение")
        qualT10Str = IIf(res.QualT10Pass, "ОК", "Нарушение")
        qualStStr = IIf(res.QualSteadyPass, "ОК", "Нарушение")
    Else
        qualStatusStr = "Н/П": qualT5Str = "Н/П": qualT10Str = "Н/П": qualStStr = "Н/П"
    End If

    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        g.Station, g.Generator, g.EquipType, _
        res.StartTime, res.FirstExceedTime, _
        res.P0, res.PTek, res.PsteadyAvg, _
        res.Df, res.Dfr, _
        res.PReq, res.PFact, _
        Round(res.AmplPctPnom, 2), res.AmplitudeTag, _
        res.QuantPct, quantStatusStr, overshootStr, _
        res.TransientType, res.NumExtrema, _
        qualStatusStr, qualT5Str, qualT10Str, qualStStr, _
        res.QualFailedList, _
        FormatSecOrNA(res.T5FactSec), FormatSecOrNA(res.T10FactSec), _
        GeneratorSheetName(g, res.EventIdx), GeneratorChartSheetName(g, res.EventIdx), _
        res.PMaxEff, res.PMinEff, res.ReservePlus, res.ReserveMinus, _
        res.PReqOrig, limitStr, _
        note, _
        res.React15Status, Round(res.DP15, 3), Round(res.DeficitMaxPct, 2), Round(res.DeficitSec, 0), _
        res.Verdict, _
        res.EventIdx, res.Applicability, res.TrTime, Round(res.TrSec, 0), _
        IIf(res.TrQuasiSteady, "Квазиуст.", "Переходный"), _
        Round(res.PReqRef, 3), Round(res.PFactRef, 3), _
        IIf(res.OscViolation, "Нарушение: " & res.OscNote, IIf(Len(res.OscNote) > 0, res.OscNote, "нет")), _
        IIf(Len(res.TelemetryNote) > 0, res.TelemetryNote, "ОК"), _
        IIf(Len(res.ReactionType) > 0, res.ReactionType, "-"), _
        IIf(res.H15 > -900, Round(res.H15, 3), ""), _
        IIf(Len(res.QualClauseVerdict) > 0, res.QualClauseVerdict, IIf(res.QualApplicable, "", "Н/П (п. 4.1.3)")) _
    )
End Sub

Public Sub WriteSummaryInvalid(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg)
    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        g.Station, g.Generator, g.EquipType, "", "", "", "", "", "", "", "", "", "", "", "", _
        "Нарушение", "", "", 0, "Н/Д", "Н/Д", "Н/Д", "Н/Д", "", "", "", "Config", "", _
        "", "", "", "", "", "", _
        "Не заполнен обязательный параметр config", _
        "Н/Д", "", "", "", "НАРУШЕНИЕ: не заполнены обязательные параметры Config", _
        "", "", "", "", "", "", "", "", "", "", "", "" _
    )
End Sub


Public Sub WriteSummaryStationRow(ByVal ws As Worksheet, ByVal r As Long, ByRef sr As TStationResult)
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

    If Not sr.QuantApplicable Then quantStatusStr = "Н/П"

    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        sr.Station, genLabel, "Сумма станции (" & sr.Count & " ген., " & IIf(sr.IsNeblok, "неблок", "блок") & ")", _
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
        IIf(sr.React15Status = "Нарушение", sr.React15Note, IIf(sr.IsNeblok, "Оценка по суммарной мощности неблочной части (п. 4.2.10 МУ)", "Сумма по блочной станции - справочно")), _
        sr.React15Status, Round(sr.DP15, 3), Round(sr.DeficitMaxPct, 2), Round(sr.DeficitSec, 0), _
        sr.Verdict, _
        sr.EventIdx, sr.Applicability, sr.TrTime, Round(sr.TrSec, 0), IIf(sr.TrQuasiSteady, "Квазиуст.", "Переходный"), _
        "", "", "", "", "", "", "" _
    )
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 52)).Font.Bold = True
End Sub

Public Sub ApplySummaryConditionalFormat(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ws.Range("A2:AZ" & lastRow).Interior.Pattern = xlNone

    Dim rr As Long
    For rr = 2 To lastRow
        If Left$(CStr(ws.Cells(rr, 2).Value), 5) = "СУММА" Then
            ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 3)).Interior.Color = RGB(221, 235, 247)
        End If
    Next rr

    ApplyStatusCF ws, ws.Range("P2:P" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("T2:T" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("U2:U" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("V2:V" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("AJ2:AJ" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("AN2:AN" & lastRow), "ОК", "НАРУШЕНИЕ"
    ApplyStatusCF ws, ws.Range("AV2:AV" & lastRow), "нет", "Нарушение"
    ApplyStatusCF ws, ws.Range("AZ2:AZ" & lastRow), "выполнены", "НАРУШЕНИЕ"
    Set rng = ws.Range("AW2:AW" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="сомнительны", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    Set rng = ws.Range("AX2:AX" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="адекватная", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:=" - «", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    Set rng = ws.Range("AP2:AP" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="не удовлетворяет", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Только", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
    End With
    Set rng = ws.Range("AL2:AL" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    
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

    Set rng = ws.Range("P2:P" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Ограничен", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

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

    Set rng = ws.Range("Q2:Q" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Да", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

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

    Set rng = ws.Range("N2:N" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Слабое", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 225, 242)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Избыточное", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 217, 102)
    End With

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

    Set rng = ws.Range("AN2:AN" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="НЕ ОЦЕНИВАЕТСЯ", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With
    Set rng = ws.Range("P2:P" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Н/П", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With

    ws.Range("A1:AZ1").Font.Bold = True
    ws.Range("A1:AZ1").Interior.Color = RGB(217, 225, 242)
End Sub

Public Sub ApplyStatusCF(ByVal ws As Worksheet, ByVal rng As Range, ByVal okText As String, ByVal badText As String)
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

Public Sub WriteVersionStamp(ByVal wsSummary As Worksheet, ByVal wsRaw As Worksheet, ByVal t0Run As Double)
    Dim col As Long
    col = 54   ' колонка BB - правее данных Summary (после AZ)
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


Public Sub InitLog()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Уровень", "Источник", "Сообщение", "Время")
    ws.Range("A1:D1").Font.Bold = True
    m_LogRow = 2
End Sub

Public Sub AppendLog(ByVal level As String, ByVal source As String, ByVal message As String)
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

Public Sub FinalizeLog()
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

Public Sub ValidateInputs(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByRef st As TSettings, ByVal timeCol As Long)
    Dim lastR As Long, r As Long, prev As Double, cur As Double, gap As Double, maxGap As Double
    Dim cfgLast As Long, g As TGenCfg, headerRow As Long
    Dim paropipes As Object, stationParopipe As Object, st2 As String, kk As Variant

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

    cfgLast = LastUsedRow(wsCfg)
    Set stationParopipe = CreateObject("Scripting.Dictionary")
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfgEff(wsCfg, r, st)
        If Not g.Enabled Then GoTo NX
        If g.ForcedNoRM Then AppendLog "INFO", g.Station & "/" & g.Generator, _
            "Без РМ, режим «Авария»: принято S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц (в Config S = " & g.SPctCfg & " %, fнч = " & g.FnchCfg & " Гц)"
        If Not g.HasRM And Not st.IsEmergency Then AppendLog "INFO", g.Station & "/" & g.Generator, _
            "Без РМ, режим «Испытание»: используются фактические S = " & g.SPct & " %, fнч = " & g.Fnch & " Гц из Config"
        If g.PNom <= 0 Then AppendLog "WARN", g.Station & "/" & g.Generator, "Pном <= 0"
        If g.SPct <= 1 Or g.SPct > 15 Then AppendLog "WARN", g.Station & "/" & g.Generator, "S вне [1..15] %: " & g.SPct
        If g.Fnch < 0 Or g.Fnch > 0.5 Then AppendLog "WARN", g.Station & "/" & g.Generator, "fнч вне [0..0.5] Гц: " & g.Fnch
        If g.T10Sec < g.T5Sec Then AppendLog "WARN", g.Station & "/" & g.Generator, "t10 < t5 (" & g.T10Sec & " < " & g.T5Sec & ")"
        If g.PMax > 0 And g.PMax < g.PMin Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax < Pmin (" & g.PMax & " < " & g.PMin & ")"
        If g.PMax > 0 And g.PNom > 0 And g.PMax > 1.3 * g.PNom Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax > 1.3*Pном (проверьте ед. измерения)"
        If g.PMin < 0 Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmin < 0: " & g.PMin
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

    If st.QuantIntervalSec < 60 Then AppendLog "WARN", "Settings", "Запасной интервал tр < 60 с: " & st.QuantIntervalSec
    If st.SteadyWindowSec < 10 Then AppendLog "WARN", "Settings", "Окно установившегося < 10 с: " & st.SteadyWindowSec
    If st.React15ThrPct > 2 Then AppendLog "WARN", "Settings", "Порог реакции 15 с > 2 %Pном (" & st.React15ThrPct & ") - п. 4.2.11 требует лишь наличия реакции"
    If st.AvgWindowSec > 30 Then AppendLog "WARN", "Settings", "Окно усреднения P0/Pтек > 30 с: " & st.AvgWindowSec
    If st.MaxReqPct > 0 And st.MaxReqPct < 10 Then AppendLog "WARN", "Settings", "Ограничение Pтреб < 10 %Pном (" & st.MaxReqPct & ") - по Приложению к МУ принимается 10 %"
    If st.MaxEventSec < st.QuantIntervalSec Then AppendLog "WARN", "Settings", "Макс. окно события (" & st.MaxEventSec & " с) меньше запасного интервала tр (" & st.QuantIntervalSec & " с)"
    AppendLog "INFO", "Settings", "Строгий дефицит п. 4.2.14.1: " & IIf(st.StrictDeficit, "вкл (дефицит в любом периоде -> НАРУШЕНИЕ)", "выкл (дефицит -> замечание)")

    Dim tFirst As Variant, tLast As Variant, spanMin As Double
    If IsDate(wsRaw.Cells(2, timeCol).Value) And IsDate(wsRaw.Cells(lastR, timeCol).Value) Then
        tFirst = wsRaw.Cells(2, timeCol).Value
        tLast = wsRaw.Cells(lastR, timeCol).Value
        spanMin = SecBetween(tFirst, tLast) / 60#
        If spanMin < 10 Then
            AppendLog "WARN", "RawData", "Длительность выборки ~ " & Format(spanMin, "0.0") & _
                      " мин. По п. 3.1 МУ желательно >= 5 мин до отклонения и до возврата в 50±0.1 Гц более 5 мин."
        Else
            AppendLog "INFO", "RawData", "Длительность выборки ~ " & Format(spanMin, "0.0") & " мин (" & _
                      Format(tFirst, "dd.mm.yyyy hh:mm:ss") & " — " & Format(tLast, "dd.mm.yyyy hh:mm:ss") & ")"
        End If
    End If

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


Public Sub CollectOldOutputSheets(ByRef names As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_RAW And ws.Name <> SH_CFG And ws.Name <> SH_SUM And ws.Name <> SH_LOG _
           And ws.Name <> SH_INSTR And ws.Name <> SH_QUERY Then
            names.Add ws.Name
        End If
    Next ws
End Sub


Public Sub DeleteOutputSheets(ByVal names As Collection)
    Dim i As Long
    Application.DisplayAlerts = False
    For i = names.Count To 1 Step -1
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(names(i))).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
End Sub

Public Function GeneratorSheetName(ByRef g As TGenCfg, ByVal evIdx As Long) As String
    Dim base As String, sfx As String
    sfx = EventSuffix(evIdx)
    base = g.Station & "_" & g.Generator
    If Len(base & sfx) > 31 Then base = Left$(base, 31 - Len(sfx))
    GeneratorSheetName = MakeSheetName(base & sfx)
End Function

Public Function GeneratorChartSheetName(ByRef g As TGenCfg, ByVal evIdx As Long) As String
    Dim base As String, full As String, sfx As String
    sfx = EventSuffix(evIdx)
    base = g.Station & "_" & g.Generator
    full = base & sfx & CHART_SUFFIX
    If Len(full) > 31 Then
        base = Left$(base, 31 - Len(CHART_SUFFIX) - Len(sfx))
        full = base & sfx & CHART_SUFFIX
    End If
    GeneratorChartSheetName = MakeSheetName(full)
End Function
