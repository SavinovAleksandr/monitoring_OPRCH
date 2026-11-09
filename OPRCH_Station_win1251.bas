Attribute VB_Name = "OPRCH_Station"
Option Explicit


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


Public Function BuildStationVerdict(ByRef sr As TStationResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String, objName As String
    objName = "по сумме " & IIf(Len(sr.Paroprovod) > 0, "паропровода " & sr.Paroprovod & " ", "станции ")
    bad = ""
    remark = ""

    If Not sr.QuantApplicable And sr.React15Status <> "Нарушение" Then
        txt = "НЕ ОЦЕНИВАЕТСЯ " & objName & ": " & sr.Applicability
        If Not sr.IsNeblok Then txt = txt & ". Блочная компоновка: п. 4.2.10 МУ не применяется, оценка - по каждому генератору"
        BuildStationVerdict = txt
        Exit Function
    End If

    If sr.QuantApplicable Then
        If sr.Limited Then
            remark = remark & "Pтреб_сум ограничен " & sr.LimitType & " (п. 4.2.8: оценка по располагаемому резерву); "
        ElseIf Not sr.QuantPass Then
            bad = bad & "Pфакт_сум(tр) = " & Format(sr.PFact, "0.00") & " МВт ниже Pтреб_сум(tр) = " & Format(sr.PReq, "0.00") & _
                  " МВт (" & Format(sr.QuantPct, "0") & " %) - п. 4.2.14.1; "
        End If
        If sr.DeficitSec > 0 And Not sr.Limited Then
            If st.StrictDeficit Or st.IsEmergency Then
                bad = bad & "дефицит до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум в течение " & _
                      Format(sr.DeficitSec, "0") & " с хотя бы в одном периоде - п. 4.2.14.1; "
            Else
                remark = remark & "дефицит до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум в течение " & _
                         Format(sr.DeficitSec, "0") & " с (п. 4.2.14.1); "
            End If
        End If
    Else
        remark = remark & sr.Applicability & "; "
    End If
    If sr.React15Status = "Нарушение" Then
        bad = bad & "отсутствует реакция РЧВ турбин через 15 с (dP15_сум = " & Format(sr.DP15, "0.00") & " МВт) - п. 4.2.11 / 4.4.2; "
    End If

    If sr.IsNeblok Then
        If Len(bad) > 0 Then
            txt = "НАРУШЕНИЕ " & objName & "(неблочная часть, " & sr.Count & " ген.): " & bad & _
                  "По п. 4.2.10 МУ требуется индивидуальная оценка каждого генератора - см. строки генераторов станции. "
        Else
            txt = "ОК " & objName & "(неблочная часть, " & sr.Count & " ген.): количественно (п. 4.2.14.1) и по реакции 15 с (п. 4.2.11) соответствует. " & _
                  "По п. 4.2.10 МУ индивидуальная оценка генераторов не требуется (строки генераторов приведены справочно). "
        End If
    Else
        If Len(bad) > 0 Then
            txt = "НАРУШЕНИЕ " & objName & "(" & sr.Count & " ген., блочная компоновка - справочно): " & bad & _
                  "Нормативная оценка - по каждому генератору (п. 4.2.10 применяется к неблочной части ТЭС). "
        Else
            txt = "ОК " & objName & "(" & sr.Count & " ген., блочная компоновка - справочно): количественно и по реакции 15 с соответствует. " & _
                  "Нормативная оценка - по каждому генератору (п. 4.2.10 применяется к неблочной части ТЭС). "
        End If
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    txt = Trim$(txt)
    If Right$(txt, 1) = ";" Then txt = Left$(txt, Len(txt) - 1)
    BuildStationVerdict = txt
End Function

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
