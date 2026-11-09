Attribute VB_Name = "OPRCH_Engine"
Option Explicit


Public Function AnalyzeOneGenerator(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                     ByRef ev As TEvent, ByVal evIdx As Long) As TGenResult
    On Error GoTo EH

    Dim res As TGenResult
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim startRow As Long, endQ As Long, endQual As Long, firstExceedRow As Long, evEnd As Long
    Dim p0 As Double, ptek As Double, df As Double, dfr As Double, preq As Double, pfact As Double
    Dim kdQuant As Double, tQuantSec As Double
    Dim qpct As Double, qpass As Boolean
    Dim calcStep As String
    Dim f0 As Double

    calcStep = "Поиск колонок"
    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
    fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
    If pCol = 0 Then Err.Raise vbObjectError + 2101, , "Не найдена колонка мощности '" & g.PowerHeader & "' для " & g.Generator
    If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator

    calcStep = "Границы события"
    startRow = ev.StartRow
    firstExceedRow = ev.FirstExceedRow
    evEnd = ev.EndRow
    If evEnd < startRow Then evEnd = startRow
    res.EventIdx = evIdx
    res.EventEndRow = evEnd

    calcStep = "Опорная частота f0 (п. 4.2.4.1)"
    f0 = ResolveF0(wsRaw, fCol, startRow, g.HasRM, st.FNom)
    res.F0 = f0

    calcStep = "Определение конца контрольного интервала"
    endQ = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)
    calcStep = "Определение конца качественного интервала"
    endQual = RowByTimeOffset(wsRaw, timeCol, startRow, g.T10Sec)

    calcStep = "Чтение P0"
    p0 = AvgPowerBackward(wsRaw, pCol, timeCol, startRow, st.AvgWindowSec)

    calcStep = "Применимость оценки (п. 4.1.2/4.1.3)"
    EvaluateApplicability wsRaw, st, g, fCol, timeCol, startRow, firstExceedRow, evEnd, f0, res

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

    calcStep = "Контрольный интервал (Config W5, запасной): Pтреб/Pфакт"
    Dim dfRef As Double, dfrRef As Double, preqRefOrig As Double, preqRef As Double
    tQuantSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(endQ, timeCol).Value)
    dfRef = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, endQ, f0)
    dfrRef = DeadbandDeviation(dfRef, g.Fnch)
    preqRefOrig = PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, tQuantSec), dfrRef, st.MaxReqPct)
    preqRef = ClampReserve(preqRefOrig, reservePlus, reserveMinus)
    res.PReqRef = preqRef
    res.PFactRef = AvgPowerBackward(wsRaw, pCol, timeCol, endQ, st.AvgWindowSec) - p0

    calcStep = "Выбор tр (п. 4.2.5/4.2.9)"
    SelectEvaluationPoint wsRaw, st, g, pCol, fCol, timeCol, startRow, firstExceedRow, evEnd, endQ, p0, f0, reservePlus, reserveMinus, res
    endQ = res.TrRow
    tQuantSec = res.TrSec

    calcStep = "Чтение Pтек(tр)"
    ptek = AvgPowerBackward(wsRaw, pCol, timeCol, endQ, st.AvgWindowSec)
    calcStep = "Расчет dF/dFr в точке tр"
    df = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, evEnd, f0)
    If res.TrQuasiSteady Then
        dfr = DeadbandDeviation(MeanFreqAround(wsRaw, fCol, timeCol, endQ, QS_HALF_WINDOW_SEC, f0) - f0, g.Fnch)
    Else
        dfr = DeadbandDeviation(NzD(wsRaw.Cells(endQ, fCol).Value, f0) - f0, g.Fnch)
    End If
    If dfr = 0 And dfrRef <> 0 And res.TrQuasiSteady Then dfr = dfrRef

    calcStep = "Расчет требуемой мощности"
    Dim preqOrig As Double
    kdQuant = DynamicKdByTime(g.EquipType, tQuantSec)
    preqOrig = PReqCalc(g.SPct, g.PNom, st.FNom, kdQuant, dfr, st.MaxReqPct)

    calcStep = "Ограничение Pтреб резервом"
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

    calcStep = "Статический Pтреб"
    Dim preqStatic As Double
    preqStatic = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, 1#, dfr, st.MaxReqPct), reservePlus, reserveMinus)
    res.PReqStatic = preqStatic
    res.PMaxEff = pMaxEff
    res.PMinEff = pMinEff
    res.ReservePlus = reservePlus
    res.ReserveMinus = reserveMinus
    res.QuantPct = qpct
    res.QuantPass = qpass
    res.Overshoot = (qpass And qpct > (100# + st.QuantTolPct))

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
    If g.PNom > 0 Then
        res.AmplPctPnom = 100# * Abs(preqOrig) / g.PNom
    Else
        res.AmplPctPnom = 0
    End If
    res.AmplitudeTag = AmplitudeTag(res.AmplPctPnom, dfr <> 0)

    calcStep = "Реакция РЧВ через 15 с (п. 4.2.11)"
    Dim r15Status As String, r15Pass As Boolean, dp15 As Double, r15Note As String
    EvaluateReaction15 wsRaw, st, pCol, fCol, timeCol, startRow, firstExceedRow, p0, f0, g.PNom, g.SPct, _
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
    EvaluateDeficit wsRaw, st, pCol, fCol, timeCol, startRow, evEnd, p0, f0, g.PNom, g.SPct, g.Fnch, _
                    g.EquipType, reservePlus, reserveMinus, defMaxPct, defSec
    res.DeficitMaxPct = defMaxPct
    res.DeficitSec = defSec

    calcStep = "Характер переходного процесса"
    EvaluateTransient wsRaw, st, g, res, pCol, timeCol

    calcStep = "Колебания мощности (п. 4.2.6, 4.2.14.4)"
    EvaluateOscillations wsRaw, st, g, res, pCol, fCol, timeCol, startRow, evEnd

    calcStep = "Качество телеизмерений"
    EvaluateTelemetry wsRaw, g, res, pCol, timeCol, startRow, evEnd

    calcStep = "Качественный критерий (Прил. 1: ступени; в «Аварии» — справочно)"
    EvaluateQualitative wsRaw, st, g, res, pCol, fCol, timeCol, f0

    calcStep = "Идентификация переходной функции h(15 с) (Прил. п. 6.2)"
    res.H15 = IdentifyH15(wsRaw, st, g, pCol, fCol, timeCol, firstExceedRow, p0, f0, res.H15Note)

    calcStep = "Тип реакции (п. 4.4.1)"
    ClassifyReactionType wsRaw, st, g, res, pCol, fCol, timeCol, startRow, evEnd

    calcStep = "Итоговый вывод"
    res.Verdict = BuildGeneratorVerdict(g, res, st)
    If Left$(res.Verdict, 9) = "НАРУШЕНИЕ" Then AppendLog "WARN", g.Station & "/" & g.Generator, res.Verdict

    AnalyzeOneGenerator = res
    Exit Function

EH:
    Err.Raise vbObjectError + 2199, , "AnalyzeOneGenerator (" & g.Generator & " / " & calcStep & "): " & Err.Description
End Function


Public Function AvgPowerBackward(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
                                  ByVal row As Long, ByVal windowSec As Double) As Double
    Dim r1 As Long, r2 As Long, r As Long, sumP As Double, cnt As Long
    If windowSec <= 0 Then
        AvgPowerBackward = NzD(wsRaw.Cells(row, pCol).Value, 0)
        Exit Function
    End If
    WindowRowsAround wsRaw, timeCol, row, windowSec, r1, r2
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


Public Sub EvaluateReaction15(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                               ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                               ByVal startRow As Long, ByVal firstExceedRow As Long, _
                               ByVal p0 As Double, ByVal f0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
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
    dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, fCol).Value, f0) - f0, fnch)
    preq15 = 0
    If dfr15 <> 0 And sPct > 0 Then
        preq15 = PReqCalc(sPct, pNom, st.FNom, DynamicKdByTime(equipType, 15#), dfr15, st.MaxReqPct)
        preq15 = ClampReserve(preq15, reservePlus, reserveMinus)
    End If
    thrMW = React15ThresholdMW(st, pNom, preq15)
    ClassifyReaction15 st, dp15, pNom, needSign, reservePlus, reserveMinus, thrMW, (dfr15 = 0), status, pass, note
End Sub


Public Function React15ThresholdMW(ByRef st As TSettings, ByVal pNom As Double, ByVal preq15 As Double) As Double
    Dim thr As Double
    thr = st.React15ThrPct / 100# * pNom
    If thr < 0 Then thr = 0
    If Abs(preq15) > 0 And 0.5 * Abs(preq15) < thr Then thr = 0.5 * Abs(preq15)
    React15ThresholdMW = thr
End Function


Public Sub ClassifyReaction15(ByRef st As TSettings, ByVal dp15 As Double, ByVal pNom As Double, _
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


Public Sub EvaluateDeficit(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                            ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                            ByVal startRow As Long, ByVal endRow As Long, _
                            ByVal p0 As Double, ByVal f0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
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
        dFrT = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, f0) - f0, fnch)
        If dFrT = 0 Then GoTo NextR
        tSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        If tSec < 15 Then GoTo NextR
        preqT = PReqCalc(sPct, pNom, st.FNom, DynamicKdByTime(equipType, tSec), dFrT, st.MaxReqPct)
        preqT = ClampReserve(preqT, reservePlus, reserveMinus)
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


Public Function BuildGeneratorVerdict(ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String, okPart As String
    bad = ""
    remark = ""

    If Not res.QuantApplicable And Not res.QualApplicable Then
        txt = "НЕ ОЦЕНИВАЕТСЯ: " & res.Applicability
        If Len(res.TelemetryNote) > 0 And Not res.TelemetryOK Then txt = txt & "; " & res.TelemetryNote
        BuildGeneratorVerdict = txt
        Exit Function
    End If

    If res.QuantApplicable Then
        If res.Limited Then
            remark = remark & "Pтреб ограничен " & res.LimitType & " (п. 4.2.8: оценка по располагаемому резерву); "
        ElseIf Not res.QuantPass Then
            bad = bad & "Pфакт(tр) = " & Format(res.PFact, "0.00") & " МВт ниже Pтреб(tр) = " & Format(res.PReq, "0.00") & _
                  " МВт (" & Format(res.QuantPct, "0") & " %) - п. 4.2.14.1 (с учётом допуска п. 4.2.7/4.2.8 и Kд(t) Приложения); "
        End If
        If res.DeficitSec > 0 And Not res.Limited Then
            If st.StrictDeficit Or st.IsEmergency Then
                bad = bad & "дефицит первичной мощности до " & Format(res.DeficitMaxPct, "0.0") & " %Pном в течение " & _
                      Format(res.DeficitSec, "0") & " с хотя бы в одном периоде - п. 4.2.14.1; "
            Else
                remark = remark & "дефицит до " & Format(res.DeficitMaxPct, "0.0") & " %Pном в течение " & _
                         Format(res.DeficitSec, "0") & " с (п. 4.2.14.1); "
            End If
        End If
        If res.OscViolation Then
            bad = bad & res.OscNote & " - п. 4.2.14.4; "
        ElseIf Len(res.OscNote) > 0 And res.OscPeriods > 0 Then
            remark = remark & res.OscNote & "; "
        End If
    Else
        remark = remark & "количественная оценка не применяется (п. 4.1.2: " & res.Applicability & "); "
    End If

    If res.React15Status = "Нарушение" Then
        bad = bad & "отсутствует реакция РЧВ турбины через 15 с (dP15 = " & Format(res.DP15, "0.00") & " МВт) - п. 4.2.11; "
    ElseIf res.React15Status = "Н/П" And Len(res.React15Note) > 0 And res.Dfr <> 0 Then
        remark = remark & "реакция 15 с не оценивается (" & res.React15Note & "); "
    End If

    If res.QualApplicable Then
        If Len(res.ReactionType) > 0 Then remark = remark & "тип реакции по п. 4.4.1: " & res.ReactionType & "; "
        If Left$(res.QualClauseVerdict, 9) = "НАРУШЕНИЕ" Then
            bad = bad & Mid$(res.QualClauseVerdict, 12) & "; "
        ElseIf Len(res.QualClauseVerdict) > 0 Then
            remark = remark & res.QualClauseVerdict & "; "
        End If
        If g.QualEnabled And res.Dfr <> 0 And Not res.QualPass Then
            If Not st.IsEmergency Then
                bad = bad & "ступени динамики (Прил. 1, табл. Е): не выполнены " & res.QualFailedList & "; "
            Else
                remark = remark & "ступени динамики (Прил. 1, табл. Е, справочно в режиме «Авария»): не выполнены " & res.QualFailedList & "; "
            End If
        End If
    End If

    If Len(res.TelemetryNote) > 0 And Not res.TelemetryOK Then remark = remark & res.TelemetryNote & "; "
    If res.H15 > -900 Then
        remark = remark & "h(15 с) = " & Format(res.H15, "0.00") & IIf(res.H15 < H15_REF, " < " & Format(H15_REF, "0.0") & " (ориентир)", "") & _
                 " (идентификация переходной функции, Прил. п. 6.2, справочно); "
    End If
    If g.ForcedNoRM Then
        remark = remark & "оборудование без РМ, режим «Авария»: принято S = " & Format(g.SPct, "0.0") & " %, fнч = " & _
                 Format(g.Fnch, "0.000") & " Гц (в Config: S = " & Format(g.SPctCfg, "0.0") & " %, fнч = " & Format(g.FnchCfg, "0.000") & " Гц); "
    End If
    If Abs(res.F0 - st.FNom) > 0.0005 Then
        remark = remark & "f0 = " & Format(res.F0, "0.000") & " Гц (без РМ: частота в t0, п. 4.2.4.1; fном = " & Format(st.FNom, "0.00") & " Гц); "
    End If

    If Len(bad) > 0 Then
        txt = "НАРУШЕНИЕ: " & bad
    Else
        okPart = "ОК: "
        If res.QuantApplicable Then okPart = okPart & "количественно (п. 4.2.14.1) и "
        okPart = okPart & "по реакции 15 с (п. 4.2.11) соответствует"
        If res.QualApplicable And res.ReactionCode = 1 Then
            okPart = okPart & "; качественно - реакция «адекватная» (п. 4.4.1)"
        ElseIf res.QualApplicable And res.ReactionCode > 1 Then
            okPart = okPart & "; качественно - тип реакции " & res.ReactionType & " (п. 4.4.1)"
        End If
        If Not st.IsEmergency And g.QualEnabled And res.QualPass Then
            okPart = okPart & "; ступени динамики (Прил. 1) выполнены"
        End If
        txt = okPart & "; "
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    If Right$(txt, 2) = "; " Then txt = Left$(txt, Len(txt) - 2)
    BuildGeneratorVerdict = txt
End Function



Public Function ResolveF0(ByVal wsRaw As Worksheet, ByVal fCol As Long, ByVal startRow As Long, _
                           ByVal hasRM As Boolean, ByVal fNom As Double) As Double
    If hasRM Then
        ResolveF0 = fNom
    Else
        ResolveF0 = NzD(wsRaw.Cells(startRow, fCol).Value, fNom)
    End If
End Function


Public Function FindEvents(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal fCol As Long, _
                            ByRef st As TSettings, ByVal fnch As Double, ByVal fRef As Double, _
                            ByRef ev() As TEvent) As Long
    Dim lastR As Long, r As Long, n As Long
    Dim prevOut As Boolean, curOut As Boolean, inEvent As Boolean
    Dim lastOutRow As Long, tStart As Double, tCur As Double, tLastOut As Double
    Dim outsideSec As Double, bestRow As Long, bestAbs As Double, val As Double
    Dim dropped As Long
    Dim backInBand As Double, tBackStart As Double

    lastR = LastUsedRow(wsRaw)
    ReDim ev(1 To 1)
    n = 0
    If fRef <= 0 Then fRef = st.FNom

    If lastR < 3 Then
        ev(1).StartRow = 2: ev(1).FirstExceedRow = 2: ev(1).EndRow = lastR
        FindEvents = 1
        Exit Function
    End If

    If Not st.AutoStart And st.EventStart > 0 Then
        bestRow = 2: bestAbs = 1E+99
        For r = 2 To lastR
            If IsDate(wsRaw.Cells(r, timeCol).Value) Then
                val = Abs(CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) - st.EventStart)
                If val < bestAbs Then bestAbs = val: bestRow = r
            End If
        Next r
        ev(1).StartRow = bestRow
        ev(1).FirstExceedRow = bestRow
        ev(1).EndRow = RowByTimeOffset(wsRaw, timeCol, bestRow, st.MaxEventSec)
        FindEvents = 1
        Exit Function
    End If

    prevOut = (Abs(NzD(wsRaw.Cells(2, fCol).Value, fRef) - fRef) > fnch)
    inEvent = False
    For r = 3 To lastR
        curOut = (Abs(NzD(wsRaw.Cells(r, fCol).Value, fRef) - fRef) > fnch)
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then tCur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) * 86400#
        If Not inEvent Then
            If curOut And Not prevOut Then
                inEvent = True
                n = n + 1
                ReDim Preserve ev(1 To n)
                ev(n).StartRow = r - 1
                ev(n).FirstExceedRow = r
                ev(n).EndRow = 0
                lastOutRow = r
                tStart = tCur
                tLastOut = tCur
                outsideSec = 0
                tBackStart = 0
            End If
        Else
            If curOut Then
                lastOutRow = r
                tLastOut = tCur
                outsideSec = outsideSec + 1
                tBackStart = 0
            ElseIf tBackStart = 0 Then
                tBackStart = tCur
            End If
            backInBand = Abs(NzD(wsRaw.Cells(r, fCol).Value, fRef) - st.FNom)
            If (Not curOut And (tCur - tLastOut) >= st.SteadyWindowSec) _
               Or (backInBand <= 0.1 And tBackStart > 0 And (tCur - tBackStart) >= st.SteadyWindowSec) _
               Or (tCur - tStart) >= st.MaxEventSec Then
                ev(n).EndRow = r
                inEvent = False
                If SecBetween(wsRaw.Cells(ev(n).FirstExceedRow, timeCol).Value, wsRaw.Cells(lastOutRow, timeCol).Value) + 1 < EVENT_MIN_OUTSIDE_SEC Then
                    n = n - 1
                    dropped = dropped + 1
                    If n >= 1 Then ReDim Preserve ev(1 To n) Else ReDim ev(1 To 1)
                End If
            End If
        End If
        prevOut = curOut
    Next r
    If inEvent Then ev(n).EndRow = lastR

    If n = 0 Then
        bestRow = 2: bestAbs = -1
        For r = 2 To lastR
            val = NzD(wsRaw.Cells(r, fCol).Value, fRef) - fRef
            If Abs(val) > bestAbs Then bestAbs = Abs(val): bestRow = r
        Next r
        ReDim ev(1 To 1)
        ev(1).StartRow = bestRow
        ev(1).FirstExceedRow = bestRow
        ev(1).EndRow = RowByTimeOffset(wsRaw, timeCol, bestRow, st.MaxEventSec)
        n = 1
    End If

    If dropped > 0 Then AppendLog "INFO", "События", "Пропущено коротких выбросов частоты за fнч (< " & EVENT_MIN_OUTSIDE_SEC & " с): " & dropped
    If st.MaxEvents >= 1 And n > st.MaxEvents Then
        AppendLog "WARN", "События", "Найдено " & n & " событий; оцениваются первые " & st.MaxEvents & " (Config W15)"
        ReDim Preserve ev(1 To st.MaxEvents)
        n = st.MaxEvents
    End If
    FindEvents = n
End Function


Public Sub EvaluateApplicability(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                  ByVal fCol As Long, ByVal timeCol As Long, _
                                  ByVal startRow As Long, ByVal firstExceedRow As Long, ByVal evEnd As Long, _
                                  ByVal f0 As Double, ByRef res As TGenResult)
    Dim r As Long, d As Double, dMax As Double, rMax As Long
    Dim run02 As Double, best02 As Double, outSec As Double, dt As Double
    Dim prevT As Double, curT As Double, minQuantSec As Double, isGpa As Boolean
    Dim txtQ As String, txtK As String

    dMax = 0: rMax = startRow
    run02 = 0: best02 = 0: outSec = 0
    prevT = -1
    For r = startRow To evEnd
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            curT = CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) * 86400#
            If prevT >= 0 Then dt = curT - prevT Else dt = 1
            If dt <= 0 Or dt > 60 Then dt = 1
            prevT = curT
        Else
            dt = 1
        End If
        d = NzD(wsRaw.Cells(r, fCol).Value, f0) - f0
        If Abs(d) > Abs(dMax) Then dMax = d: rMax = r
        If Abs(d) >= QUANT_DF_HZ - 0.0000001 Then
            run02 = run02 + dt
            If run02 > best02 Then best02 = run02
        Else
            run02 = 0
        End If
        If Abs(d) > g.Fnch Then outSec = outSec + dt
    Next r

    res.DfMaxEvent = dMax
    res.DurAbove02Sec = best02
    res.DurOutsideSec = outSec
    If firstExceedRow > 0 And rMax >= firstExceedRow Then
        res.JumpSec = SecBetween(wsRaw.Cells(firstExceedRow, timeCol).Value, wsRaw.Cells(rMax, timeCol).Value)
    Else
        res.JumpSec = 0
    End If

    If Not st.IsEmergency Then
        res.QuantApplicable = True
        res.QualApplicable = True
        res.Applicability = "Испытание: колич. (Прил. 14) + динамика t5/t10 (Прил. 1); применимость 4.1.2/4.1.3 не отсекает"
        Exit Sub
    End If

    isGpa = (InStr(UCase$(g.EquipType), "ГПА") > 0 Or InStr(UCase$(g.EquipType), "ДГУ") > 0)
    minQuantSec = IIf(isGpa, QUANT_MIN_SEC_GPA, QUANT_MIN_SEC_TES)

    res.QuantApplicable = (best02 >= minQuantSec - 0.5)
    If res.QuantApplicable Then
        txtQ = "колич. (п. 4.1.2: |dF| >= 0.2 Гц в течение " & Format(best02, "0") & " с)"
    ElseIf best02 > 0 Then
        txtQ = "|dF| >= 0.2 Гц лишь " & Format(best02, "0") & " с (< " & Format(minQuantSec, "0") & " с по п. 4.1.2)"
    Else
        txtQ = "|dF|max = " & Format(Abs(dMax), "0.000") & " Гц < 0.2 Гц (п. 4.1.2)"
    End If

    Dim jumpOk As Boolean, durOk As Boolean, levelOk As Boolean
    jumpOk = (res.JumpSec <= QUAL_JUMP_SEC + 0.5)
    durOk = (outSec >= QUAL_MIN_SEC - 0.5)
    If g.HasRM Then
        levelOk = (Abs(dMax) > QUAL_RM_DF_HZ)
    Else
        levelOk = (Abs(dMax) >= QUANT_DF_HZ - 0.0000001 And Abs(dMax) <= QUAL_NORM_DF_MAX_HZ + 0.0000001)
    End If
    res.QualApplicable = jumpOk And durOk And levelOk
    If res.QualApplicable Then
        txtK = "кач. (п. 4.1.3, " & IIf(g.HasRM, "п. 4.4.4: с РМ, |dF| > 0.15 Гц", "п. 4.4.3: без РМ, |dF| в 0.2..0.4 Гц") & _
               "; нарастание " & Format(res.JumpSec, "0") & " с, вне fнч " & Format(outSec, "0") & " с)"
    Else
        txtK = "кач. не применяется ("
        If Not jumpOk Then txtK = txtK & "нарастание " & Format(res.JumpSec, "0") & " с > 10 с; "
        If Not durOk Then txtK = txtK & "вне fнч лишь " & Format(outSec, "0") & " с < 20 с; "
        If Not levelOk Then txtK = txtK & "|dF|max = " & Format(Abs(dMax), "0.000") & " Гц вне диапазона " & _
            IIf(g.HasRM, "> 0.15 Гц (п. 4.4.4)", "0.2..0.4 Гц (п. 4.4.3)") & "; "
        If Right$(txtK, 2) = "; " Then txtK = Left$(txtK, Len(txtK) - 2)
        txtK = txtK & ")"
    End If

    If res.QuantApplicable And res.QualApplicable Then
        res.Applicability = "Колич. + кач.: " & txtQ & "; " & txtK
    ElseIf res.QuantApplicable Then
        res.Applicability = "Только колич.: " & txtQ & "; " & txtK
    ElseIf res.QualApplicable Then
        res.Applicability = "Только кач.: " & txtQ & "; " & txtK
    Else
        res.Applicability = "отклонение частоты не удовлетворяет п. 4.1.2/4.1.3 МУ: " & txtQ & "; " & txtK
    End If
End Sub


Public Sub SelectEvaluationPoint(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                  ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                  ByVal startRow As Long, ByVal firstExceedRow As Long, ByVal evEnd As Long, _
                                  ByVal endRef As Long, ByVal p0 As Double, ByVal f0 As Double, _
                                  ByVal reservePlus As Double, ByVal reserveMinus As Double, ByRef res As TGenResult)
    Dim r As Long, fMean As Double, dfrT As Double, tSec As Double, preqT As Double
    Dim have As Double, deficit As Double, tol As Double
    Dim bestRow As Long, bestDef As Double, bestDfrRow As Long, bestDfr As Double
    Dim anyQS As Boolean, stepR As Long, lastR As Long
    Dim fbRow As Long, fbDef As Double, fbDfrRow As Long, fbDfr As Double
    Dim fHere As Double

    tol = 0.01 * g.PNom
    bestRow = 0: bestDef = -1E+99
    bestDfrRow = 0: bestDfr = 0
    anyQS = False
    fbRow = 0: fbDef = -1E+99
    fbDfrRow = 0: fbDfr = 0
    lastR = LastUsedRow(wsRaw)
    If evEnd > lastR Then evEnd = lastR

    stepR = 1
    If evEnd - startRow > 600 Then stepR = (evEnd - startRow) \ 600 + 1

    For r = startRow To evEnd Step stepR
        tSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        If tSec < 15 Then GoTo NextR
        fHere = NzD(wsRaw.Cells(r, fCol).Value, f0)
        dfrT = DeadbandDeviation(fHere - f0, g.Fnch)
        If dfrT <> 0 Then
            preqT = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, tSec), dfrT, st.MaxReqPct), reservePlus, reserveMinus)
            have = SgnNZ(preqT) * (AvgPowerBackward(wsRaw, pCol, timeCol, r, st.AvgWindowSec) - p0)
            deficit = (Abs(preqT) - tol) - have
            If deficit > fbDef Then fbDef = deficit: fbRow = r
            If Abs(dfrT) > Abs(fbDfr) Then fbDfr = dfrT: fbDfrRow = r
        End If
        If Not IsQuasiSteady(wsRaw, fCol, timeCol, r, f0, fMean) Then GoTo NextR
        dfrT = DeadbandDeviation(fMean - f0, g.Fnch)
        If dfrT = 0 Then GoTo NextR
        anyQS = True
        preqT = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, tSec), dfrT, st.MaxReqPct), reservePlus, reserveMinus)
        have = SgnNZ(preqT) * (AvgPowerBackward(wsRaw, pCol, timeCol, r, st.AvgWindowSec) - p0)
        deficit = (Abs(preqT) - tol) - have
        If deficit > bestDef Then bestDef = deficit: bestRow = r
        If Abs(dfrT) > Abs(bestDfr) Then bestDfr = dfrT: bestDfrRow = r
NextR:
    Next r

    If anyQS Then
        If bestDef > 0 Then
            res.TrRow = bestRow
        Else
            res.TrRow = IIf(bestDfrRow > 0, bestDfrRow, bestRow)
        End If
        res.TrQuasiSteady = True
    ElseIf fbRow > 0 Then
        If fbDef > 0 Then
            res.TrRow = fbRow
        Else
            res.TrRow = IIf(fbDfrRow > 0, fbDfrRow, fbRow)
        End If
        res.TrQuasiSteady = False
    Else
        res.TrRow = endRef
        res.TrQuasiSteady = False
    End If
    If res.TrRow < startRow Then res.TrRow = startRow
    res.TrTime = wsRaw.Cells(res.TrRow, timeCol).Value
    res.TrSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(res.TrRow, timeCol).Value)
End Sub


Public Sub EvaluateOscillations(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                 ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                 ByVal startRow As Long, ByVal evEnd As Long)
    Dim n As Long, i As Long, r As Long
    Dim p() As Double, f() As Double, ps() As Double
    Dim extrIdx() As Long, extrCnt As Long
    Dim prevDir As Integer, curDir As Integer, dP As Double, noise As Double
    Dim amp As Double, maxAmp As Double, cntBig As Long, bestRun As Long, run As Long
    Dim fTied As Long, j As Long, fLo As Double, fHi As Double, k As Long
    Dim s As Double, c As Long

    res.OscViolation = False
    res.OscPeriods = 0
    res.OscAmplPct = 0
    res.OscNote = ""
    n = evEnd - startRow + 1
    If n < 20 Or g.PNom <= 0 Then Exit Sub
    If n > 5000 Then n = 5000

    ReDim p(1 To n): ReDim f(1 To n): ReDim ps(1 To n)
    For i = 1 To n
        r = startRow + i - 1
        p(i) = NzD(wsRaw.Cells(r, pCol).Value, 0)
        f(i) = NzD(wsRaw.Cells(r, fCol).Value, st.FNom)
    Next i
    For i = 1 To n
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= n Then s = s + p(j): c = c + 1
        Next j
        ps(i) = s / c
    Next i

    noise = 0.002 * g.PNom   ' 0.2 %Pном - порог значимого изменения
    ReDim extrIdx(1 To n)
    extrCnt = 0
    prevDir = 0
    For i = 2 To n
        dP = ps(i) - ps(i - 1)
        If dP > noise Then
            curDir = 1
        ElseIf dP < -noise Then
            curDir = -1
        Else
            curDir = 0
        End If
        If prevDir <> 0 And curDir <> 0 And curDir <> prevDir Then
            extrCnt = extrCnt + 1
            extrIdx(extrCnt) = i - 1
        End If
        If curDir <> 0 Then prevDir = curDir
    Next i
    If extrCnt < 2 Then Exit Sub

    maxAmp = 0: cntBig = 0: bestRun = 0: run = 0
    For i = 2 To extrCnt
        amp = Abs(ps(extrIdx(i)) - ps(extrIdx(i - 1))) / 2#
        If amp > maxAmp Then maxAmp = amp
        If amp > OSC_MIN_AMPL_PCT / 100# * g.PNom Then
            run = run + 1
            If run > bestRun Then bestRun = run
            cntBig = cntBig + 1
        Else
            run = 0
        End If
    Next i
    res.OscAmplPct = 100# * maxAmp / g.PNom
    res.OscPeriods = bestRun \ 2    ' два полупериода = один период

    fTied = 0
    For i = 1 To extrCnt
        fLo = 1E+99: fHi = -1E+99
        r = startRow + extrIdx(i) - 1
        k = extrIdx(i)
        Do While k >= 1
            If Abs(SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(startRow + k - 1, timeCol).Value)) > 3.0001 Then Exit Do
            If f(k) < fLo Then fLo = f(k)
            If f(k) > fHi Then fHi = f(k)
            k = k - 1
        Loop
        k = extrIdx(i) + 1
        Do While k <= n
            If Abs(SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(startRow + k - 1, timeCol).Value)) > 3.0001 Then Exit Do
            If f(k) < fLo Then fLo = f(k)
            If f(k) > fHi Then fHi = f(k)
            k = k + 1
        Loop
        If fHi - fLo > QS_TOL_HZ Then fTied = fTied + 1
    Next i

    If res.OscPeriods >= OSC_MIN_PERIODS And res.OscAmplPct > OSC_MIN_AMPL_PCT Then
        If fTied * 2 > extrCnt Then
            res.OscNote = "колебания мощности (" & res.OscPeriods & " периодов, амплитуда до " & Format(res.OscAmplPct, "0.0") & _
                          " %Pном) связаны с колебаниями частоты - не нарушение по п. 4.2.14.4"
        Else
            res.OscViolation = True
            res.OscNote = "незатухающие колебания мощности: " & res.OscPeriods & " периодов подряд с амплитудой до " & _
                          Format(res.OscAmplPct, "0.0") & " %Pном (> 1 %Pном), не связанные с частотой"
        End If
    ElseIf res.OscPeriods > 0 And res.OscAmplPct > OSC_MIN_AMPL_PCT Then
        res.OscNote = "колебания мощности " & res.OscPeriods & " пер., амплитуда до " & Format(res.OscAmplPct, "0.0") & " %Pном (п. 4.2.6, < 5 периодов)"
    End If
End Sub


Public Sub EvaluateTelemetry(ByVal wsRaw As Worksheet, ByRef g As TGenCfg, ByRef res As TGenResult, _
                              ByVal pCol As Long, ByVal timeCol As Long, ByVal startRow As Long, ByVal evEnd As Long)
    Dim r As Long, r1 As Long, total As Long, blanks As Long, run As Long, maxRun As Long
    Dim prevV As Double, curV As Double, hasPrev As Boolean, spikes As Long, note As String
    Dim lastR As Long
    res.TelemetryOK = True
    res.TelemetryNote = ""
    lastR = LastUsedRow(wsRaw)
    r1 = RowByTimeOffset(wsRaw, timeCol, startRow, -15#)
    If r1 < 2 Then r1 = 2
    If evEnd > lastR Then evEnd = lastR
    hasPrev = False
    For r = r1 To evEnd
        total = total + 1
        If Not IsNumeric(wsRaw.Cells(r, pCol).Value) Or Len(Trim$(CStr(wsRaw.Cells(r, pCol).Value))) = 0 Then
            blanks = blanks + 1
            run = 0
            hasPrev = False
        Else
            curV = CDbl(wsRaw.Cells(r, pCol).Value)
            If hasPrev Then
                If Abs(curV - prevV) < 0.0000001 Then
                    run = run + 1
                    If run > maxRun Then maxRun = run
                Else
                    run = 0
                End If
                If g.PNom > 0 Then
                    If Abs(curV - prevV) > 0.2 * g.PNom Then spikes = spikes + 1
                End If
            End If
            prevV = curV
            hasPrev = True
        End If
    Next r
    If total = 0 Then Exit Sub
    note = ""
    If blanks > 0 Then
        note = note & "пропуски ТИ мощности: " & blanks & " из " & total & " точек; "
        If blanks * 10 > total Then res.TelemetryOK = False
    End If
    If maxRun + 1 > TELEM_MAX_REPEAT Then
        note = note & "одинаковые значения мощности " & (maxRun + 1) & " точек подряд (> " & TELEM_MAX_REPEAT & ", возможное «замораживание» ТИ); "
        res.TelemetryOK = False
    End If
    If spikes > 0 Then
        note = note & "скачки мощности > 20 %Pном за шаг: " & spikes & "; "
        If spikes > 2 Then res.TelemetryOK = False
    End If
    If Len(note) > 0 Then
        If Right$(note, 2) = "; " Then note = Left$(note, Len(note) - 2)
        res.TelemetryNote = IIf(res.TelemetryOK, "ТИ: ", "ТИ сомнительны: ") & note
    End If
End Sub


Public Sub ClassifyReactionType(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                 ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                 ByVal startRow As Long, ByVal evEnd As Long)
    Dim needSign As Integer, reserveNeed As Double, thr1 As Double
    Dim r As Long, dP As Double, maxGood As Double, maxBad As Double, rMaxGood As Long
    Dim dPTr As Double, code As Long, tName As String
    Dim lastR As Long

    res.ReactionCode = 0
    res.ReactionType = ""
    res.QualClauseVerdict = ""
    needSign = SgnNZ(res.PReqOrig)
    If needSign = 0 Then needSign = -SgnNZ(res.DfMaxEvent)
    If needSign = 0 Then Exit Sub

    thr1 = 0.01 * g.PNom
    reserveNeed = IIf(needSign = 1, res.ReservePlus, res.ReserveMinus)
    lastR = LastUsedRow(wsRaw)
    If evEnd > lastR Then evEnd = lastR

    maxGood = 0: maxBad = 0: rMaxGood = startRow
    For r = startRow To evEnd
        dP = needSign * (NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0)
        If dP > maxGood Then maxGood = dP: rMaxGood = r
        If -dP > maxBad Then maxBad = -dP
    Next r
    dPTr = needSign * res.PFact

    If Not res.TelemetryOK And maxGood < thr1 And maxBad < thr1 Then
        code = 8: tName = "телеизмерение"
    ElseIf reserveNeed < thr1 Then
        code = 7: tName = "без резерва"
    ElseIf res.OscViolation Then
        code = 5: tName = "колебания"
    ElseIf maxBad > thr1 And maxGood < thr1 Then
        code = 4: tName = "противоположная"
    ElseIf maxGood < thr1 Then
        code = 6: tName = "нет реакции"
    ElseIf res.React15Status = "ОК" And dPTr < 0.5 * maxGood And dPTr < Abs(res.PReq) - thr1 And res.Dfr <> 0 Then
        code = 2: tName = "с провалом"
    ElseIf res.React15Status = "Нарушение" And dPTr >= 0.5 * Abs(res.PReqStatic) Then
        code = 3: tName = "котельная"
    ElseIf res.React15Status = "Нарушение" And maxGood >= thr1 Then
        code = 3: tName = "котельная"
    Else
        code = 1: tName = "адекватная"
    End If
    res.ReactionCode = code
    res.ReactionType = code & " - «" & tName & "»"

    If Not res.QualApplicable Then Exit Sub
    If code = 1 Then
        res.QualClauseVerdict = "реакция «адекватная» - требования п. " & IIf(g.HasRM, "4.4.4", "4.4.3") & " выполнены"
    ElseIf code = 7 Then
        res.QualClauseVerdict = "резерв на изменение мощности отсутствует - отсутствие реакции не является нарушением (п. " & IIf(g.HasRM, "4.4.4", "4.4.3") & ")"
    ElseIf code = 8 Then
        res.QualClauseVerdict = "качественная оценка невозможна из-за качества ТИ (тип 8 по п. 4.4.1)"
    ElseIf g.HasRM Then
        res.QualClauseVerdict = "НАРУШЕНИЕ: отсутствие реакции типа «адекватная» (тип " & res.ReactionType & ") при наличии резерва " & _
                                Format(reserveNeed, "0.0") & " МВт - п. 4.4.4"
    Else
        res.QualClauseVerdict = "зафиксирован факт отсутствия реакции типа «адекватная» (тип " & res.ReactionType & ") при резерве " & _
                                Format(reserveNeed, "0.0") & " МВт - п. 4.4.3 (нарушение фиксируется при >= 3 фактах за календарный год)"
    End If
End Sub


Public Sub EvaluateQualitative(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, ByVal f0 As Double)
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

    If Not hit10 Then
        Dim dFrEnd As Double, preqEnd As Double, haveEnd As Double
        dFrEnd = DeadbandDeviation(NzD(wsRaw.Cells(res.EndQualRow, fCol).Value, f0) - f0, g.Fnch)
        preqEnd = PReqCalc(g.SPct, g.PNom, st.FNom, _
                  DynamicKdByTime(g.EquipType, SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, _
                                                          wsRaw.Cells(res.EndQualRow, timeCol).Value)), dFrEnd, st.MaxReqPct)
        preqEnd = ClampReserve(preqEnd, res.ReservePlus, res.ReserveMinus)
        haveEnd = NzD(wsRaw.Cells(res.EndQualRow, pCol).Value, 0) - res.P0
        If SgnNZ(preqEnd) = 0 Or SgnNZ(preqEnd) <> signReq Then
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

    res.PReqSteady = ComputeSteadyPReqMean(wsRaw, fCol, timeCol, res.EndQualRow, _
                                           st.SteadyWindowSec, f0, st.FNom, g.SPct, g.PNom, g.Fnch, _
                                           g.EquipType, wsRaw.Cells(res.StartRow, timeCol).Value, st.MaxReqPct)
    If Not st.IsEmergency And Abs(res.PReqStatic) > Abs(res.PReqSteady) Then
        res.PReqSteady = res.PReqStatic
    End If
    res.PReqSteady = ClampReserve(res.PReqSteady, res.ReservePlus, res.ReserveMinus)

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

Public Function ComputeSteadyMean(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
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

Public Function ComputeSteadyPReqMean(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal timeCol As Long, _
                                        ByVal endRow As Long, ByVal windowSec As Double, _
                                        ByVal f0 As Double, ByVal fNom As Double, ByVal sPct As Double, ByVal pNom As Double, _
                                        ByVal fnch As Double, _
                                        ByVal equipType As String, ByVal tStart As Variant, _
                                        ByVal maxReqPct As Double) As Double
    Dim startRow As Long, r As Long
    Dim sumP As Double, cnt As Long
    Dim dFr As Double, fv As Double, kdEff As Double, tSec As Double
    If windowSec <= 0 Then windowSec = 30
    If sPct <= 0 Then Exit Function
    If f0 <= 0 Then f0 = fNom
    startRow = RowByTimeOffset(wsRaw, timeCol, endRow, -windowSec)
    If startRow < 2 Then startRow = 2
    If startRow >= endRow Then startRow = endRow
    For r = startRow To endRow
        If IsNumeric(wsRaw.Cells(r, freqCol).Value) Then
            fv = CDbl(wsRaw.Cells(r, freqCol).Value) - f0
            dFr = DeadbandDeviation(fv, fnch)
            tSec = SecBetween(tStart, wsRaw.Cells(r, timeCol).Value)
            kdEff = DynamicKdByTime(equipType, tSec)
            sumP = sumP + PReqCalc(sPct, pNom, fNom, kdEff, dFr, maxReqPct)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then ComputeSteadyPReqMean = sumP / cnt Else ComputeSteadyPReqMean = 0
End Function


Public Sub EvaluateTransient(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
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
