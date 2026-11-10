Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' Точки входа AnalyzeOPRCH. Импорт 6.
' OPRCH_Monitor v1.8.5 — раздельный модуль (лимит VBA ~64 КБ).

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

    wsCfg.Range("A1:V1").Value = Array( _
        "Станция", "Генератор", "Колонка_мощности", "Колонка_частоты", "Тип_оборудования", _
        "Pном, МВт", "S, %", "fнч, Гц", "Вкл (1/0)", "Кач_вкл (1/0)", _
        "t5, c", "dP5, %Pном", "t10, c", "dP10, %Pном", "Уст_допуск, %Pном", _
        "В сумму станции (1/0)", "Контр_уст (1/0)", "Паропровод", _
        "Pmax, МВт", "Pmin, МВт", "РМ (1/0)", "Компоновка (блок/неблок)" _
    )

    wsCfg.Cells(2, 1).Resize(1, 22).Value = Array("Сосногорская ТЭЦ", "ТГ-5", "ТГ-5", "Частота", "ПТУ_неблок", 55, 4.2, 0.105, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 55, 15, 1, "неблок")
    wsCfg.Cells(3, 1).Resize(1, 22).Value = Array("Сосногорская ТЭЦ", "ТГ-7", "ТГ-7", "Частота", "ПТУ_неблок", 60, 4.5, 0.11, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 60, 20, 1, "неблок")
    wsCfg.Cells(4, 1).Resize(1, 22).Value = Array("ТЭЦ СЛПК", "ТГ-2Э", "ТГ-2Э", "Частота", "ПТУ_неблок", 50, 4.5, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "Э", 50, 15, 1, "неблок")
    wsCfg.Cells(5, 1).Resize(1, 22).Value = Array("ТЭЦ СЛПК", "ТГ-5У", "ТГ-5У", "Частота", "ПТУ_неблок", 87.7, 4.2, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "У", 87.7, 25, 1, "неблок")

    wsCfg.Range("W1").Value = "Глобальные настройки"
    wsCfg.Cells(2, 23).Resize(1, 2).Value = Array("fном, Гц", 50)
    wsCfg.Cells(3, 23).Resize(1, 2).Value = Array("Время начала события", "")
    wsCfg.Cells(4, 23).Resize(1, 2).Value = Array("Автопоиск старта (1/0)", 1)
    wsCfg.Cells(5, 23).Resize(1, 2).Value = Array("Запасной интервал tр (если нет QS), с", 300)
    wsCfg.Cells(6, 23).Resize(1, 2).Value = Array("Допуск количеств., %", 10)
    wsCfg.Cells(7, 23).Resize(1, 2).Value = Array("Порог включения в работу, МВт", 1)
    wsCfg.Cells(8, 23).Resize(1, 2).Value = Array("Pre-start буфер, с (реком. >=300)", 15)
    wsCfg.Cells(9, 23).Resize(1, 2).Value = Array("Интервал графика, с", 120)
    wsCfg.Cells(10, 23).Resize(1, 2).Value = Array("Окно установив./конец события, с", 30)
    wsCfg.Cells(11, 23).Resize(1, 2).Value = Array("Порог реакции 15 с, %Pном", 0.5)
    wsCfg.Cells(12, 23).Resize(1, 2).Value = Array("Усреднение P0/Pтек, с (0=текущие)", 0)
    wsCfg.Cells(13, 23).Resize(1, 2).Value = Array("Режим (Авария/Испытание)", "Авария")
    wsCfg.Cells(14, 23).Resize(1, 2).Value = Array("Макс. окно события, с", 600)
    wsCfg.Cells(15, 23).Resize(1, 2).Value = Array("Макс. число событий (1=только первое)", 3)
    wsCfg.Cells(16, 23).Resize(1, 2).Value = Array("Ограничение Pтреб, %Pном (0=нет)", 10)
    wsCfg.Cells(17, 23).Resize(1, 2).Value = Array("Строгий дефицит 4.2.14.1 (1/0)", 1)
    wsCfg.Range("X13").NumberFormat = "@"
    WriteInstructionSheet
    EnsureSheet(SH_LOG)
    EnsureSheet(SH_SUM)

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

' Экспорт отчёта (Summary, Протокол, Log и листы графиков) в PDF рядом с книгой

Public Sub ExportOPRCHReport()
    On Error GoTo EH
    Dim ws As Worksheet, names As Collection, i As Long
    Dim arr() As String, pdfPath As String, baseName As String
    Set names = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SH_SUM Or ws.Name = SH_PROT Or ws.Name = SH_LOG Then
            names.Add ws.Name
        ElseIf Right$(ws.Name, Len(CHART_SUFFIX)) = CHART_SUFFIX Or Right$(ws.Name, 3) = "_Гр" Then
            names.Add ws.Name
        End If
    Next ws
    If names.Count = 0 Then
        MsgBox "Нет листов для экспорта. Сначала запустите мониторинг.", vbExclamation
        Exit Sub
    End If
    ReDim arr(0 To names.Count - 1)
    For i = 1 To names.Count
        arr(i - 1) = CStr(names(i))
    Next i
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Сначала сохраните книгу - PDF создаётся в той же папке.", vbExclamation
        Exit Sub
    End If
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & baseName & "_Отчет_ОПРЧ_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    For i = 0 To UBound(arr)
        With ThisWorkbook.Worksheets(arr(i)).PageSetup
            On Error Resume Next
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            On Error GoTo EH
        End With
    Next i
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets(SH_SUM).Select
    MsgBox "Отчёт сохранён: " & pdfPath, vbInformation
    Exit Sub
EH:
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SUM).Select
    MsgBox "Ошибка экспорта в PDF: " & Err.Description, vbCritical
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
    Dim evList() As TEvent, nEv As Long, k As Long
    Dim fCol As Long

    t0Run = Timer
    stepName = "Подготовка листов"
    Set wsRaw = GetRequiredSheet(SH_RAW)
    Set wsCfg = GetRequiredSheet(SH_CFG)
    Set wsSummary = EnsureSheet(SH_SUM)

    ' Кнопки сразу - даже если дальше будет ошибка, интерфейс уже на месте
    stepName = "Интерфейс (кнопки)"
    On Error Resume Next
    EnsureControlButtons wsCfg
    WriteInstructionSheet
    On Error GoTo EH

    stepName = "Чтение настроек"
    st = ReadSettings(wsCfg)
    LoadKdProfiles wsCfg
    LoadChartDisplayOptions wsCfg
    On Error Resume Next
    EnsureChartOptionCheckboxes wsCfg
    On Error GoTo EH
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Err.Raise vbObjectError + 2001, , "В RawData не найдена колонка 'Время'."

    cfgLast = LastUsedRow(wsCfg)
    If cfgLast < 2 Then Err.Raise vbObjectError + 2002, , "В Config нет строк генераторов."

    stepName = "Подготовка лога"
    InitLog
    Set m_Protocol = New Collection
    AppendLog "INFO", "Settings", "Режим: " & st.Mode & IIf(st.IsEmergency, _
        " (для оборудования без РМ принимаются S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц)", _
        " (используются фактические S/fнч из Config; применимость по п. 4.1.2/4.1.3 не проверяется)")

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
    wsSummary.Range("A1:AZ1").Value = Array( _
        "Станция", "Генератор", "Тип", _
        "Старт (расч.)", "Время выхода за fнч", _
        "P0, МВт", "Pтек(tр), МВт", "Pуст_сред, МВт", _
        "dF, Гц", "dFr(tр), Гц", _
        "Pтреб(tр), МВт", "Pфакт(tр), МВт", _
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
        "Вывод", _
        "Событие", "Применимость (п. 4.1.2/4.1.3)", "tр", "tр, с от старта", "Режим в tр", _
        "Pтреб(Tинт), МВт", "Pфакт(Tинт), МВт", _
        "Колебания (п. 4.2.14.4)", "Качество ТИ", "Тип реакции (п. 4.4.1)", _
        "h(15 с), Прил. п. 6.2", "Кач. вывод (п. 4.4.3/4.4.4)" _
    )
    outRow = 2

    For r = 2 To cfgLast
        stepName = "Чтение Config, строка " & r
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextGen

        g = ReadGenCfgEff(wsCfg, r, st)
        If Not g.Enabled Then GoTo NextGen

        If Not ValidateGenCfg(g) Then
            WriteSummaryInvalid wsSummary, outRow, g
            AppendLog "Config", g.Station & "/" & g.Generator, "Пропущен: не заполнены обязательные параметры"
            outRow = outRow + 1
            GoTo NextGen
        End If

        stepName = "Поиск событий для " & g.Generator
        fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
        If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator
        ' Для поиска границ: с РМ — fном; без РМ — частота в начале выборки (уточняется до f(t0) в анализе).
        nEv = FindEvents(wsRaw, timeCol, fCol, st, g.Fnch, _
                         IIf(g.HasRM, st.FNom, NzD(wsRaw.Cells(2, fCol).Value, st.FNom)), evList)
        If nEv > 1 Then AppendLog "INFO", g.Station & "/" & g.Generator, "Найдено событий (выходов за fнч): " & nEv

        For k = 1 To nEv
            stepName = "Расчет генератора " & g.Generator & " (событие " & k & ")"
            res = AnalyzeOneGenerator(wsRaw, st, g, evList(k), k)
            stepName = "Запись листа данных " & g.Generator
            WriteGeneratorSheet wsRaw, st, g, res
            stepName = "Запись графика " & g.Generator
            WriteGeneratorChartSheet st, g, res
            stepName = "Запись Summary для " & g.Generator
            WriteSummaryRow wsSummary, outRow, g, res
            AddProtocolGenerator g, res
            outRow = outRow + 1
        Next k

NextGen:
    Next r

    stepName = "Расчет суммарных листов станций"
    BuildStationAggregates wsRaw, wsCfg, wsSummary, st, outRow

    stepName = "Оформление Summary"
    wsSummary.Columns("A:AZ").AutoFit
    wsSummary.Columns("AN").ColumnWidth = 90
    wsSummary.Columns("AZ").ColumnWidth = 70
    wsSummary.Columns("AI").ColumnWidth = 60
    wsSummary.Columns("AP").ColumnWidth = 45
    wsSummary.Range("D:E").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsSummary.Range("AQ:AQ").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ApplySummaryConditionalFormat wsSummary
    WriteVersionStamp wsSummary, wsRaw, t0Run

    stepName = "Формирование протокола"
    WriteProtocolSheet st, wsRaw

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
        If Len(Trim$(CStr(wsCfg.Cells(r, 21).Value))) = 0 Then wsCfg.Cells(r, 21).Value = 1: changed = changed + 1
        If Len(Trim$(CStr(wsCfg.Cells(r, 22).Value))) = 0 Then
            wsCfg.Cells(r, 22).Value = IIf(InStr(UCase$(et), "НЕБЛОК") > 0, "неблок", "блок"): changed = changed + 1
        End If
NX:
    Next r
    MsgBox "Пресеты применены. Заполнено ячеек: " & changed, vbInformation
End Sub

' ==========================================================
' Анализ одного генератора
' ==========================================================

Public Sub WriteInstructionSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_INSTR)
    ws.Cells.Clear
    ws.Range("A1").Value = "Мониторинг ОПРЧ v" & OPRCH_VERSION & " — универсальный шаблон (любая станция / любое событие)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "0. КАК ПОЛЬЗОВАТЬСЯ (универсально)"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Программа не привязана к одной электростанции. Для каждого события: заполните RawData (выгрузка ОИК/архива) и Config (состав оборудования)."
    ws.Range("A5").Value = "Отдельной вкладки «Запрос» нет — код запроса к ОИК храните у себя; в книгу кладутся только измеренные ряды и параметры расчёта."
    ws.Range("A7").Value = "1. ИСХОДНЫЕ ДАННЫЕ (обязательно)"
    ws.Range("A7").Font.Bold = True
    ws.Range("A8").Value = "• Лист RawData: колонка «Время» (шаг <= 5 с) + частота + активные мощности генераторов (МВт). Имена колонок — любые, главное совпадение с Config."
    ws.Range("A9").Value = "• Окно данных по п. 3.1 МУ: желательно >= 5 мин до начала отклонения и до устойчивого возврата в 50±0.1 Гц более 5 мин."
    ws.Range("A10").Value = "• Лист Config: по каждому генератору — станция, имя, заголовки колонок мощности/частоты, тип, Pном, S, fнч, РМ (1/0), компоновка (блок/неблок)."
    ws.Range("A11").Value = "• Пустые нормативные поля (t5/t10, fнч…) — кнопка «Применить пресеты типов»."
    ws.Range("A12").Value = "• Глобальные настройки — колонки W/X: режим Авария/Испытание, fном, автопоиск старта события, строгий дефицит (W17)."
    ws.Range("A14").Value = "2. ОДИН ЗАПУСК = ВСЁ СРАЗУ"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15").Value = "Кнопка «Запустить мониторинг ОПРЧ» (AnalyzeOPRCH) — количественная и качественная оценка в одном проходе."
    ws.Range("A16").Value = "Отдельно «только качественно» / «только количественно» не нужно: неприменимое отсекается по п. 4.1.2 / 4.1.3 (в режиме «Авария»)."
    ws.Range("A17").Value = "• Колич. (Прил. 14 п. 4.2): P0, f0, Pтреб(t), tр, дефицит по периодам, реакция 15 с, сумма неблока."
    ws.Range("A18").Value = "• Кач. (Прил. 14 п. 4.4): тип реакции 1-8. Ступени t5/t10 (Прил. 1) — при «Кач_вкл=1»."
    ws.Range("A20").Value = "3. РЕЖИМЫ РАСЧЁТА (Config W13)"
    ws.Range("A20").Font.Bold = True
    ws.Range("A21").Value = "• Авария — реальные события (Прил. 14). Без РМ: S=6 %, fнч=0.15 Гц. Дефицит «хотя бы в одном периоде» -> НАРУШЕНИЕ. t5/t10 только в замечаниях."
    ws.Range("A22").Value = "• Испытание — контрольные испытания (Прил. 1 + формулы Прил. 14). S/fнч из Config. Несданные t5/t10 -> НАРУШЕНИЕ."
    ws.Range("A24").Value = "4. РЕЗУЛЬТАТЫ"
    ws.Range("A24").Font.Bold = True
    ws.Range("A25").Value = "• Summary — сводная таблица; колонка «Вывод» — главный вердикт."
    ws.Range("A26").Value = "• Листы <Ст>_<Ген> / <Ст>_Сумма и *_Граф — ряды и графики; Протокол — текстовая сводка; Log — предупреждения."
    ws.Range("A27").Value = "• Для неблочной ТЭС нормативный вывод — по сумме станции (п. 4.2.10); строки генераторов при ОК по сумме — справочно."
    ws.Range("A29").Value = "5. ЗАПУСК"
    ws.Range("A29").Font.Bold = True
    ws.Range("A30").Value = "Разрешите макросы -> лист Config -> кнопка «Запустить мониторинг ОПРЧ» (или Alt+F8 / Option+F8 -> AnalyzeOPRCH)."
    ws.Range("A31").Value = "Если кнопок нет: Alt+F8 -> InitOPRCHUI. SetupOPRCHTemplate на заполненной книге не запускайте - он перезапишет RawData/Config."
    ws.Columns("A").ColumnWidth = 140
    On Error Resume Next
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    On Error GoTo 0
End Sub

' Создать/обновить кнопки и лист «Инструкция» (без пересчёта).
' Alt+F8 / Option+F8 -> InitOPRCHUI. Также вызывается при открытии книги.

Public Sub InitOPRCHUI()
    On Error GoTo EH
    Dim wsCfg As Worksheet
    WriteInstructionSheet
    Set wsCfg = EnsureSheet(SH_CFG)
    EnsureControlButtons wsCfg
    On Error Resume Next
    EnsureControlButtons EnsureSheet(SH_INSTR)
    On Error GoTo 0
    Exit Sub
EH:
    MsgBox "Ошибка InitOPRCHUI: " & Err.Description, vbCritical
End Sub

Public Sub InitOPRCHUISilent()
    On Error Resume Next
    WriteInstructionSheet
    EnsureControlButtons EnsureSheet(SH_CFG)
    EnsureControlButtons EnsureSheet(SH_INSTR)
End Sub

Public Sub EnsureControlButtons(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    AddOrReplaceButton ws, "btnRunOPRCH", ws.Range("E30"), 260, 32, _
        "Запустить мониторинг ОПРЧ", "AnalyzeOPRCH", RGB(40, 120, 220), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnClearOPRCH", ws.Range("I30"), 200, 32, _
        "Очистить результаты", "ClearOPRCHResults", RGB(150, 150, 150), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnPresets", ws.Range("M30"), 220, 32, _
        "Применить пресеты типов", "ApplyPresetsToConfig", RGB(80, 160, 100), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnExportPdf", ws.Range("Q30"), 200, 32, _
        "Экспорт отчёта в PDF", "ExportOPRCHReport", RGB(160, 90, 160), RGB(255, 255, 255)
End Sub

Public Sub AddOrReplaceButton(ByVal ws As Worksheet, ByVal nm As String, ByVal anchor As Range, _
                               ByVal w As Double, ByVal h As Double, _
                               ByVal caption As String, ByVal onAction As String, _
                               ByVal fillColor As Long, ByVal textColor As Long)
    Dim shp As Shape
    Dim macroRef As String
    On Error Resume Next
    ws.Shapes(nm).Delete
    Err.Clear
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, w, h)
    If shp Is Nothing Then
        ' запасной вариант: кнопка формы (стабильнее на части сборок Excel)
        Dim btn As Button
        Set btn = ws.Buttons.Add(anchor.Left, anchor.Top, w, h)
        If Not btn Is Nothing Then
            btn.Name = nm
            btn.Caption = caption
            btn.OnAction = "'" & ThisWorkbook.Name & "'!" & onAction
        End If
        Exit Sub
    End If
    shp.Name = nm
    macroRef = "'" & ThisWorkbook.Name & "'!" & onAction
    shp.OnAction = macroRef
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(70, 70, 70)
    ' TextFrame совместимее TextFrame2 на Windows Excel
    Err.Clear
    shp.TextFrame.Characters.Text = caption
    If Err.Number <> 0 Then
        Err.Clear
        shp.TextFrame2.TextRange.Text = caption
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
        shp.TextFrame2.TextRange.Font.Size = 11
    Else
        shp.TextFrame.Characters.Font.Color = textColor
        shp.TextFrame.Characters.Font.Size = 11
        shp.TextFrame.HorizontalAlignment = xlHAlignCenter
        shp.TextFrame.VerticalAlignment = xlVAlignCenter
    End If
    On Error GoTo 0
End Sub
