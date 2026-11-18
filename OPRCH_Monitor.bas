Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' Точки входа AnalyzeOPRCH. Импорт 6.
' OPRCH_Monitor v1.8.6 — раздельный модуль (лимит VBA ~64 КБ).

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

Public Function PdfTextHeight(ByVal text As String, ByVal widthChars As Double, ByVal fontSize As Double) As Double
    Dim perLine As Long, lines As Long, i As Long, chunk As Long
    perLine = CLng(widthChars * 0.95)
    If perLine < 24 Then perLine = 24
    lines = 1
    chunk = 0
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = vbLf Or Mid$(text, i, 1) = vbCr Then
            lines = lines + 1
            chunk = 0
        Else
            chunk = chunk + 1
            If chunk >= perLine Then
                lines = lines + 1
                chunk = 0
            End If
        End If
    Next i
    If lines > 14 Then lines = 14
    If lines < 1 Then lines = 1
    PdfTextHeight = lines * (fontSize + 5) + 6
    If PdfTextHeight < 18 Then PdfTextHeight = 18
End Function

Public Sub PdfColorStatus(ByVal cell As Range)
    Dim s As String
    s = CStr(cell.Value)
    If InStr(1, s, "Нарушение", vbTextCompare) > 0 Or Left$(s, 9) = "НАРУШЕНИЕ" Then
        cell.Font.Color = RGB(192, 0, 0)
        cell.Font.Bold = True
    ElseIf InStr(1, s, "Ограничен", vbTextCompare) > 0 Then
        cell.Font.Color = RGB(156, 87, 0)
        cell.Font.Bold = True
    ElseIf Left$(s, 2) = "ОК" Then
        cell.Font.Color = RGB(0, 97, 0)
        cell.Font.Bold = True
    ElseIf InStr(1, s, "Н/П", vbTextCompare) > 0 Then
        cell.Font.Color = RGB(89, 89, 89)
    End If
End Sub

Public Sub PdfPageSetup(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 999
        .LeftMargin = Application.InchesToPoints(0.45)
        .RightMargin = Application.InchesToPoints(0.45)
        .TopMargin = Application.InchesToPoints(0.55)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.25)
        .CenterHorizontally = True
        .PrintGridlines = False
        .LeftFooter = "Мониторинг ОПРЧ"
        .RightFooter = "Стр. &P из &N"
    End With
    On Error GoTo 0
End Sub

Public Function PdfVerdictText(ByVal ws As Worksheet) As String
    Dim a As String, b As String
    a = Trim$(CStr(ws.Range("B12").Value))
    b = Trim$(CStr(ws.Range("B7").Value))
    If Left$(a, 2) = "ОК" Or Left$(a, 3) = "НАР" Or Left$(a, 3) = "НЕ " Then
        PdfVerdictText = a
    ElseIf Left$(b, 2) = "ОК" Or Left$(b, 3) = "НАР" Or Left$(b, 3) = "НЕ " Then
        PdfVerdictText = b
    ElseIf Len(a) > 0 Then
        PdfVerdictText = a
    Else
        PdfVerdictText = b
    End If
End Function

Public Function PdfHasExportData() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    If LastUsedRow(ThisWorkbook.Worksheets(SH_SUM)) >= 2 Then
        PdfHasExportData = True
        Exit Function
    End If
    For Each ws In ThisWorkbook.Worksheets
        If Right$(ws.Name, Len(CHART_SUFFIX)) = CHART_SUFFIX Or Right$(ws.Name, 3) = "_Гр" Then
            If ws.ChartObjects.Count > 0 Then
                PdfHasExportData = True
                Exit Function
            End If
        End If
    Next ws
    PdfHasExportData = False
End Function

Public Sub PdfBuildSummary(ByVal ws As Worksheet)
    Dim src As Worksheet, lastR As Long, r As Long, outR As Long
    Dim txt As String, h As Double
    ws.Name = "Сводка"
    ws.Cells.Clear
    Set src = ThisWorkbook.Worksheets(SH_SUM)
    ws.Range("A1").Value = "Сводка мониторинга ОПРЧ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = "Количественная и качественная оценка в одном расчёте. Длинный текст переносится, таблица умещается в ширину альбомного листа."
    ws.Range("A2").Font.Size = 10
    ws.Range("A2:K2").Merge
    ws.Range("A2").WrapText = True
    ws.Rows(2).RowHeight = 32
    ws.Range("A4:K4").Value = Array( _
        "Станция", "Генератор", "Колич.", "%", "Кач.", _
        "t5", "t10", "уст", "Реакция 15 с", "Дефицит", "Вывод")
    ws.Range("A4:K4").Font.Bold = True
    ws.Range("A4:K4").Font.Color = RGB(255, 255, 255)
    ws.Range("A4:K4").Interior.Color = RGB(31, 78, 121)
    ws.Range("A4:K4").WrapText = True
    ws.Range("A4:K4").VerticalAlignment = xlCenter
    ws.Rows(4).RowHeight = 30
    outR = 5
    lastR = LastUsedRow(src)
    For r = 2 To lastR
        If Len(Trim$(CStr(src.Cells(r, 1).Value))) = 0 And Len(Trim$(CStr(src.Cells(r, 2).Value))) = 0 Then GoTo NextSum
        ws.Cells(outR, 1).Value = src.Cells(r, 1).Value
        ws.Cells(outR, 2).Value = src.Cells(r, 2).Value
        ws.Cells(outR, 3).Value = src.Cells(r, 16).Value
        If IsNumeric(src.Cells(r, 15).Value) Then
            ws.Cells(outR, 4).Value = Format(CDbl(src.Cells(r, 15).Value), "0")
        Else
            ws.Cells(outR, 4).Value = src.Cells(r, 15).Value
        End If
        ws.Cells(outR, 5).Value = src.Cells(r, 20).Value
        ws.Cells(outR, 6).Value = src.Cells(r, 21).Value
        ws.Cells(outR, 7).Value = src.Cells(r, 22).Value
        ws.Cells(outR, 8).Value = src.Cells(r, 23).Value
        ws.Cells(outR, 9).Value = src.Cells(r, 36).Value
        If IsNumeric(src.Cells(r, 38).Value) Then
            ws.Cells(outR, 10).Value = Format(CDbl(src.Cells(r, 38).Value), "0.0")
        Else
            ws.Cells(outR, 10).Value = src.Cells(r, 38).Value
        End If
        ws.Cells(outR, 11).Value = src.Cells(r, 40).Value
        ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).WrapText = True
        ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).VerticalAlignment = xlTop
        ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).Font.Size = 9
        PdfColorStatus ws.Cells(outR, 3)
        PdfColorStatus ws.Cells(outR, 5)
        PdfColorStatus ws.Cells(outR, 6)
        PdfColorStatus ws.Cells(outR, 7)
        PdfColorStatus ws.Cells(outR, 8)
        PdfColorStatus ws.Cells(outR, 9)
        PdfColorStatus ws.Cells(outR, 11)
        If Left$(CStr(src.Cells(r, 2).Value), 5) = "СУММА" Then
            ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).Font.Bold = True
            ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).Interior.Color = RGB(221, 235, 247)
        ElseIf outR Mod 2 = 0 Then
            ws.Range(ws.Cells(outR, 1), ws.Cells(outR, 11)).Interior.Color = RGB(248, 250, 252)
        End If
        txt = CStr(ws.Cells(outR, 11).Value)
        h = PdfTextHeight(txt, 28, 9)
        If h < 22 Then h = 22
        ws.Rows(outR).RowHeight = h
        outR = outR + 1
NextSum:
    Next r
    If outR > 5 Then
        ws.Range(ws.Cells(4, 1), ws.Cells(outR - 1, 11)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(4, 1), ws.Cells(outR - 1, 11)).Borders.Color = RGB(180, 198, 214)
    End If
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 8
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 9
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 9
    ws.Columns("I").ColumnWidth = 13
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 28
    PdfPageSetup ws
    ws.PageSetup.PrintTitleRows = "$4:$4"
End Sub

Public Sub PdfBuildProtocol(ByVal rep As Workbook)
    Dim src As Worksheet, ws As Worksheet
    Dim lastR As Long, r As Long, outR As Long, txt As String
    On Error Resume Next
    Set src = ThisWorkbook.Worksheets(SH_PROT)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub
    lastR = LastUsedRow(src)
    If lastR < 1 Then Exit Sub
    Set ws = rep.Worksheets.Add(After:=rep.Worksheets(rep.Worksheets.Count))
    ws.Name = "Протокол PDF"
    outR = 1
    For r = 1 To lastR
        txt = Trim$(CStr(src.Cells(r, 1).Value))
        If Len(txt) = 0 Then GoTo NextProt
        ws.Cells(outR, 1).Value = txt
        ws.Cells(outR, 1).WrapText = True
        ws.Cells(outR, 1).VerticalAlignment = xlTop
        ws.Cells(outR, 1).Font.Size = 11
        If src.Cells(r, 1).Font.Bold Then ws.Cells(outR, 1).Font.Bold = True
        If src.Cells(r, 1).Font.Color <> 0 Then ws.Cells(outR, 1).Font.Color = src.Cells(r, 1).Font.Color
        ws.Rows(outR).RowHeight = PdfTextHeight(txt, 108, 11)
        outR = outR + 1
NextProt:
    Next r
    ws.Columns("A").ColumnWidth = 112
    PdfPageSetup ws
End Sub

Public Sub PdfBuildCharts(ByVal rep As Workbook)
    Dim wsSrc As Worksheet, ws As Worksheet, co As ChartObject
    Dim r As Long, cap As String, verdict As String, nCharts As Long
    Dim shp As Shape
    Set ws = rep.Worksheets.Add(After:=rep.Worksheets(rep.Worksheets.Count))
    ws.Name = "Графики"
    ws.Columns("A").ColumnWidth = 112
    r = 1
    nCharts = 0
    rep.Activate
    For Each wsSrc In ThisWorkbook.Worksheets
        If Right$(wsSrc.Name, Len(CHART_SUFFIX)) = CHART_SUFFIX Or Right$(wsSrc.Name, 3) = "_Гр" Then
            For Each co In wsSrc.ChartObjects
                If r > 1 Then ws.HPageBreaks.Add Before:=ws.Rows(r)
                cap = Trim$(CStr(wsSrc.Range("A1").Value))
                On Error Resume Next
                If co.Chart.HasTitle Then cap = co.Chart.ChartTitle.Text
                On Error GoTo 0
                ws.Cells(r, 1).Value = cap
                ws.Cells(r, 1).Font.Bold = True
                ws.Cells(r, 1).Font.Size = 14
                ws.Cells(r, 1).WrapText = True
                ws.Rows(r).RowHeight = 24
                r = r + 1
                verdict = PdfVerdictText(wsSrc)
                If Len(verdict) = 0 Then verdict = " "
                ws.Cells(r, 1).Value = verdict
                ws.Cells(r, 1).WrapText = True
                ws.Cells(r, 1).VerticalAlignment = xlTop
                ws.Cells(r, 1).Font.Size = 10
                PdfColorStatus ws.Cells(r, 1)
                ws.Rows(r).RowHeight = PdfTextHeight(verdict, 108, 10)
                r = r + 1
                On Error Resume Next
                co.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                If Err.Number <> 0 Then
                    Err.Clear
                    co.Chart.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
                End If
                rep.Activate
                ws.Activate
                ws.Cells(r, 1).Select
                ws.Paste
                Set shp = Nothing
                Set shp = ws.Shapes(ws.Shapes.Count)
                If Not shp Is Nothing Then
                    shp.LockAspectRatio = False
                    shp.Left = ws.Cells(r, 1).Left + 4
                    shp.Top = ws.Cells(r, 1).Top + 2
                    shp.Width = Application.InchesToPoints(9.6)
                    shp.Height = Application.InchesToPoints(5.15)
                End If
                On Error GoTo 0
                ws.Rows(r & ":" & (r + 25)).RowHeight = 15
                ws.Cells(r + 25, 1).Value = " "
                r = r + 27
                nCharts = nCharts + 1
            Next co
        End If
    Next wsSrc
    If nCharts = 0 Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    PdfPageSetup ws
End Sub

Public Sub InstrHeading(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String)
    ws.Cells(r, 1).Value = title
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    With ws.Cells(r, 1)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 24
    r = r + 1
End Sub

Public Sub InstrNote(ByVal ws As Worksheet, ByRef r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = text
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    With ws.Cells(r, 1)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 11
    End With
    ws.Rows(r).RowHeight = PdfTextHeight(text, 128, 11)
    r = r + 1
End Sub

Public Sub InstrField(ByVal ws As Worksheet, ByRef r As Long, ByVal place As String, ByVal fieldName As String, ByVal howTo As String)
    ws.Cells(r, 1).Value = place
    ws.Cells(r, 2).Value = fieldName
    ws.Cells(r, 3).Value = howTo
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).WrapText = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).VerticalAlignment = xlTop
    ws.Cells(r, 1).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Size = 10
    If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = RGB(242, 246, 250)
    ws.Rows(r).RowHeight = PdfTextHeight(howTo, 74, 10)
    r = r + 1
End Sub

Public Sub ExportOPRCHReport()
    Dim rep As Workbook
    Dim pdfPath As String, baseName As String
    Dim prevScreen As Boolean, i As Long
    On Error GoTo EH
    If Not PdfHasExportData() Then
        MsgBox "Нет результатов для отчёта. Сначала запустите мониторинг.", vbExclamation
        Exit Sub
    End If
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Сначала сохраните книгу — PDF создаётся в той же папке.", vbExclamation
        Exit Sub
    End If
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & baseName & "_Отчет_ОПРЧ_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set rep = Workbooks.Add
    Do While rep.Worksheets.Count > 1
        rep.Worksheets(rep.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    PdfBuildSummary rep.Worksheets(1)
    PdfBuildProtocol rep
    PdfBuildCharts rep
    On Error Resume Next
    For i = 1 To rep.Worksheets.Count
        rep.Worksheets(i).PageSetup.FitToPagesWide = 1
        rep.Worksheets(i).PageSetup.FitToPagesTall = 999
    Next i
    On Error GoTo EH
    rep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    rep.Close SaveChanges:=False
    Set rep = Nothing
    Application.CutCopyMode = False
    Application.ScreenUpdating = prevScreen
    ThisWorkbook.Activate
    MsgBox "Отчёт сохранён:" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
           "Сводка идёт с переносом длинного текста." & vbCrLf & _
           "Каждый график — на отдельном альбомном листе.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not rep Is Nothing Then rep.Close SaveChanges:=False
    ThisWorkbook.Activate
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
    Dim previousCalculation As XlCalculation, previousUpdating As Boolean
    Dim stateSaved As Boolean, failureNumber As Long, failureText As String

    previousCalculation = Application.Calculation
    previousUpdating = Application.ScreenUpdating
    stateSaved = True
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

    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousUpdating

    OPRCHMsg "Мониторинг ОПРЧ завершен (v" & OPRCH_VERSION & "). Время: " _
        & Format(Timer - t0Run, "0.0") & " c.", vbInformation
    Exit Sub

EH:
    failureNumber = Err.Number
    failureText = Err.Description
    On Error Resume Next
    If stateSaved Then
        Application.Calculation = previousCalculation
        Application.ScreenUpdating = previousUpdating
    End If
    AppendLog "ERROR", stepName, "Ошибка #" & CStr(failureNumber) & ": " & failureText
    On Error GoTo 0
    OPRCHMsg "Ошибка AnalyzeOPRCH (" & stepName & ", #" & failureNumber & "): " & _
             failureText & vbCrLf & "Подробности: лист Log.", vbCritical
End Sub

' Автопрогон / песочница: тот же расчёт, без диалогов MsgBox.

Public Sub AnalyzeOPRCHSilent()
    Dim prev As Boolean
    prev = gOPRCHQuiet
    gOPRCHQuiet = True
    On Error GoTo EH
    AnalyzeOPRCH
    gOPRCHQuiet = prev
    Exit Sub
EH:
    gOPRCHQuiet = prev
End Sub

Public Sub ClearOPRCHResults()
    Dim targetSheets As Collection
    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    DeleteOutputSheets targetSheets
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
    Dim ws As Worksheet, r As Long
    Set ws = EnsureSheet(SH_INSTR)
    ws.Cells.Clear
    On Error Resume Next
    ws.Range("A1:C200").UnMerge
    On Error GoTo 0
    r = 1
    ws.Cells(r, 1).Value = "Мониторинг участия в ОПРЧ  —  инструкция  v" & OPRCH_VERSION
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Color = RGB(31, 78, 121)
    ws.Rows(r).RowHeight = 26
    r = 3
    InstrHeading ws, r, "1. Что умеет макрос"
    InstrNote ws, r, "Один запуск считает и количественную оценку (Приложение 14, величина первичной мощности), и качественную (тип реакции по п. 4.4 и ступени динамики Приложения 1). Отдельно нажимать «только количество» или «только качество» не нужно: обе оценки получаются вместе."
    InstrNote ws, r, "На выходе: лист Summary с вердиктом, графики по каждому генератору и по сумме станции, текстовый протокол и, по кнопке, PDF. Макрос подходит для любой станции: в книгу кладут ряды измерений и параметры оборудования, а не код запроса к ОИК."
    r = r + 1
    InstrHeading ws, r, "2. Что заполнить перед запуском"
    ws.Cells(r, 1).Value = "Куда"
    ws.Cells(r, 2).Value = "Поле"
    ws.Cells(r, 3).Value = "Что писать"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = RGB(217, 226, 243)
    r = r + 1
    InstrField ws, r, "RawData", "Время", "Первая строка — заголовки. Колонка «Время», шаг не больше 5 с. Желательно не меньше 5 минут до отклонения частоты и до возврата в 50±0,1 Гц."
    InstrField ws, r, "RawData", "Частота и мощность", "Отдельная колонка частоты (общая или своя у генератора) и отдельная колонка активной мощности каждого генератора, МВт. Имя колонки потом точно повторяется в Config."
    InstrField ws, r, "Config A–B", "Станция, генератор", "Как назвать листы и строки сводки. Одна строка Config — один генератор."
    InstrField ws, r, "Config C–D", "Колонка мощности, колонка частоты", "Точные заголовки из строки 1 листа RawData. Без этого совпадения расчёт останавливается."
    InstrField ws, r, "Config E", "Тип оборудования", "ПТУ_блок, ПТУ_блок_уголь, ПТУ_неблок, ГТУ, ПГУ_утил, ПГУ_сбросн или ГПА. От типа берутся пустые нормативы и профиль Kд(t)."
    InstrField ws, r, "Config F, G, H", "Pном, S %, fнч", "Обязательны для включённого генератора. Pном — в МВт, S — статизм в процентах, fнч — зона нечувствительности в Гц. Пустую fнч заполняет кнопка «Применить пресеты типов»."
    InstrField ws, r, "Config I", "Вкл (1/0)", "1 — генератор участвует в этом запуске. 0 — строка пропускается целиком, ни количественно, ни качественно."
    InstrField ws, r, "Config U, V", "РМ и компоновка", "U: 1 — есть регулятор мощности, 0 — нет. В режиме «Авария» при РМ=0 программа сама ставит S=6 % и fнч=0,15 Гц. V: блок или неблок. Для неблочной части нормативный вывод — по сумме станции."
    InstrField ws, r, "Config S, T", "Pmax, Pmin", "Можно оставить пустыми: тогда Pmax=Pном, Pmin=0. Если запас меньше требуемой первичной мощности, Pтреб ограничивается резервом."
    InstrField ws, r, "Config W/X", "Режим и общие уставки", "X13 — «Авария» или «Испытание». Рядом: fном (обычно 50), автопоиск события, длительность количественного окна, допуск в процентах. Эти ячейки общие на всю книгу, не на одну машину."
    r = r + 1
    InstrHeading ws, r, "3. Количественная оценка"
    InstrNote ws, r, "Количественная оценка отвечает на вопрос: в момент tр выдана ли требуемая первичная мощность с учётом статизма, Kд(t), зоны нечувствительности и резерва. Она считается всегда, если в колонке I стоит 1. Отдельное поле «включить количество» не нужно."
    InstrField ws, r, "Что влияет", "S, fнч, Pном, РМ, Pmax/Pmin", "Из них считаются Pтреб и Pфакт. Профиль Kд(t) берётся из таблицы AA:AI по типу оборудования. Если профиля нет, расчёт останавливается с ошибкой."
    InstrField ws, r, "Config W/X", "Допуск и дефицит", "Допуск количественной оценки — насколько процентов Pфакт может не добрать до Pтреб. При включённом строгом дефиците нехватка первичной мощности хотя бы в одном периоде (п. 4.2.14.1) даёт НАРУШЕНИЕ в режиме «Авария»."
    InstrField ws, r, "Summary", "Колич. статус, Колич. %", "ОК — требуемая мощность выдана. Нарушение — не выдана. Ограничен Pmax/Pmin — требуемое упёрлось в резерв, это не нарушение. Н/П — событие не подходит под п. 4.1.2 (в «Аварии»): слишком слабое или слишком короткое отклонение частоты."
    InstrField ws, r, "Summary", "Реакция 15 с, Дефицит", "Реакция 15 с (п. 4.2.11): изменилась ли мощность в нужную сторону через 15 с после выхода за fнч. Дефицит — наибольшая нехватка и сколько секунд она длилась."
    r = r + 1
    InstrHeading ws, r, "4. Качественная оценка"
    InstrNote ws, r, "Качественная оценка отвечает на два разных вопроса. Первый: какой тип реакции по п. 4.4 (адекватная или нет). Второй: успела ли мощность пройти ступени t5 и t10 и удержаться у цели (Приложение 1). Ступени включаются колонкой J."
    InstrField ws, r, "Config J", "Кач_вкл (1/0)", "1 — проверять ступени t5, t10 и установившееся значение. 0 — качественную проверку ступеней не делать. Количественная оценка при этом всё равно считается."
    InstrField ws, r, "Config K–O", "t5, dP5, t10, dP10, допуск уст.", "Времена ступеней в секундах и доли мощности в % Pном. Пустые клетки заполняет кнопка «Применить пресеты типов» по колонке E. Свои числа, если они уже введены, кнопка не затирает."
    InstrField ws, r, "Config Q", "Контр_уст (1/0)", "1 — сравнивать установившуюся мощность с целью. 0 — установившееся значение не контролировать. Для установившегося отклонения в сводке используется оранжевый цвет: это замечание, а не такой же жёсткий провал, как невыдача Pтреб."
    InstrField ws, r, "Режим «Авария»", "Ступени справочно", "Несданные t5/t10 в «Аварии» не делают общий вердикт НАРУШЕНИЕМ сами по себе: они попадают в замечание. НАРУШЕНИЕМ по качеству здесь становится отсутствие адекватной реакции у оборудования с регулятором мощности (п. 4.4.4)."
    InstrField ws, r, "Режим «Испытание»", "Ступени в вердикте", "Для контрольных испытаний несданные t5 или t10 — это уже НАРУШЕНИЕ. Берутся фактические S и fнч из Config, подстановка 6 % и 0,15 Гц не делается."
    InstrField ws, r, "Summary", "Кач. статус, t5, t10, уст", "Три подпункта и общий качественный статус. Рядом — фактические секунды t5 и t10 и тип реакции. Н/П по п. 4.1.3 значит, что событие для качественной оценки неприменимо."
    r = r + 1
    InstrHeading ws, r, "5. Вместе, только количество или только качество"
    InstrField ws, r, "Вместе", "Обычный запуск", "Оставьте I=1 и J=1. Кнопка «Запустить мониторинг ОПРЧ» за один проход считает количество и качество, графики и протокол. Так и задумано для рабочего анализа."
    InstrField ws, r, "Только количество", "Config J = 0", "Качественные ступени не проверяются и не портят вердикт. Количественная оценка, реакция 15 с и дефицит остаются."
    InstrField ws, r, "Только качество", "Отдельной кнопки нет", "Количественная оценка для включённого генератора считается всегда. Если в режиме «Авария» событие не проходит п. 4.1.2, количественный статус будет «Н/П», а в колонке «Вывод» останется качественная часть. Это ближайший к «только качеству» результат."
    InstrField ws, r, "Config X13", "Авария или Испытание", "«Авария» — разбор реального отклонения в сети. «Испытание» — контрольное испытание, где ступени t5/t10 входят в вердикт. Режим один на весь запуск, он лежит в ячейке X13."
    r = r + 1
    InstrHeading ws, r, "6. Куда смотреть результат"
    InstrNote ws, r, "Главная колонка — Summary «Вывод»: текст начинается с ОК, НАРУШЕНИЕ или НЕ ОЦЕНИВАЕТСЯ. Для неблочной станции нормативный вывод — строка СУММА. Листы с окончанием «_Граф» — графики. Лист «Протокол» — тот же вывод связным текстом. Лист Log — предупреждения о данных, а не вердикт."
    InstrNote ws, r, "Кнопка «Экспорт отчёта в PDF» не печатает широкую сводку Summary как есть. Она собирает альбомный отчёт: короткая таблица с переносом длинного «Вывода», протокол с переносами строк и каждый график на своём листе. Файл PDF появляется в той же папке, что и книга."
    r = r + 1
    InstrHeading ws, r, "7. Запуск"
    InstrNote ws, r, "Разрешите макросы. Лист Config — кнопка «Запустить мониторинг ОПРЧ» (или Alt+F8, макрос AnalyzeOPRCH). Если кнопок нет: Alt+F8 — InitOPRCHUI. Кнопка «Применить пресеты типов» заполняет только пустые нормативы. SetupOPRCHTemplate на уже заполненной книге не запускайте: он перезапишет RawData и Config."
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 78
    On Error Resume Next
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    On Error GoTo 0
End Sub

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
    OPRCHMsg "Ошибка InitOPRCHUI: " & Err.Description, vbCritical
End Sub

Public Sub InitOPRCHUISilent()
    Dim previousQuiet As Boolean, errorInfo As String
    previousQuiet = gOPRCHQuiet
    gOPRCHQuiet = True
    On Error GoTo InitFail
    WriteInstructionSheet
    EnsureControlButtons EnsureSheet(SH_CFG)
    EnsureControlButtons EnsureSheet(SH_INSTR)
    gOPRCHQuiet = previousQuiet
    Exit Sub
InitFail:
    errorInfo = "Ошибка инициализации UI #" & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    AppendLog "ERROR", "InitOPRCHUISilent", errorInfo
    gOPRCHQuiet = previousQuiet
End Sub

' Единая точка диалогов: в тихом режиме ничего не показывает.

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
    Set shp = ws.Shapes.AddShape(5, anchor.Left, anchor.Top, w, h)  ' 5 = msoShapeRoundedRectangle
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
