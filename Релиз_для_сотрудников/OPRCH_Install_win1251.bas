Attribute VB_Name = "OPRCH_Install"
Option Explicit

' Установщик OPRCH v1.8.2 (универсальный шаблон: RawData + Config)
' Импортируйте ТОЛЬКО этот файл, затем Alt+F8 -> InstallOPRCHMacros.
' Нужна галочка: Trust access to the VBA project object model.
' Предпочтительнее открыть готовый Шаблон_мониторинг_ОПРЧ.xlsm без установки.

Public Sub InstallOPRCHMacros()
    Dim folder As String, i As Long
    Dim files As Variant, f As String
    Dim comp As Object, j As Long
    Dim dest As String, wb As Workbook

    On Error GoTo EH
    Set wb = ThisWorkbook
    folder = wb.Path
    If Len(folder) = 0 Then
        MsgBox "Сначала сохраните книгу на диск (лучше сразу как .xlsm) в папку с файлами *_win1251.bas.", vbExclamation
        Exit Sub
    End If
    folder = folder & Application.PathSeparator

    files = Array( _
        "OPRCH_Common_win1251.bas", _
        "OPRCH_Engine_win1251.bas", _
        "OPRCH_Station_win1251.bas", _
        "OPRCH_Sheets_win1251.bas", _
        "OPRCH_Charts_win1251.bas", _
        "OPRCH_Monitor_win1251.bas" _
    )

    For i = LBound(files) To UBound(files)
        f = folder & CStr(files(i))
        If Dir(f) = "" Then
            MsgBox "Не найден файл:" & vbCrLf & f & vbCrLf & vbCrLf & _
                   "Положите книгу в ту же папку, где лежат шесть *_win1251.bas.", vbCritical
            Exit Sub
        End If
    Next i

    For j = wb.VBProject.VBComponents.Count To 1 Step -1
        Set comp = wb.VBProject.VBComponents(j)
        If Left$(comp.Name, 6) = "OPRCH_" And comp.Name <> "OPRCH_Install" Then
            wb.VBProject.VBComponents.Remove comp
        End If
    Next j

    For i = LBound(files) To UBound(files)
        wb.VBProject.VBComponents.Import folder & CStr(files(i))
    Next i

    On Error Resume Next
    Application.Run "WriteInstructionSheet"
    On Error GoTo EH

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" Then
        dest = folder & Replace(wb.Name, ".xlsx", "") & ".xlsm"
        If LCase$(Right$(dest, 5)) <> ".xlsm" Then dest = dest & ".xlsm"
        wb.SaveAs Filename:=dest, FileFormat:=52
    Else
        wb.Save
    End If

    MsgBox "Макросы OPRCH v1.8.2 установлены." & vbCrLf & _
           "Config -> «Запустить мониторинг ОПРЧ»." & vbCrLf & _
           "Модуль OPRCH_Install можно удалить (Alt+F11).", vbInformation
    Exit Sub
EH:
    MsgBox "Ошибка установки: " & Err.Description & vbCrLf & vbCrLf & _
           "Проверьте: Параметры макросов -> «Доверить доступ к объектной модели проектов VBA».", vbCritical
End Sub
