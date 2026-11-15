Attribute VB_Name = "OPRCH_Install"
Option Explicit

' Установщик OPRCH v1.8.6 (Windows fixes) (универсальный шаблон: RawData + Config)
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
        "OPRCH_Common", _
        "OPRCH_Engine", _
        "OPRCH_Station", _
        "OPRCH_Sheets", _
        "OPRCH_Charts", _
        "OPRCH_Monitor" _
    )

    Dim resolved() As String
    ReDim resolved(LBound(files) To UBound(files))
    For i = LBound(files) To UBound(files)
        f = ResolveOPRCHBas(folder, CStr(files(i)))
        If Len(f) = 0 Then
            MsgBox "Не найден модуль " & CStr(files(i)) & vbCrLf & vbCrLf & _
                   "Положите шесть файлов OPRCH_*.bas рядом с книгой" & vbCrLf & _
                   "или в подпапку макросы_для_импорта.", vbCritical
            Exit Sub
        End If
        resolved(i) = f
    Next i

    For j = wb.VBProject.VBComponents.Count To 1 Step -1
        Set comp = wb.VBProject.VBComponents(j)
        If Left$(comp.Name, 6) = "OPRCH_" And comp.Name <> "OPRCH_Install" Then
            wb.VBProject.VBComponents.Remove comp
        End If
    Next j

    For i = LBound(files) To UBound(files)
        wb.VBProject.VBComponents.Import resolved(i)
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

    MsgBox "Макросы OPRCH v1.8.6 установлены." & vbCrLf & _
           "Config -> «Запустить мониторинг ОПРЧ»." & vbCrLf & _
           "Модуль OPRCH_Install можно удалить (Alt+F11).", vbInformation
    Exit Sub
EH:
    MsgBox "Ошибка установки: " & Err.Description & vbCrLf & vbCrLf & _
           "Проверьте: Параметры макросов -> «Доверить доступ к объектной модели проектов VBA».", vbCritical
End Sub

Private Function ResolveOPRCHBas(ByVal folder As String, ByVal baseName As String) As String
    Dim dirs As Variant, names As Variant
    Dim d As Variant, n As Variant, f As String
    dirs = Array( _
        folder, _
        folder & "макросы_для_импорта" & Application.PathSeparator _
    )
    names = Array(baseName & "_win1251.bas", baseName & ".bas")
    For Each d In dirs
        For Each n In names
            f = CStr(d) & CStr(n)
            If Len(Dir(f)) > 0 Then
                ResolveOPRCHBas = f
                Exit Function
            End If
        Next n
    Next d
    ResolveOPRCHBas = ""
End Function
