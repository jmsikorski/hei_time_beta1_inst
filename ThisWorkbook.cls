VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If Environ$("username") = "jsikorski" Then
        ExportVBA
    End If

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If Environ$("username") = "jsikorski" Then
        ExportVBA
    End If
    On Error GoTo 0
    If Environ$("username") <> "jsikorski" Then
        MsgBox "ERROR: File is read only", vbCritical + vbOKOnly, "READ ONLY!"
        Cancel = True
    Else
        If ThisWorkbook.ReadOnly Then
            unlock_workbook
        End If
    End If
End Sub

Private Sub Workbook_Open()
    Application.WindowState = xlMaximized
    main
End Sub

Private Sub unlock_workbook()
    Application.DisplayAlerts = False
    On Error GoTo kill_old
    ThisWorkbook.SaveAs ThisWorkbook.path & "\2" & ThisWorkbook.name
    On Error GoTo 0
    SetAttr ThisWorkbook.path & "\" & Right(ThisWorkbook.name, Len(ThisWorkbook.name) - 1), vbNormal
    ThisWorkbook.SaveAs Right(ThisWorkbook.name, Len(ThisWorkbook.name) - 1)
    Kill ThisWorkbook.path & "\2" & ThisWorkbook.name
    Application.DisplayAlerts = True
    Exit Sub
kill_old:
    Kill ThisWorkbook.path & "\2" & ThisWorkbook.name
    Resume Next
End Sub
