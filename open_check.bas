Attribute VB_Name = "open_check"
Option Explicit
Option Private Module

Public Function openCheck(ByRef xlFile As Variant, ByRef xlPath As String) As Integer
'This function attempts to open an Excel file
'If the file is not found it asks user to locate file
'Return values:
'2: Success, xlFile changed
'1: Success,
'-1: Error opening unresolved
'-2: Error opening no file selected


'NOTE: current application has Application as listed Excel application
    On Error GoTo retry
    Dim ans As Integer
    Dim f As Variant
    Dim strFileToOpen As Variant
    Application.Workbooks.Open Filename:=xlFile
    openCheck = 1
    On Error GoTo 0
    ans = 0
'    Application.GetOpenFilename
retry:
    ChDir ThisWorkbook.path
    strFileToOpen = Application.GetOpenFilename _
    (Title:=xlFile & " Not found", _
    filefilter:="Excel Files *.xls* (*.xls*),", _
    MultiSelect:=False)
    If strFileToOpen = False Then
        ans = MsgBox("No file selected.", vbExclamation + vbAbortRetryIgnore, "Sorry!")
        If ans = vbRetry Then
            GoTo retry
        ElseIf ans = vbAbort Or ans = vbCancel Then
            openCheck = -2
        ElseIf ans = vbIgnore Then
            ThisWorkbook.Close False
        Else
            openCheck = -1
        End If
    Else
        On Error GoTo load_err
        Dim wb As Workbook
        Set wb = Application.Workbooks.Open(Filename:=strFileToOpen)
        If wb.path <> xlPath Then
            openCheck = 2
        Else
            openCheck = 1
        End If
    End If
    On Error GoTo 0
    Exit Function
load_err:
    ans = MsgBox("Unable to open selected file!", vbCritical + vbAbortRetryIgnore, "ERROR!")
    If ans = vbAbort Or ans = vbCancel Then
        openCheck = -2
    ElseIf ans = vbRetry Then
        GoTo retry
    ElseIf ans = vbIgnore Then
        ans = -1
    End If
    On Error GoTo 0
End Function



