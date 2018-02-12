Attribute VB_Name = "main_program"
Private Const pw = "hei3078"
Private Const dt = "DATA"
Private Const exeName = "TimeCardGen.xlsm"
Public Enum mAns
    go = 3
    uninstall = 2
    install = 1
End Enum
Public Enum fileType
    master = 1
    builder = 2
    Installer = 3
End Enum
Public Sub showbook()
    ThisWorkbook.Worksheets("DATA").Visible = True
    ThisWorkbook.Worksheets("BUILD").Visible = True

End Sub
Public Sub main()
    For i = 1 To ThisWorkbook.Sheets.count - 1
        ThisWorkbook.Worksheets(i).Visible = xlVeryHidden
    Next
    Dim mMenu As mainMenu
    Dim ans As Integer
    Dim retryAns As Integer
    Dim reinstall As Boolean
    Dim testPath As String
    testPath = Dir(ThisWorkbook.Worksheets(1).Range("aPath"), vbDirectory)
    If testPath = "" Then
        ThisWorkbook.Worksheets(1).Range("appinstalled") = False
    Else
        ThisWorkbook.Worksheets(1).Range("appinstalled") = True
    End If
    reinstall = False
    ans = 0
    Set mMenu = New mainMenu
retry_line:
    mMenu.Show
    If mMenu.ans = mAns.go Then
        main_run
    ElseIf mMenu.ans = mAns.install Then
        ans = main_install
        If ans = 1 Then
            MsgBox "Installed " & ThisWorkbook.Worksheets(dt).Range("aFile") & vbNewLine & "File is located in /Documents/Time Card Generator", vbOKOnly + vbInformation, "SUCCESS!"
            Set mMenu = Nothing
            If Application.Workbooks.count = 1 Then
                Stop
                Application.DisplayAlerts = False
                ThisWorkbook.Saved = True
                Application.Quit
            Else
                ThisWorkbook.Close False
            End If
        ElseIf ans = 2 Then
            If Environ$("username") = "jsikorski" Then
                reinstall = True
                GoTo uninstall_line
            End If
            retryAns = MsgBox("File already installed, would you like to repair installation?", _
                vbYesNoCancel + vbExclamation, "ERROR!")
            If retryAns = vbCancel Then
                GoTo retry_line
            ElseIf retryAns = vbNo Then
                Exit Sub
            ElseIf retryAns = vbYes Then
                reinstall = True
                GoTo uninstall_line
            End If
        Else
            GoTo clear_bad_install
            MsgBox "Unable to install!", vbCritical + vbOKOnly, "ERROR"
        End If
    ElseIf mMenu.ans = mAns.uninstall Then
uninstall_line:
        ans = main_uninstall(reinstall)
        If Not reinstall Then
            If ans = 1 Then
                MsgBox "Uninstall Complete!", vbOKOnly + vbInformation, "SUCCESS!"
            ElseIf ans = 2 Then
                retryAns = MsgBox("File already uninstalled, would you like to reinstall?", _
                    vbYesNoCancel + vbExclamation, "ERROR!")
                If retryAns = vbCancel Then
                    GoTo retry_line
                ElseIf retryAns = vbNo Then
                    Exit Sub
                ElseIf retryAns = vbYes Then
                    main_install
                End If
            Else
                MsgBox "Unable to uninstall!" & vbNewLine & vbNewLine & "Please close all files and try again", vbCritical + vbOKOnly, "ERROR"
            End If
        End If
    Else
        If Environ$("username") <> "jsikorski" Then
            'ThisWorkbook.Close , False
        End If
    End If
    GoTo clean_up:
clear_bad_install:
    If clearFolder(ws.Range("aPath")) <> 1 Then
        MsgBox "Error cleaning up improper installation", vbOKOnly, "ERROR"
    End If
    RmDir ws.Range("aPath")
clean_up:
    Set mMenu = Nothing
End Sub

Public Function main_uninstall(Optional reinstall As Boolean) As Integer
    'On Error GoTo uninstall_err
    ans = MsgBox("Are you sure you want to uninstall " & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "?", vbExclamation + vbOKCancel, "CONFIRM UNINSTALL")
    If ans <> vbOK Then
        Exit Function
    End If

    Dim testPath As String
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Set ws = ThisWorkbook.Worksheets(dt)
    testPath = Dir(ws.Range("aPath"), vbDirectory)
    Dim shl As WshShell
    Dim iPath As String
    Set shl = New WshShell
    iPath = shl.SpecialFolders(16) ' "Documents" Folder
    If Right(iPath, 1) <> "\" Then
        iPath = iPath & "\"
    End If
    
    If testPath = "" Then
        main_uninstall = 2
        Exit Function
    End If
    If clearFolder(ws.Range("aPath")) <> 1 Then
        main_uninstall = -1
        GoTo clean_up
    End If
    FSO.DeleteFolder ws.Range("aPath")
    If FSO.FolderExists(iPath & "Time Card Generator") Then
        On Error GoTo uninstall_err
        FSO.DeleteFolder iPath & "Time Card Generator"
        On Error GoTo 0
    End If

    ThisWorkbook.Worksheets(1).Range("reg_user") = vbNullString
    ThisWorkbook.Worksheets(1).Range("reg_password") = vbNullString
    If reinstall Then
        If main_install <> 1 Then
            MsgBox "Unable to install!", vbCritical + vbOKOnly, "ERROR"
        Else
            MsgBox "Installation Complete!", vbOKOnly + vbInformation, "SUCCESS!"
        End If
    End If
    GoTo clean_up
uninstall_err:
    If Environ$("username") = "jsikorski" Then
        Debug.Print "**UNINSTALL ERROR DETAILS**"
        Debug.Print "Description: " & Err.Description
        Debug.Print "Help Context: " & Err.HelpContext
        Debug.Print "Help File: " & Err.HelpFile
        Debug.Print "Last DLL Error: " & Err.LastDllError
        Debug.Print "Number: " & Err.Number
        Debug.Print "Source: " & Err.Source
        'Err.Raise Err.Number
    End If
    main_uninstall = -1
    If Err.Number <> 70 Then
        Err.Clear
        Exit Function
    End If
clean_up:
    Set ws = Nothing
    Set shl = Nothing
    If FSO.FileExists(iPath & "Time Card Generator/" & ThisWorkbook.Name) Then
        Set FSO = Nothing
        killThisFile
    End If
    Set FSO = Nothing
    main_uninstall = 1
End Function

Private Sub killThisFile()
' Original code from Tom Urtis
    Dim ans As Integer
    With ThisWorkbook
        
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        Application.Quit

    End With
End Sub
 
 Public Function getXPass() As String
    Dim xPass As String
    Dim tString As String
    tString = Environ$("username")
    For i = 1 To Len(tString)
        xPass = xPass & Chr(Asc(Left(tString, 1)) Xor (Len(tString) + 1) * 4)
        tString = Right(tString, Len(tString) - 1)
    Next
    getXPass = xPass
End Function

Public Function main_install() As Integer
    'On Error GoTo install_err
    Dim ws As Worksheet
    Dim testPath As String
    Set ws = ThisWorkbook.Worksheets(dt)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Dim shl As WshShell
    Dim iPath As String
    Set shl = New WshShell
    iPath = shl.SpecialFolders(16) ' "Documents" Folder
    If Right(iPath, 1) <> "\" Then
        iPath = iPath & "\"
    End If
    testPath = Dir(ws.Range("aPath"), vbDirectory)
    If testPath <> "" Then
        main_install = 2
        Exit Function
    End If
    Set lMenu = New loginMenu
    Dim frm As Object
    Do While ws.Range("reg_user") = vbNullString _
        Or ws.Range("reg_password") = vbNullString
        Dim cnt As Integer
        cnt = 0
        For Each frm In VBA.UserForms
            cnt = cnt + 1
        Next
        If cnt = 1 Then
            Exit Do
        Else
            lMenu.Show
        End If
    Loop
    MkDir ws.Range("aPath")
    makeLnkPath ws.Range("sp_path")
    rebuildFile (fileType.master)
    ExportVisualBasicCode.importDataFile
    Workbooks(ws.Range("aFile").Value).Worksheets("HOME").Range("reg_user").Value = ws.Range("reg_user").Value
    Workbooks(ws.Range("aFile").Value).Worksheets("HOME").Range("reg_pass").Value = ws.Range("reg_password").Value
    ws.Range("reg_user").Clear
    ws.Range("reg_password").Clear
    ws.Range("appinstalled").Value = True
    Workbooks(ws.Range("aFile").Value).Protect getXPass
    Workbooks(ThisWorkbook.Worksheets(dt).Range("aFile").Value).Save
    Workbooks(ThisWorkbook.Worksheets(dt).Range("aFile").Value).Close
    If FSO.FolderExists(iPath & "Time Card Generator") Then
        ThisWorkbook.SaveAs iPath & "Time Card Generator\" & exeName
    Else
        FSO.CreateFolder iPath & "Time Card Generator"
        ThisWorkbook.SaveAs iPath & "Time Card Generator\" & exeName
    End If
    main_install = 1
    Exit Function
install_err:
    If Environ$("username") = "jsikorski" Then
        Debug.Print "**INSTALL ERROR DETAILS**"
        Debug.Print "Description: " & Err.Description
        Debug.Print "Help Context: " & Err.HelpContext
        Debug.Print "Help File: " & Err.HelpFile
        Debug.Print "Last DLL Error: " & Err.LastDllError
        Debug.Print "Number: " & Err.Number
        Debug.Print "Source: " & Err.Source
        Err.Raise Err.Number
    End If
    Err.Clear
    main_install = -1
End Function

Public Function makeLnkPath(ByVal lnk As String) As Integer
   On Error Resume Next
    Dim WshShell As Object
    Dim oURLlink As Object
    Set WshShell = CreateObject("WScript.Shell")
    Set oURLlink = WshShell.CreateShortcut(ThisWorkbook.Worksheets(dt).Range("aPath") & "\Data.URL")
    With oURLlink
        .TargetPath = lnk
        .Save
        makeLnkPath = 1
        Exit Function
   End With
   makeLnkPath = -1
End Function

Public Function Getlnkpath(ByVal lnk As String) As String
   On Error Resume Next
   With CreateObject("Wscript.Shell").CreateShortcut(lnk)
       Getlnkpath = .TargetPath
       .Close
   End With
End Function

Public Sub t12()
    Dim xlFile As String
    xlFile = Getlnkpath("C:\ProgramData\HelixTimeCard\Data.URL") & "Lead Card.xlsx"
    Debug.Print xlFile
    Workbooks.Open xlFile
End Sub

Private Function check_URL(url As String) As Boolean
    check_URL = False
    Dim WinHttp As New WinHttpRequest
    WinHttp.Open "get", url, False
    WinHttp.Send
    If WinHttp.Status = 200 Then check_URL = True
End Function

Public Sub main_run()
    Dim xlFile As String, xlPath As String
    Dim ans As Integer
    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(dt)
    ws.Unprotect pw
    i = 0
    Stop
    Application.WindowState = xlMaximized
'    If ws.Range("apprunning") = False Then
'        ans = MsgBox("Quit?", vbQuestion + vbYesNoCancel, "Helix Time Card Gen")
'        If ans = vbYes Then
'            ThisWorkbook.Close False
'        ElseIf ans = vbNo Then
'            ws.Range("apprunning") = True
'            main
'        ElseIf vbCancel Then
'            MsgBox "This file is not for editing, Please close", vbApplicationModal + vbOKOnly + vbExclamation, "PLEASE CLOSE"
'            On Error Resume Next
'            ws.Visible = False
'            ws.Protect pw
'            On Error GoTo 0
'            SetAttr ThisWorkbook.path & "\" & ThisWorkbook.name, vbReadOnly
'            Exit Sub
'        End If
'    End If
    On Error GoTo data_sht_not_found
    If Worksheets(1).Name <> dt Then
        ws.Move before:=ThisWorkbook.Worksheets(1)
    End If
    On Error GoTo 0
    
    xlPath = ws.Range("aPath")
    If xlPath = vbNullString Then
        xlPath = "C:\ProgramData\HelixTimeCard"
    End If
    xlFile = ws.Range("aFile")
    On Error Resume Next
    Set wb = Workbooks(xlFile)
    ws.Range("appRunning") = True
    On Error GoTo file_not_open
    If wb.Name = xlFile Then
        Application.Run "'" & xlFile & "'" & "!Timecard.main"
        On Error GoTo 0
        Exit Sub
    End If
    Exit Sub
file_not_open:
    Err.Clear
    On Error GoTo file_not_found
    Application.Workbooks.Open xlPath & "\" & xlFile
    Exit Sub
file_not_found:
    openCheck xlPath, xlFile
    Resume Next
data_sht_not_found:
    ans = MsgBox("ERROR: Could not find Data sheet!", vbAbortRetryIgnore + vbCritical, "ERROR!")
    If ans = vbAbort Then
        ThisWorkbook.Close False
    ElseIf ans = vbRetry Then
        main
    ElseIf ans = vbIgnore Then
        MsgBox "To correct this error please repair " & dt & " worksheet!", vbInformation + vbOKOnly, "EXITING TO WORKBOOK"
    End If
    On Error GoTo 0
End Sub
