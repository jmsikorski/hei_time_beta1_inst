Attribute VB_Name = "ExportVisualBasicCode"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model

Public Sub ExportVBA()

Const Module = 1
Const ClassModule = 2
Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim dir_main As String
    Dim extension As String
    Dim FSO As New FileSystemObject
    Dim dirs() As String
    Dim directory As Variant
    ReDim dirs(2)
    dirs(0) = "C:\Users\jsikorski\Desktop\Time Card Project - JASON\ALL VBA CODE\" & ThisWorkbook.name & "_VBA_" & Format(Now(), "mm.dd.yy_hh.mm.ss")
    dirs(1) = "C:\Users\jsikorski\Desktop\Time Card Project - JASON\hei_time\" & ThisWorkbook.name & "_VBA"
    dirs(2) = "C:\Users\jsikorski\Desktop\Time Card Project\Installer"
    count = 0

    For Each directory In dirs
        If Not FSO.FolderExists(directory) Then
            Call FSO.CreateFolder(directory)
        Else
            clearFolder (directory)
        End If
    Next
    Set FSO = Nothing

    For Each directory In dirs
        For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
            Select Case VBComponent.Type
                Case ClassModule, Document
                    extension = ".cls"
                Case Form
                    extension = ".frm"
                Case Module
                    extension = ".bas"
                Case Else
                    extension = ".txt"
            End Select
    
    
            On Error Resume Next
            Err.Clear
    
            path = directory & "\" & VBComponent.name & extension
            Call VBComponent.Export(path)
    
            If Err.Number <> 0 Then
                Call MsgBox("Failed to export " & VBComponent.name & " to " & path, vbCritical)
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.name & ":" & space(Padding), Padding) & path
            End If
    
            On Error GoTo 0
        Next
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & dir_main
    Application.StatusBar = False
End Sub

Public Sub clearFolder(xFolder As String)
    Dim FSO As New FileSystemObject
    Dim xFile As file
    If Not FSO.FolderExists(xFolder) Then
        Call FSO.CreateFolder(xFolder)
        Exit Sub
    End If
    Set FSO = Nothing
    
    For Each xFile In FSO.GetFolder(xFolder).Files
        On Error GoTo close_file
        Kill xFile
    Next
    Exit Sub
close_file:
    Err.Clear
    Dim ans As Integer
    ans = MsgBox("ERROR! Unable to remove files", vbAbortRetryIgnore + vbCritical, "ERROR!")
    If ans = vbRetry Then
        clearFolder (xFolder)
    ElseIf ans = vbAbort Then
        main
    Else
        ThisWorkbook.Close False
    End If
End Sub

Public Sub importDataFile()
    Dim FSO As FileSystemObject
    Dim objShell As WshShell
    Set FSO = New FileSystemObject
    Dim xFolder As String
    Dim xFile As Object
    Set objShell = New WshShell
    xFolder = objShell.SpecialFolders("Desktop") & "\Time Card Project\"
    
    For Each xFile In FSO.GetFolder(xFolder).Files
        If FSO.GetExtensionName(xFile.name) = "xlsx" Or FSO.GetExtensionName(xFile.name) = "xlsm" Then
            FSO.CopyFile xFile, ThisWorkbook.Worksheets(1).Range("aPath") & "/" & xFile.name
        End If
    Next
End Sub
Public Sub rebuildFile(rFile As Integer)
Attribute rebuildFile.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim xlFile As String
    Dim rng As Range
    Dim cFolder As String
    Dim templateName As String
    templateName = " TEMPLATE.xlsm"
    Set objShell = New WshShell
    cFolder = objShell.SpecialFolders("Desktop")
    
    Select Case rFile
        Case 1 ' Rebuild Master File
            xlFile = ThisWorkbook.Worksheets(1).Range("aFile")
            xlFile = Left(xlFile, Len(xlFile) - 5) & templateName
            cFolder = cFolder & "\Time Card Project\DEMO"
        Case 2 ' Rebuild Builder File
'            xlFile =
'            cFolder = cFolder & "\Time Card Project\Builder"
        Case 3 ' Rebuild Installer File
'            xlFile =
'            cFolder = cFolder & "\Time Card Project\Installer"
        Case Else
            Debug.Print "ERROR REBUILDING FILE"
            Exit Sub
    End Select
    Application.EnableEvents = False
    Workbooks.Open ThisWorkbook.Worksheets(1).Range("sp_Path") & xlFile
    Application.EnableEvents = True
    Workbooks(xlFile).Activate
    ImportModules cFolder
    For Each rng In ThisWorkbook.Worksheets("BUILD").Range("A1", ThisWorkbook.Worksheets("BUILD").Range("A1").End(xlDown))
        AddReference rng.Value, rng.Offset(0, 1).Value
    Next
    Application.DisplayAlerts = False
    Dim newFile As String
    newFile = ThisWorkbook.Worksheets(1).Range("aPath") & "\" & ThisWorkbook.Worksheets(1).Range("aFile")
    ActiveWorkbook.SaveAs newFile
    Application.DisplayAlerts = True

End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select code Folder"
        .AllowMultiSelect = False
        .InitialFileName = FolderWithVBAProjectFiles
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub Unzip1()
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip", MultiSelect:=False)
    If Fname = False Then
        'Do nothing
    Else
        'Root folder for the new folder.
        'You can also use DefPath = "C:\Users\Ron\test\"
        'DefPath = Application.DefaultFilePath
        DefPath = "C:"
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        'Create the folder name
        strDate = Format(Now, " dd-mm-yy h-mm-ss")
        FileNameFolder = DefPath & "MyUnzipFolder " & strDate & "\"

        'Make the normal folder in DefPath
        MkDir FileNameFolder

        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")

        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items

        'If you want to extract only one file you can use this:
        'oApp.Namespace(FileNameFolder).CopyHere _
         'oApp.Namespace(Fname).items.Item("test.txt")

        MsgBox "You find the files here: " & FileNameFolder

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Sub

Public Sub ImportModules(Optional codeFolder As String)
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    Dim objShell As WshShell
    
    If codeFolder = "" Then
        Set objShell = New WshShell
        codeFolder = objShell.SpecialFolders("Desktop")
        codeFolder = codeFolder & "\Time Card Project\"
    End If
    
    If ActiveWorkbook.name = ThisWorkbook.name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles(codeFolder) = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.name
    
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
'    If wkbTarget.VBProject.Protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = codeFolder & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms
    Dim cnt As Integer
    cnt = 0
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.name) = "cls") Then
            Debug.Print objFile.name
            If Left(objFile.name, 12) = "ThisWorkbook" Then
                With ActiveWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
                    .DeleteLines StartLine:=1, count:=.CountOfLines
                    .AddFromFile objFile.path
                    .DeleteLines StartLine:=1, count:=4
                End With
            ElseIf Left(objFile.name, 5) = "Sheet" Then
                On Error Resume Next
                With ActiveWorkbook.VBProject.VBComponents(Left(objFile.name, Len(objFile.name) - 4)).CodeModule
                    .DeleteLines StartLine:=1, count:=.CountOfLines
                    .AddFromFile objFile.path
                    .DeleteLines StartLine:=1, count:=4
                End With
            Else
                cmpComponents.Import objFile.path
                On Error GoTo 0
            End If
        
        ElseIf (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.path
            Debug.Print objFile.name
        End If
        cnt = cnt + 1
    Next objFile
    
    Debug.Print "Imported " & cnt & " Files"
End Sub

Function FolderWithVBAProjectFiles(Optional xFolder As String) As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String
    
    If xFolder = "" Then
        Set WshShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("scripting.filesystemobject")
    
        SpecialPath = WshShell.SpecialFolders("DEsktop")
    
        If Right(SpecialPath, 1) <> "\" Then
            SpecialPath = SpecialPath & "\"
        End If
        xFolder = SpecialPath & "Time Card Project - JASON\hei_time\Time Card Gen BETA.xlsm_VBA"
    End If
    
    Set FSO = New FileSystemObject
    If FSO.FolderExists(Left(xFolder, Len(xFolder) - 1)) = False Then
        On Error Resume Next
        MkDir xFolder
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(xFolder) = True Then
        FolderWithVBAProjectFiles = xFolder
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim vbProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set vbProj = ActiveWorkbook.VBProject
        For Each VBComp In vbProj.VBComponents
            If VBComp.name <> "main_module" Then
                If VBComp.Type = vbext_ct_Document Then
                    'Thisworkbook or worksheet module
                    'We do nothing
                Else
                    vbProj.VBComponents.Remove VBComp
                End If
            End If
        Next VBComp
End Function

Sub AddReference(rName As String, rLoc As String)
    Debug.Print rName
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    '~~> Check if "Microsoft VBScript Regular Expressions 5.5" is already added
    For Each chkRef In vbProj.References
        If chkRef.Description = rName Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile rLoc

CleanUp:
    If BoolExists = True Then
        Debug.Print rName & " already exists"
    Else
        Debug.Print rName & " Added Successfully"
    End If

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub
