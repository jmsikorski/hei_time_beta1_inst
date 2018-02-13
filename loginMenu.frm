VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "loginMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "loginMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub loginButton_Click()
    If userBox = vbNullString And pwBox = vbNullString Then
        MsgBox "Username & Password are required!", vbOKOnly + vbCritical, "ERROR!"
        Exit Sub
    ElseIf userBox = vbNullString Then
        MsgBox "Please enter a Username", vbOKOnly + vbCritical, "ERROR!"
        Exit Sub
    ElseIf pwBox = vbNullString Then
        MsgBox "Please enter a Password", vbOKOnly + vbCritical, "ERROR!"
        Exit Sub
    End If
    ThisWorkbook.Worksheets(1).Range("reg_password") = Me.pwBox.Value
    ThisWorkbook.Worksheets(1).Range("reg_user") = Me.userBox.Value
    Unload Me
End Sub

Private Sub pw_reset_Click()
    ThisWorkbook.Worksheets(1).Range("reg_password") = vbNullString
    ThisWorkbook.Worksheets(1).Range("reg_user") = vbNullString
    Unload Me
End Sub

Private Sub pwBox_Change()
    
End Sub

Private Sub userBox_Change()

End Sub

Private Sub UserForm_Initialize()
    With Me
        .userBox.Value = Environ$("username")
        .pwBox.SetFocus
        .Caption = "LOGIN"
    End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        pw_reset_Click
    End If
End Sub
