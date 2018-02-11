VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainMenu 
   Caption         =   "Helix Time Card Generator"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "mainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const p1 = 6
Private Const p2 = 144
Private Const p3 = 278
Public ans As Integer

Private Sub pExit_Click()
    Me.ans = -1
    Me.Hide
End Sub

Private Sub pInstall_Click()
    Me.ans = 1
    Me.Hide
End Sub

Private Sub prun_Click()
    ans = 3
    Me.Hide
End Sub

Private Sub pUninstall_Click()
    Me.ans = 2
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    
    If ThisWorkbook.Worksheets(1).Range("appinstalled") = False Then
        Me.pInstall.Left = p1
        Me.pUninstall.Left = p2
        Me.pExit.Left = p3
        Me.pRun.Visible = False
    Else
        Me.pRun.Left = p1
        Me.pUninstall.Left = p2
        Me.pExit.Left = p3
        Me.pInstall.Visible = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       ans = -1
    End If
End Sub

