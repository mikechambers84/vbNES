VERSION 5.00
Begin VB.Form frmConfigInput 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmConfigInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblWhichKey 
      Alignment       =   2  'Center
      Caption         =   "Press key for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmConfigInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public curkey As Long

Private Sub changedispkey()
    Select Case curkey
        Case 0, 8: lblWhichKey = "Press key for button A"
        Case 1, 9: lblWhichKey = "Press key for button B"
        Case 2, 10: lblWhichKey = "Press key for select"
        Case 3, 11: lblWhichKey = "Press key for start"
        Case 4, 12: lblWhichKey = "Press key for up"
        Case 5, 13: lblWhichKey = "Press key for down"
        Case 6, 14: lblWhichKey = "Press key for left"
        Case 7, 15: lblWhichKey = "Press key for right"
    End Select
End Sub

Public Sub startkeyconfig1()
    curkey = 0
    changedispkey
    Me.Caption = "Configure controller 1"
End Sub

Public Sub startkeyconfig2()
    curkey = 8
    changedispkey
    Me.Caption = "Configure controller 2"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    keymap(curkey) = KeyCode
    curkey = curkey + 1
    changedispkey
    If (curkey = 8) Or (curkey = 16) Then
        frmMain.saveconfig
        Unload Me
        frmMain.Show
    End If
End Sub

