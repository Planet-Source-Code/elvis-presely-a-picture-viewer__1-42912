VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1305
   ClientLeft      =   2610
   ClientTop       =   2130
   ClientWidth     =   1725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Password.frx":0442
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   1725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter administrator password"
      Height          =   675
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    frmPassword.Width = txtPassword.Width + 300

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Dim msg As Integer
    
    On Error Resume Next
    
    If KeyAscii = 13 Then
        If txtPassword.Text = "346611" Then
        frmpicture.File1.Archive = True
        frmpicture.File1.Hidden = True
        frmpicture.File1.System = True
        frmpicture.MnuHidden.Checked = True
        frmpicture.File1.Refresh
        frmpicture.picture1.Refresh
        Unload Me
        Else
        msg = MsgBox("Invalid password please try again", vbCritical + vbOKOnly, "Warning")
        txtPassword.Text = ""
        frmpicture.MnuHidden.Checked = False
        frmpicture.File1.Archive = False
        frmpicture.File1.Hidden = False
        frmpicture.File1.System = False
        End If
    End If
    
    If KeyAscii = 27 Then
    frmpicture.MnuHidden.Checked = False
    frmpicture.File1.Archive = False
    frmpicture.File1.Hidden = False
    frmpicture.File1.System = False
    frmpicture.File1.Refresh
    frmpicture.picture1.Refresh
    Unload Me
    End If
End Sub
