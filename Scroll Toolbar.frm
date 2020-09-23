VERSION 5.00
Begin VB.Form frmScroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "scroll toolbar"
   ClientHeight    =   1440
   ClientLeft      =   450
   ClientTop       =   1755
   ClientWidth     =   1725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Scroll Toolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   1725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtScroll 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblScroll 
      Alignment       =   2  'Center
      Caption         =   "Please Choose a scroll speed:"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmpicture.Enabled = True
    frmpicture.Timer2.Enabled = False
    frmpicture.MnuStop.Enabled = False
    frmpicture.MnuStart.Enabled = True
    frmpicture.Drive1.Enabled = True
    frmpicture.Dir1.Enabled = True
    Unload frmScroll
    
End Sub

Private Sub cmdOK_Click()
    Dim msg As Integer
    Dim temp As String
    
    temp = txtScroll.Text
    
    If txtScroll.Text = "" Then
    msg = MsgBox("Please enter in a value", vbCritical, "Warning")
    Else
    frmpicture.Enabled = True
    frmpicture.lbltime.Caption = txtScroll.Text
    frmpicture.Timer2.Enabled = True
    frmScroll.Visible = False
    End If
End Sub

Private Sub Form_Load()
    lblScroll.Caption = "Please Choose A Scroll Speed:"
    
    txtScroll.MaxLength = "10"
    txtScroll.Text = ""
    frmpicture.Timer2.Enabled = False
End Sub


Private Sub txtScroll_KeyPress(KeyAscii As Integer)
    Dim msg As Integer
    Dim temp As String
    
    temp = txtScroll.Text
    
    If KeyAscii = 13 Then
        If txtScroll.Text = "" Then
        msg = MsgBox("Please enter in a value", vbCritical, "Warning")
        Else
            If Asc(temp) < 48 Or Asc(temp) > 57 Then
            msg = MsgBox("Value must be numerical, please try again", vbCritical, "WSarning")
            txtScroll.Text = ""
            Else
            frmpicture.Enabled = True
            frmpicture.lbltime.Caption = txtScroll.Text
            frmpicture.Timer2.Enabled = True
            frmScroll.Visible = False
            End If
        End If
    End If
End Sub

