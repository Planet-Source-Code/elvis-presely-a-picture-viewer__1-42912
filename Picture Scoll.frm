VERSION 5.00
Begin VB.Form frmPScroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   4155
   ClientTop       =   1950
   ClientWidth     =   1725
   ClipControls    =   0   'False
   Icon            =   "Picture Scoll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   1725
   Begin VB.Timer Timer1 
      Interval        =   501
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblnavigation 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuStop 
         Caption         =   "&STOP"
         Shortcut        =   ^S
      End
      Begin VB.Menu Mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSpeed 
         Caption         =   "&Speed"
         Begin VB.Menu MnuFaster 
            Caption         =   "&Faster"
            Shortcut        =   ^A
         End
         Begin VB.Menu MnuSlower 
            Caption         =   "&Slower"
            Shortcut        =   ^Z
         End
      End
      Begin VB.Menu MnuNavigation 
         Caption         =   "&Navigation"
         Begin VB.Menu MnuForward 
            Caption         =   "&Forwards"
            Shortcut        =   ^D
         End
         Begin VB.Menu MnuBackward 
            Caption         =   "&Backward"
            Shortcut        =   ^C
         End
      End
   End
End
Attribute VB_Name = "frmPScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 27 Then
    Unload Me
    frmpicture.Enabled = True
    frmpicture.Visible = True
    End If
    
    If vbKeyPageUp = True Then
    Timer1.Enabled = False
    End If
    
    If KeyAscii = vbKeyPageDown Then
    Timer1.Interval = Timer1.Interval - "1000"
    End If
    
End Sub

Private Sub Form_Load()
    Dim Screen1, Screen2 As String
    
    On Error Resume Next
    
    frmPScroll.Visible = False
    frmPScroll.Caption = "rsa's mini picture viewer"
    frmPScroll.Height = "3750"
    frmPScroll.Width = picture1.Width + "250"
    
    Screen1 = Screen.Height
    Screen2 = frmPScroll.Height + "500"
    frmPScroll.Left = "0"
    frmPScroll.Top = Screen1 - Screen2
    
    MnuStop.Caption = "STOP"
    MnuForward.Checked = True
    lblnavigation.Caption = "F"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    Unload Me
    frmpicture.Enabled = True
    frmpicture.Visible = True
    
End Sub

Private Sub MnuBackward_Click()
    On Error Resume Next
    
    If MnuBackward.Checked = True Then
    MnuBackward.Checked = False
    MnuForward.Checked = True
    lblnavigation.Caption = "F"
    Else
    MnuBackward.Checked = True
    MnuForward.Checked = False
    lblnavigation.Caption = "B"
    End If
    
End Sub

Private Sub MnuExit_Click()
    On Error Resume Next
    
    Unload Me
    frmpicture.Enabled = True
    frmpicture.Visible = True
    
End Sub

Private Sub MnuFaster_Click()
    On Error Resume Next
    
    Timer1.Interval = Timer1.Interval - 500
    
End Sub

Private Sub MnuForward_Click()
    On Error Resume Next
    
    If MnuForward.Checked = True Then
    MnuForward.Checked = False
    MnuBackward.Checked = True
    lblnavigation.Caption = "B"
    Else
    MnuForward.Checked = True
    MnuBackward.Checked = False
    lblnavigation.Caption = "F"
    End If
    
End Sub

Private Sub MnuSlower_Click()
    On Error Resume Next
    
    Timer1.Interval = Timer1.Interval + 500
    
End Sub

Private Sub MnuStop_Click()
    On Error Resume Next
    
    If MnuStop.Caption = "STOP" Then
    MnuStop.Caption = "START"
    Timer1.Enabled = False
    Else
    MnuStop.Caption = "STOP"
    Timer1.Enabled = True
    End If
    
End Sub

Private Sub picture1_DblClick()
    Dim Strfile As String
    
    On Error Resume Next
    
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    
    MnuStop.Caption = "START"
    Timer1.Enabled = False
    frmPictureL.Show
    frmPictureL.picture1.Show Strfile
    frmPictureL.Caption = "rsa's mini picture viewer - " & File1.FileName & ""
    
End Sub

Private Sub Timer1_Timer()
    Dim Strfile, Sizing, Height, Screen1, Screen2 As String
    
    On Error Resume Next
    
    Strfile = frmpicture.File1.FileName
    Height = frmpicture.Height - "1000"
    frmPScroll.Visible = True
    
    If lblnavigation.Caption = "F" Then
    frmpicture.File1.Listindex = frmpicture.File1.Listindex + 1
    End If

    If lblnavigation.Caption = "B" Then
    frmpicture.File1.Listindex = frmpicture.File1.Listindex - 1
    End If
    
    picture1.Picture = frmpicture.picture1.Picture
    frmPScroll.Caption = Strfile
    
    picture1.Stretch = True
    picture1.Height = Height
    Sizing = picture1.Picture.Height / Height
    picture1.Width = picture1.Picture.Width / Sizing

    frmPScroll.Height = Height + "1000"
    frmPScroll.Width = picture1.Width + "300"
    
    Screen1 = Screen.Height
    Screen2 = frmPScroll.Height + "500"
    frmPScroll.Left = "0"
    frmPScroll.Top = Screen1 - Screen2
    
End Sub
