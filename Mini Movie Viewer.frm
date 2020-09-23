VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMovie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "rsa's mini movie viewer"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   2820
   ClientWidth     =   6180
   Icon            =   "Mini Movie Viewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6180
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   2520
      Pattern         =   "*.mpg;*.mpeg;*.avi*"
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   120
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MediaPlayerCtl.MediaPlayer mp 
      Height          =   3375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   30
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   1
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   2.5
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   -600
      WindowlessVideo =   0   'False
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
      Begin VB.Menu MnuShowAll 
         Caption         =   "Show &All Files"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuPicture 
         Caption         =   "Show &Picture Viewer"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDefault 
         Caption         =   "D&efault"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuColour 
         Caption         =   "&Colour"
         Begin VB.Menu MnuFont 
            Caption         =   "&Font"
            Shortcut        =   ^F
         End
         Begin VB.Menu MnuBackground 
            Caption         =   "&Background"
            Shortcut        =   ^B
         End
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuTime 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDate 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuControls 
      Caption         =   "&Controls"
      Begin VB.Menu MnuPLay 
         Caption         =   "&Play"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MnuStop 
         Caption         =   "&Stop"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MnuPause 
         Caption         =   "P&ause"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNavigate 
         Caption         =   "&Navigate"
         Begin VB.Menu MnuRewind 
            Caption         =   "&Rewind"
            Shortcut        =   +{F6}
         End
         Begin VB.Menu MnuForward 
            Caption         =   "&Fast Forward"
            Shortcut        =   +{F7}
         End
         Begin VB.Menu mnu10 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSkipBack 
            Caption         =   "Skip &Back"
            Shortcut        =   +{F8}
         End
         Begin VB.Menu MnuSkipForward 
            Caption         =   "Skip F&orward"
            Shortcut        =   +{F9}
         End
      End
      Begin VB.Menu MnuSpeed 
         Caption         =   "&Speed"
         Begin VB.Menu MnuSlow 
            Caption         =   "&Slow"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu MnuNormal 
            Caption         =   "&Normal"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MnuFast 
            Caption         =   "&Fast"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnu7 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSlower 
            Caption         =   "S&lower"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu MnuFaster 
            Caption         =   "F&aster"
            Shortcut        =   +{F5}
         End
      End
      Begin VB.Menu MnuVolume 
         Caption         =   "&Volume"
         Begin VB.Menu MnuVolup 
            Caption         =   "&Up"
            Shortcut        =   +{F11}
         End
         Begin VB.Menu MnuVolDown 
            Caption         =   "&Down"
            Shortcut        =   +{F12}
         End
         Begin VB.Menu mnu4 
            Caption         =   "-"
         End
         Begin VB.Menu MnuMute 
            Caption         =   "&Mute"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu MnuZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu Mnu50 
            Caption         =   "&50%"
            Shortcut        =   ^Q
         End
         Begin VB.Menu Mnu100 
            Caption         =   "&100%"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnu9 
            Caption         =   "-"
         End
         Begin VB.Menu MnuFullScreen 
            Caption         =   "&Full Screen"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu MnuShow 
         Caption         =   "S&how"
         Begin VB.Menu MnuAudio 
            Caption         =   "&Audio Controls"
         End
         Begin VB.Menu MnuCaptioning 
            Caption         =   "&Captioning"
         End
         Begin VB.Menu MnuDisplay 
            Caption         =   "&Display"
         End
         Begin VB.Menu MnuGotoBar 
            Caption         =   "&Goto Bar"
         End
         Begin VB.Menu MnuPositionControl 
            Caption         =   "&Position Control"
         End
         Begin VB.Menu MnuStatusBar 
            Caption         =   "&Status Bar"
         End
         Begin VB.Menu MnuTracker 
            Caption         =   "&Tracker"
         End
         Begin VB.Menu mnu11 
            Caption         =   "-"
         End
         Begin VB.Menu MnuStatistics 
            Caption         =   "S&tatistics..."
         End
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu MnuMenu 
         Caption         =   "&Menu..."
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&Help..."
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuShowControls 
         Caption         =   "Sh&ow Controls"
         Shortcut        =   ^{F8}
      End
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    On Error Resume Next
    
    If Err.Number = cdlCancel Then
    Err.Clear
    Else
    Drive1 = Dir1
    File1 = Dir1
    mp.Stop
    mp.Visible = False
    mp.Enabled = False
    frmMovie.Width = Drive1.Width + File1.Width + "500"
    frmMovie.Caption = "rsa's mini movie viewer"
    End If
    
End Sub

Private Sub Drive1_Change()
    Dim msg As Integer

    On Error GoTo 1
    
    If Error = True Then
1   msg = MsgBox("Please insert disk and try again", vbOKOnly + vbExclamation, "Warning")
    Err.Clear
    End If
    
    If msg = 1 Then
    Drive1 = Dir1
    End If
    
    Dir1 = Drive1
    File1 = Dir1
    mp.Stop
    mp.Visible = False
    mp.Enabled = False
    frmMovie.Width = Drive1.Width + File1.Width + "500"
    frmMovie.Caption = "rsa's mini movie viewer"
    
End Sub

Private Sub File1_Click()
    Dim Strfile, Sizing, Max As String
    Dim msg As Integer
    
    On Error GoTo 1
    
    File1 = Dir1
    Drive1 = Dir1
    Strfile = File1.Path & "\" & File1.FileName
    
    mp.Visible = False
    mp.Enabled = False
    
    If frmMovie.WindowState = 0 Then
    mp.FileName = Strfile
    mp.Height = File1.Height
    frmMovie.Width = Drive1.Width + File1.Width + "500" + mp.Width
    End If
    
    If Error = True Then
1   msg = MsgBox("Invalid picture format", vbExclamation + vbOKOnly, "Warning")
    End If
    
    mp.Visible = True
    mp.Enabled = True
    
    
    frmMovie.Caption = "rsa's mini movie viewer - " & File1.FileName & ""
    frmMovie.Show
    mp.Rate = 1

End Sub
Private Sub Form_Load()
    Dim Screen1, Screen2 As String
    
    On Error GoTo 1
    
    Dir1.Path = "c:\documents and settings\administrator\documents\movies"
    Screen1 = Screen.Height
    Screen2 = frmMovie.Height + 450
    
    If Error = True Then
    Err.Clear
1   Dir1.Path = "c:\"
    Drive1 = Dir1
    End If

    File1.Pattern = "*.mpg;*.mpeg;*.avi;*.wmv;*.asf;*.wm;*.wma;*.wmv;*.m1v;*.mp2;*.mp3;*.cda;*.wav;*.snd;*.mid;*.rmi;*.midi"
    
    mp.Visible = False
    
    MnuShowControls.Checked = True
    MnuAudio.Checked = True
    MnuPositionControl.Checked = True
    MnuTracker.Checked = True
    Mnu100.Checked = True
    MnuNormal.Checked = True
    
    frmMovie.WindowState = 0
    frmMovie.Left = "0"
    Screen1 = Screen.Height
    Screen2 = frmMovie.Height + 450
    frmMovie.Top = Screen1 - Screen2
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg As Integer
    
    On Error Resume Next
    
    msg = MsgBox("Are you sure that you want to exit?", vbYesNo + vbExclamation, "Warning")
    
    If msg = vbYes Then
    Unload Me
    End If
    
    If msg = vbNo Then
    Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    Dim Screensize, Screensize1 As String
    
    On Error Resume Next
    
    Screensize = Screen.Height / 2
    Screensize1 = Screensize - 350
    
    If frmMovie.WindowState = 0 Then
    File1.Height = Screensize1
    Dir1.Height = File1.Height - "250"
    frmMovie.Height = File1.Height + "1000"
    End If
    
End Sub

Private Sub Mnu100_Click()
    On Error Resume Next
    
    If Mnu100.Checked = False Then
    Mnu50.Checked = False
    Mnu100.Checked = True
    MnuFullScreen.Checked = False
    mp.DisplaySize = 4
    End If
End Sub

Private Sub Mnu50_Click()
    On Error Resume Next
    
    If Mnu50.Checked = False Then
    Mnu100.Checked = False
    Mnu50.Checked = True
    MnuFullScreen.Checked = False
    mp.DisplaySize = 1
    End If
End Sub

Private Sub MnuAudio_Click()
    On Error Resume Next
    
    If MnuAudio.Checked = True Then
    MnuAudio.Checked = False
    mp.ShowAudioControls = False
    Else
    MnuAudio.Checked = True
    mp.ShowAudioControls = True
    End If
End Sub

Private Sub MnuBackground_Click()
    On Error Resume Next
    
    cd.CancelError = True
    cd.Action = 3
    
    If Err.Number = cdlCancel Then
    Err.Clear
    cd.Color = File1.BackColor
    Else
    Drive1.BackColor = cd.Color
    Dir1.BackColor = cd.Color
    File1.BackColor = cd.Color
    frmMovie.BackColor = cd.Color
    mp.DisplayBackColor = cd.Color
    End If
End Sub

Private Sub MnuCaptioning_Click()
    On Error Resume Next
    
    If MnuCaptioning.Checked = False Then
    MnuCaptioning.Checked = True
    mp.ShowCaptioning = True
    Else
    MnuCaptioning.Checked = False
    mp.ShowCaptioning = False
    End If
End Sub

Private Sub MnuDefault_Click()
    On Error Resume Next
    
    MnuShowAll.Checked = False
    Drive1.BackColor = &HFFFFFF
    Dir1.BackColor = &HFFFFFF
    File1.BackColor = &HFFFFFF
    frmMovie.BackColor = &H8000000F
    mp.DisplayBackColor = &H0&
    Drive1.ForeColor = &H0&
    Dir1.ForeColor = &H0&
    File1.ForeColor = &H0&
    mp.DisplayForeColor = &HFFFFFF
    
End Sub

Private Sub MnuDelete_Click()
    Dim Strfile As String
    Dim msg As Integer
    
    Strfile = File1.Path & "\" & File1.FileName
    msg = MsgBox("Are you sure that you want to delete - " & File1.FileName & "", vbExclamation + vbYesNo, "Warning")
    
    On Error Resume Next
    
    If msg = vbYes Then
    
:   Kill (Strfile)
    File1.Refresh
    File1.Listindex = File1.TopIndex
    End If
    
    If msg = vbNo Then
    Cancel = True
    End If
End Sub

Private Sub MnuDisplay_Click()
    On Error Resume Next
    
    If MnuDisplay.Checked = False Then
    MnuDisplay.Checked = True
    mp.ShowDisplay = True
    Else
    MnuDisplay.Checked = False
    mp.ShowDisplay = False
    End If
End Sub

Private Sub MnuExit_Click()
    On Error Resume Next
    
    Unload Me
End Sub

Private Sub MnuFast_Click()
    On Error Resume Next
    
    If MnuFast.Checked = False Then
    MnuFast.Checked = True
    MnuSlow.Checked = False
    MnuNormal.Checked = False
    MnuSlower.Checked = False
    MnuFaster.Checked = False
    mp.Rate = 2
    End If
End Sub

Private Sub MnuFaster_Click()
    On Error Resume Next
    
    mp.Rate = mp.Rate + 0.25
    
    If mp.Rate < 1 Then
    MnuSlow.Checked = True
    MnuSlower.Checked = False
    MnuNormal.Checked = False
    MnuFast.Checked = False
    MnuFaster.Checked = True
    End If
     
    If mp.Rate = 1 Then
    MnuSlow.Checked = False
    MnuSlower.Checked = False
    MnuNormal.Checked = True
    MnuFast.Checked = False
    MnuFaster.Checked = True
    End If
    
    If mp.Rate > 1 Then
    MnuSlow.Checked = False
    MnuSlower.Checked = False
    MnuNormal.Checked = False
    MnuFast.Checked = True
    MnuFaster.Checked = True
    End If
End Sub

Private Sub MnuFont_Click()
    On Error Resume Next
    
    cd.CancelError = True
    cd.Action = 3
    
    If Err.Number = cdlCancel Then
    Err.Clear
    cd.Color = File1.ForeColor
    Else
    Drive1.ForeColor = cd.Color
    Dir1.ForeColor = cd.Color
    File1.ForeColor = cd.Color
    mp.DisplayForeColor = cd.Color
    End If
End Sub

Private Sub MnuForward_Click()
    On Error Resume Next
    
    mp.CurrentPosition = mp.CurrentPosition + 50
End Sub

Private Sub MnuFullScreen_Click()
    On Error Resume Next
    
    If mp.DisplaySize = 3 Then
    Mnu50.Checked = False
    Mnu100.Checked = False
    MnuFullScreen.Checked = False
    Else
    mp.DisplaySize = 3
    Mnu50.Checked = False
    Mnu100.Checked = True
    MnuFullScreen.Checked = False
    End If
    
End Sub

Private Sub MnuGotoBar_Click()
    On Error Resume Next
    
    If MnuGotoBar.Checked = False Then
    MnuGotoBar.Checked = True
    mp.ShowGotoBar = True
    Else
    MnuGotoBar.Checked = False
    mp.ShowGotoBar = False
    End If
End Sub

Private Sub MnuHelp_Click()
    On Error Resume Next
    
    mp.ShowDialog mpShowDialogHelp
End Sub

Private Sub MnuMenu_Click()
    On Error Resume Next
    
    mp.ShowDialog mpShowDialogContextMenu
End Sub

Private Sub MnuMute_Click()
    On Error Resume Next
    
    If MnuMute.Checked = False Then
    MnuMute.Checked = True
    mp.Mute = True
    Else
    MnuMute.Checked = False
    mp.Mute = False
    End If
End Sub

Private Sub MnuNormal_Click()
    On Error Resume Next
    
    If MnuNormal.Checked = False Then
    MnuNormal.Checked = True
    MnuSlow.Checked = False
    MnuFast.Checked = False
    MnuSlower.Checked = False
    MnuFaster.Checked = False
    mp.Rate = 1
    End If
End Sub

Private Sub MnuOptions_Click()
    On Error Resume Next
    
    mp.ShowDialog mpShowDialogOptions
End Sub

Private Sub MnuPause_Click()
    On Error Resume Next
    
    mp.Pause
End Sub

Private Sub MnuPicture_Click()
    On Error Resume Next
    
    frmpicture.Show
End Sub

Private Sub MnuPlay_Click()
    On Error Resume Next
    
    mp.Play
End Sub

Private Sub MnuPositionControl_Click()
    On Error Resume Next
    
    If MnuPositionControl.Checked = True Then
    MnuPositionControl.Checked = False
    mp.ShowPositionControls = False
    Else
    MnuPositionControl.Checked = True
    mp.ShowPositionControls = True
    End If
End Sub

Private Sub MnuRefresh_Click()
    On Error Resume Next
    
    File1.Refresh
    Dir1.Refresh
    Drive1.Refresh
End Sub

Private Sub MnuRewind_Click()
    On Error Resume Next
    
    mp.CurrentPosition = mp.CurrentPosition - 50
End Sub

Private Sub MnuShowAll_Click()
    On Error Resume Next
    
    If MnuShowAll.Checked = False Then
    MnuShowAll.Checked = True
    File1.Pattern = "*.*"
    Else
    MnuShowAll.Checked = False
    File1.Pattern = "*.mpg;*.mpeg;*.avi;*.wmv;*.asf;*.wm;*.wma;*.wmv;*.m1v;*.mp2;*.mp3;*.cda;*.wav;*.snd;*.mid;*.rmi;*.midi"
    End If
End Sub

Private Sub MnuShowControls_Click()
    On Error Resume Next
    
    If MnuShowControls.Checked = False Then
    MnuShowControls.Checked = True
    mp.ShowControls = True
    MnuAudio.Enabled = True
    MnuAudio.Checked = True
    mp.ShowAudioControls = True
    MnuCaptioning.Enabled = True
    MnuCaptioning.Checked = False
    mp.ShowCaptioning = False
    MnuDisplay.Enabled = True
    MnuDisplay.Checked = False
    mp.ShowDisplay = False
    MnuGotoBar.Enabled = True
    MnuGotoBar.Checked = False
    mp.ShowGotoBar = False
    MnuPositionControl.Enabled = True
    MnuPositionControl.Checked = True
    mp.ShowPositionControls = True
    MnuStatusBar.Enabled = True
    MnuStatusBar.Checked = False
    mp.ShowStatusBar = False
    MnuTracker.Enabled = True
    MnuTracker.Checked = True
    mp.ShowTracker = True
    Else
    MnuShowControls.Checked = False
    mp.ShowControls = False
    MnuAudio.Enabled = False
    MnuAudio.Checked = False
    mp.ShowAudioControls = False
    MnuCaptioning.Enabled = False
    MnuCaptioning.Checked = False
    mp.ShowCaptioning = False
    MnuDisplay.Enabled = False
    MnuDisplay.Checked = False
    mp.ShowDisplay = False
    MnuGotoBar.Enabled = False
    MnuGotoBar.Checked = False
    mp.ShowGotoBar = False
    MnuPositionControl.Enabled = False
    MnuPositionControl.Checked = False
    mp.ShowPositionControls = False
    MnuStatusBar.Enabled = False
    MnuStatusBar.Checked = False
    mp.ShowStatusBar = False
    MnuTracker.Enabled = False
    MnuTracker.Checked = False
    mp.ShowTracker = False
    End If
End Sub

Private Sub MnuSkipBack_Click()
    On Error Resume Next
    
    mp.CurrentPosition = mp.CurrentPosition - 25
End Sub

Private Sub MnuSkipForward_Click()
    On Error Resume Next
    
    mp.CurrentPosition = mp.CurrentPosition + 25
End Sub

Private Sub MnuSlow_Click()
    On Error Resume Next
    
    If MnuSlow.Checked = False Then
    MnuSlow.Checked = True
    MnuNormal.Checked = False
    MnuFast.Checked = False
    MnuSlower.Checked = False
    MnuFaster.Checked = False
    mp.Rate = 0.5
    End If
End Sub

Private Sub MnuSlower_Click()
    On Error Resume Next
    
    mp.Rate = mp.Rate - 0.25
    
    If mp.Rate < 1 Then
    MnuSlow.Checked = True
    MnuSlower.Checked = True
    MnuNormal.Checked = False
    MnuFast.Checked = False
    MnuFaster.Checked = False
    End If
     
    If mp.Rate = 1 Then
    MnuSlow.Checked = False
    MnuSlower.Checked = True
    MnuNormal.Checked = True
    MnuFast.Checked = False
    MnuFaster.Checked = False
    End If
    
    If mp.Rate > 1 Then
    MnuSlow.Checked = False
    MnuSlower.Checked = True
    MnuNormal.Checked = False
    MnuFast.Checked = True
    MnuFaster.Checked = False
    End If
End Sub

Private Sub MnuStatistics_Click()
    On Error Resume Next
    
    mp.ShowDialog mpShowDialogStatistics
End Sub

Private Sub MnuStatusBar_Click()
    On Error Resume Next
    
    If MnuStatusBar.Checked = False Then
    MnuStatusBar.Checked = True
    mp.ShowStatusBar = True
    Else
    MnuStatusBar.Checked = False
    mp.ShowStatusBar = False
    End If
End Sub

Private Sub MnuStop_Click()
    On Error Resume Next
    
    mp.Stop
    mp.SelectionStart = 0
End Sub

Private Sub mnuw_Click()
    On Error Resume Next
    
    mp.FastForward
End Sub

Private Sub MnuTracker_Click()
    On Error Resume Next
    
    If MnuTracker.Checked = True Then
    MnuTracker.Checked = False
    mp.ShowTracker = False
    Else
    MnuTracker.Checked = True
    mp.ShowTracker = True
    End If
End Sub

Private Sub MnuVolDown_Click()
    On Error Resume Next
    
    mp.Volume = mp.Volume - 100
End Sub

Private Sub MnuVolup_Click()
    On Error Resume Next
    
    mp.Volume = mp.Volume + 100
End Sub

Private Sub mp_DblClick(Button As Integer, ShiftState As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Call MnuFullScreen_Click
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    
    MnuTime.Caption = Time
    MnuDate.Caption = Date
End Sub
