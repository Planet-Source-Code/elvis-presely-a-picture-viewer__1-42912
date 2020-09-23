VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "rsa's mini picture viewer"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   5085
   ClientWidth     =   5535
   Icon            =   "Mini Picture Viewer.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Mini Picture Viewer.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   3405
   ScaleWidth      =   5535
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2520
      MouseIcon       =   "Mini Picture Viewer.frx":0D0C
      MousePointer    =   1  'Arrow
      Pattern         =   "*.jpg;*.gif;*.bmp;*.png;*.psd;*.tif;*.pcx;*.wbmp;*.wmf;*.emf;*.cur;*.ico;*.ico|"
      System          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2175
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
      MouseIcon       =   "Mini Picture Viewer.frx":15D6
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   2295
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
      Height          =   2790
      Left            =   120
      MouseIcon       =   "Mini Picture Viewer.frx":1EA0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtMsg 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   120
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox txtedit 
      Height          =   735
      Left            =   720
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1800
      Top             =   1440
   End
   Begin VB.Label lblMovie 
      Caption         =   "0"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDir 
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblFile 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image picture1 
      Height          =   1815
      Left            =   4800
      MousePointer    =   1  'Arrow
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHide1 
         Caption         =   "&Hide"
         Shortcut        =   ^H
      End
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
      Begin VB.Menu MnuHidden 
         Caption         =   "Show &Hidden Files..."
         Shortcut        =   ^K
      End
      Begin VB.Menu MnuPScroll 
         Caption         =   "Show &Picture Scroller..."
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuMovie 
         Caption         =   "Show &Movie Player..."
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MnuDeleteFolder 
         Caption         =   "D&elete Folder"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDefault 
         Caption         =   "D&efault"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuColour 
         Caption         =   "&Colour"
         Begin VB.Menu MnuFont 
            Caption         =   "&Font..."
            Shortcut        =   ^F
         End
         Begin VB.Menu MnuBackground 
            Caption         =   "&Background..."
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
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDate 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuScroll 
      Caption         =   "&Scroll"
      Begin VB.Menu MnuStart 
         Caption         =   "St&art..."
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuStop 
         Caption         =   "St&op"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu MnuTimeS 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuDateS 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMovieS 
         Caption         =   "Show &Movie Player..."
      End
      Begin VB.Menu Mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExitS 
         Caption         =   "E&xit"
      End
      Begin VB.Menu MnuHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu MnuShow 
         Caption         =   "&Show"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmpicture"
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
    End If
    
    frmpicture.ScaleMode = 1
    frmpicture.Caption = "rsa's mini picture viewer - " & " " & File1.Listcount
    picture1.Visible = False
    frmpicture.Width = Dir1.Width + File1.Width + 500
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
    
    frmpicture.ScaleMode = 1
    frmpicture.Caption = "rsa's mini picture viewer"
    picture1.Visible = False
    frmpicture.Width = Dir1.Width + File1.Width + 500
    
End Sub

Private Sub File1_Click()
    Dim Strfile, Sizing, Screensize, List, Listindex, Listcount As String
    Dim msg As Integer
    
    On Error GoTo 1
    
    frmpicture.ScaleMode = 1
    Strfile = File1.Path & "\" & File1.FileName
    Listindex = File1.Listindex + 1
    Listcount = File1.Listcount
    List = Listindex & " of " & Listcount
       
    picture1.Picture = LoadPicture(Strfile)
    picture1.Visible = False
    picture1.Stretch = False
    Screensize = Screen.Width - (Dir1.Width + File1.Width + "500")
    
    If Right$(File1.FileName, 4) = ".gif" Then
    Strfile = File1.Path & "\" & File1.FileName
    frmPictureL.Show
    frmPictureL.picture1.Show (Strfile)
    frmPictureL.Caption = "rsa's mini picture viewer - " & File1.FileName & ""
    End If
    
    If frmpicture.WindowState = 0 Then
        
        frmpicture.Width = Drive1.Width + File1.Width + "500" + picture1.Width
        
        If picture1.Picture.Height > File1.Height Then
        picture1.Stretch = True
        picture1.Height = File1.Height
        Sizing = picture1.Picture.Height / File1.Height
        picture1.Width = picture1.Picture.Width / Sizing
        frmpicture.Width = Drive1.Width + File1.Width + "500" + picture1.Width
        frmpicture.Height = File1.Height + "1000"
        End If
    
        If picture1.Width > Screensize Then
        picture1.Stretch = True
        picture1.Width = Screensize
        Sizing = picture1.Picture.Width / Screensize
        picture1.Height = picture1.Picture.Height / Sizing
        frmpicture.Width = Drive1.Width + File1.Width + "500" + picture1.Width
        frmpicture.Height = File1.Height + "1000"
        End If
        
    Else
        
        If picture1.Picture.Height > File1.Height Then
        picture1.Stretch = True
        picture1.Height = File1.Height
        Sizing = picture1.Picture.Height / File1.Height
        picture1.Width = picture1.Picture.Width / Sizing
        End If
        
        If picture1.Width > Screensize Then
        picture1.Stretch = True
        picture1.Width = Screensize
        Sizing = picture1.Picture.Width / Screensize
        picture1.Height = picture1.Picture.Height / Sizing
        End If
        
    End If
    
    picture1.Visible = True
    
    If Error = True Then
1   msg = MsgBox("Invalid picture format", vbExclamation + vbOKOnly, "Warning")
    End If

    frmpicture.Caption = "rsa's mini picture viewer - " & List & " - " & File1.FileName
    lblFile.Caption = File1.Listindex
    
    frmpicture.ScaleMode = 3

End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    
    InitControlsXP
End Sub

Private Sub Form_Load()
    Dim Screensize, Screensize1, List, Listindex, Listcount, Screen1, Screen2 As String
    
    Listindex = File1.Listindex + 1
    Listcount = File1.Listcount
    List = Listindex & " of " & Listcount
    Screensize = Screen.Height / 2
    Screensize1 = Screensize - 350
    Screen1 = Screen.Height
    
    On Error GoTo 1
    
    Dir1.Path = "c:\documents\pics"
    
    If Error = True Then
    Err.Clear
1   Dir1.Path = "c:\"
    Drive1 = Dir1
    End If

    frmpicture.ScaleMode = 1

    File1.Pattern = "*.jpg;*.jpeg;*.gif;*.bmp;*.png;*.psd;*.tif;*.pcx;*.wbmp;*.wmf;*.emf;*.cur;*.ico"
    picture1.BorderStyle = 1
    
    frmpicture.Caption = "rsa's mini picture viewer"
    
    MnuStop.Visible = True
    MnuStop.Enabled = False
    
    Timer2.Enabled = False
    
    File1.Height = Screensize1
    Dir1.Height = File1.Height - "250"
    frmpicture.Height = File1.Height + "1000"
    Screen2 = frmpicture.Height + 450
    
    With IconData

    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmpicture.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    
    frmpicture.WindowState = 1
    frmpicture.Left = "0"
    frmpicture.Top = Screen1 - Screen2
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    Call MnuHide_Click
    Cancel = True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If frmpicture.WindowState = 2 Then
    lblSize.Caption = "Max"
    frmpicture.ScaleMode = 1
    Else
    lblSize.Caption = "Min"
    frmpicture.ScaleMode = 1
    End If
    
    If frmpicture.WindowState = 1 Then
    With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmpicture.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    frmpicture.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    frmpicture.Hide
    End If
    
End Sub

Private Sub lblSize_Change()
    Dim Screensize, Screensize1 As String
    
    On Error Resume Next
    
    Screensize = Screen.Height / 2
    Screensize1 = Screensize - "350"
    
    frmpicture.ScaleMode = 1
    
    If lblSize.Caption = "Max" Then
    File1.Height = frmpicture.Height
    Dir1.Height = File1.Height - "1250"
    File1.Height = File1.Height - "900"
    End If
    
    If lblSize.Caption = "Min" Then
    File1.Height = Screensize1
    Dir1.Height = File1.Height - "250"
    frmpicture.Height = File1.Height + "1000"
    End If
    
End Sub

Private Sub MnuAbout_Click()
    On Error Resume Next
    
    frmAbout.Show
    frmpicture.Enabled = False
    frmpicture.Visible = False
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
    frmpicture.BackColor = cd.Color
    End If
End Sub

Private Sub MnuDefault_Click()
    On Error Resume Next
    
    MnuShowAll.Checked = False
    Drive1.BackColor = &HFFFFFF
    Dir1.BackColor = &HFFFFFF
    File1.BackColor = &HFFFFFF
    frmpicture.BackColor = &H8000000F
    Drive1.ForeColor = &H0&
    Dir1.ForeColor = &H0&
    File1.ForeColor = &H0&
End Sub

Private Sub MnuDelete_Click()
    On Error Resume Next
    
    Dim Strfile As String
    Dim msg As Integer
    
    Strfile = File1.Path & "\" & File1.FileName
    msg = MsgBox("Are you sure that you want to delete - " & File1.FileName & "", vbExclamation + vbYesNo, "Warning")
    
    If msg = vbYes Then
    
:   Kill (Strfile)
    File1.Refresh
    File1.Listindex = lblFile.Caption - 1
    End If
    
    If msg = vbNo Then
    Cancel = True
    End If
End Sub

Private Sub MnuDeleteFolder_Click()
    On Error Resume Next
    
    Dim Strfile, StrDir As String
    Dim msg As Integer
    
    Strfile = Dir1 & "\" & "*.*"
    StrDir = Dir1.Path
    
    msg = MsgBox("Are you sure that you want to delete - " & Dir1.Path & " and all of its contents", vbExclamation + vbYesNo, "Warning")
    
    If msg = vbYes Then
:   Kill (Strfile)
:   RmDir (StrDir)
    End If
    
    If vbNo Then
    Cancel = True
    End If

End Sub

Private Sub MnuExit_Click()
    On Error Resume Next
    
    Unload Me
    
End Sub

Private Sub MnuExitS_Click()
    On Error Resume Next
    
    Shell_NotifyIcon NIM_DELETE, IconData
    End
    
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
    End If
End Sub

Private Sub MnuHidden_Click()
    On Error Resume Next
    
    If MnuHidden.Checked = True Then
    frmpicture.MnuHidden.Checked = False
    frmpicture.File1.Archive = False
    frmpicture.File1.Hidden = False
    frmpicture.File1.System = False
    File1.Refresh
    picture1.Refresh
    Else
    frmPassword.Show
    End If
    
End Sub

Private Sub MnuHide_Click()
    On Error Resume Next
    
    frmpicture.WindowState = 1
    
End Sub

Private Sub MnuHide1_Click()
    On Error Resume Next
    
    Call MnuHide_Click
    
End Sub

Private Sub MnuMovie_Click()
    On Error Resume Next
    
    frmMovie.Show
    lblMovie.Caption = "1"
    frmMovie.SetFocus
    
End Sub

Private Sub MnuMovieS_Click()
    On Error Resume Next

    Call MnuMovie_Click
End Sub

Private Sub MnuPrint_Click()
    Dim strprint As String
    
    strprint = File1.Path & "\" & File1.FileName
    
    On Error Resume Next
    
    cd.CancelError = True
    cd.ShowPrinter

    If Err.Number = cdlCancel Then
    Err.Clear
    Else
    Printer.Copies = cd.Copies
    Printer.Print picture1.Picture
    Printer.EndDoc
    End If
End Sub

Private Sub MnuPScroll_Click()
    On Error Resume Next
    
    frmPScroll.Show
    frmPScroll.Visible = False
    frmpicture.Visible = False
    
    
End Sub

Private Sub MnuRefresh_Click()
    On Error Resume Next
    
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
    
    File1.Listindex = lblFile.Caption
End Sub

Private Sub MnuShow_Click()
    Dim Screensize, Screensize1 As String
    
    On Error Resume Next
    
    frmpicture.WindowState = vbNormal
    frmpicture.Show
    Screensize = Screen.Height / 2
    Screensize1 = Screensize - "350"
      
      With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = frmpicture.Caption & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
    End With
    frmpicture.ScaleMode = 3
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    
    frmpicture.ScaleMode = 1
    
    File1.Height = Screensize1
    Dir1.Height = File1.Height - "250"
    frmpicture.Height = File1.Height + "1000"
    
End Sub

Private Sub MnuShowAll_Click()
    On Error Resume Next
    
    If MnuShowAll.Checked = False Then
    MnuShowAll.Checked = True
    File1.Pattern = "*.*"
    Else
    MnuShowAll.Checked = False
    File1.Pattern = "*.jpg;*.jpeg;*.gif;*.bmp;*.png;*.psd;*.tif;*.pcx;*.wbmp;*.wmf;*.emf;*.cur;*.ico"
    End If
End Sub

Private Sub MnuStart_Click()
    Dim msg As Integer
    
    On Error Resume Next
    
    If File1.FileName = "" Then
    msg = MsgBox("Please select a start picture", vbCritical, "Warning")
    Else
    lbltime.Caption = frmScroll.txtScroll.Text
    Timer2.Enabled = False
    MnuStop.Enabled = True
    MnuStart.Enabled = False
    frmScroll.Show
    frmpicture.Enabled = False
    frmScroll.txtScroll.Text = ""
    Dir1.Enabled = False
    Drive1.Enabled = False
    End If
    
End Sub

Private Sub MnuStop_Click()
    On Error Resume Next
    
    Timer2.Enabled = False
    MnuStart.Enabled = True
    MnuStop.Enabled = False
    Dir1.Enabled = True
    Drive1.Enabled = True
    Unload frmScroll
End Sub

Private Sub picture1_DblClick()
    Dim Strfile As String
    
    On Error Resume Next
    
    Strfile = File1.Path & "\" & File1.FileName
    
    frmPictureL.Show
    frmPictureL.picture1.Show Strfile
    frmPictureL.Caption = "rsa's mini picture viewer - " & File1.FileName & ""
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    
    MnuTime.Caption = Time
    MnuDate.Caption = Date
    MnuTimeS.Caption = Time
    MnuDateS.Caption = Date
End Sub

Private Sub Timer2_Timer()
    Dim msg As Integer
    
    On Error GoTo 1
    
    lbltime.Caption = lbltime.Caption - 1
    
    If lbltime.Caption = "0" Then
        If File1.FileName = "" Then
        Timer2.Enabled = False
        msg = MsgBox("Please select start picture", vbCritical, "Warning")
        Else
        lbltime.Caption = frmScroll.txtScroll.Text
        File1.Listindex = File1.Listindex + 1
        End If
    End If
    
    If Error = True Then
1   File1.Listindex = File1.Listindex
    Timer2.Enabled = False
    End If
    
    If Timer2.Enabled = False Then
    Drive1.Enabled = True
    Dir1.Enabled = True
    MnuStop.Enabled = False
    MnuStart.Enabled = True
    Unload frmScroll
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim msg As Long

    msg = x

    If msg = WM_LBUTTONDBLCLK Then
        If frmpicture.WindowState = 1 Then
        Call MnuShow_Click
        Else
        Call MnuHide_Click
        End If
    Else
        If msg = WM_RBUTTONDOWN Then
        PopupMenu MnuPopUp
        End If
    End If
End Sub
