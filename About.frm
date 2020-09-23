VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   5100
   ClientTop       =   2445
   ClientWidth     =   4185
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "< &Previous"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version: 1.1.13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright © 2002 rsa corp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblinfo2 
      Caption         =   "Last Modified on Jan 26th 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Caption         =   "This program was created on July 2nd 2002. by rsa for rsa corp. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      Caption         =   "Filename"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
   Begin VB.Image picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   120
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub CmdNext_Click()
    Dim Strfile, Sizing As String
    
    On Error Resume Next
    
    picture1.Visible = False
    frmpicture.File1.Listindex = frmpicture.File1.Listindex + 1
    picture1.Picture = frmpicture.picture1.Picture
    
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    lblFilename.Caption = Strfile
    
    If picture1.Picture.Height > "2535" Then
    picture1.Stretch = True
    picture1.Height = "2535"
    Sizing = picture1.Picture.Height / "2535"
    picture1.Width = picture1.Picture.Width / Sizing
    End If
    
    If picture1.Width > "2535" Then
    picture1.Stretch = True
    picture1.Width = "2535"
    Sizing = picture1.Picture.Width / "2535"
    picture1.Height = picture1.Picture.Height / Sizing
    End If
    
    picture1.Visible = True
    
End Sub

Private Sub CmdPrevious_Click()
    Dim Strfile, Sizing As String
    
    On Error Resume Next
    
    picture1.Visible = False
    frmpicture.File1.Listindex = frmpicture.File1.Listindex - 1
    picture1.Picture = frmpicture.picture1.Picture
    
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    lblFilename.Caption = Strfile
    
    If picture1.Picture.Height > "2535" Then
    picture1.Stretch = True
    picture1.Height = "2535"
    Sizing = picture1.Picture.Height / "2535"
    picture1.Width = picture1.Picture.Width / Sizing
    End If
    
    If picture1.Width > "2535" Then
    picture1.Stretch = True
    picture1.Width = "2535"
    Sizing = picture1.Picture.Width / "2535"
    picture1.Height = picture1.Picture.Height / Sizing
    End If
    
    picture1.Visible = True
    
End Sub

Private Sub cmdSysInfo_Click()
    On Error Resume Next
  
    Call StartSysInfo
  
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    frmpicture.Enabled = True
    frmpicture.Visible = True
    frmpicture.Show
    Unload Me
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()
    Dim Strfile, Sizing As String
    Dim msg As Integer
    
    If frmpicture.File1.Listindex = 0 Then
    Call CmdNext_Click
    End If
    
    On Error GoTo 1
    
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    
    frmAbout.Caption = "about rsa's mini picture viewer"
    picture1.Picture = LoadPicture(Strfile)
    lblFilename.Caption = Strfile
    
    If picture1.Picture.Height > "2535" Then
    picture1.Stretch = True
    picture1.Height = "2535"
    Sizing = picture1.Picture.Height / "2535"
    picture1.Width = picture1.Picture.Width / Sizing
    End If
    
    If picture1.Width > "2535" Then
    picture1.Stretch = True
    picture1.Width = "2535"
    Sizing = picture1.Picture.Width / "2535"
    picture1.Height = picture1.Picture.Height / Sizing
    End If
    
    If Err.Number = cdlCancel Then
1   Err.Clear
    End If
    
End Sub

Private Sub picture1_DblClick()
    Dim Strfile As String
    
    On Error Resume Next
    
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    
    frmPictureL.Show
    frmPictureL.picture1.Show Strfile
    frmPictureL.Caption = "rsa's mini picture viewer - " & frmpicture.File1.FileName & ""
    
End Sub
