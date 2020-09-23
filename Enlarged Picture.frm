VERSION 5.00
Object = "{50F16B18-467E-11D1-8271-00C04FC3183B}#1.0#0"; "shimgvw.dll"
Begin VB.Form frmPictureL 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   1485
   ClientWidth     =   12930
   Icon            =   "Enlarged Picture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   12930
   Begin VB.CommandButton Previous 
      Caption         =   "< Previous"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Next 
      Caption         =   "Next >"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   10080
      Width           =   1215
   End
   Begin PREVIEWLibCtl.Preview picture1 
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "frmPictureL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    
    frmPictureL.Caption = "rsa's mini picture viewer - " & frmpicture.File1.FileName & ""
    frmPictureL.WindowState = 2
    
End Sub

Private Sub Next_Click()
    Dim Strfile As String
    
    On Error Resume Next
    
    frmpicture.File1.Listindex = frmpicture.File1.Listindex + 1
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    picture1.Show Strfile
    frmAbout.picture1.Picture = LoadPicture(Strfile)
    frmPictureL.Caption = "rsa's mini picture viewer - " & frmpicture.File1.FileName & ""
    
End Sub

Private Sub Previous_Click()
    Dim Strfile As String
    
    On Error Resume Next
    
    frmpicture.File1.Listindex = frmpicture.File1.Listindex - 1
    Strfile = frmpicture.File1.Path & "\" & frmpicture.File1.FileName
    picture1.Show Strfile
    frmAbout.picture1.Picture = LoadPicture(Strfile)
    frmPictureL.Caption = "rsa's mini picture viewer - " & frmpicture.File1.FileName & ""
    
End Sub
