Attribute VB_Name = "Mod1"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Function AppPath(ByVal zPath As String) As String
  If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"
End Function

Private Function FileExist(ByVal strPath As String) As Boolean
  On Local Error GoTo ErrFile
  Open strPath For Input Access Read As #1
  Close #1
  FileExist = True
  Exit Function
ErrFile:
  FileExist = False
End Function

Private Sub MakeManifest()
  Dim file$, file2$, qwe As String
  file$ = AppPath(App.Path) & App.EXEName & ".exe.MANIFEST"
  If Not FileExist(file$) Then
    qwe = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf _
        & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf _
        & "<assemblyIdentity type=""win32"" processorArchitecture=""*"" version=""6.0.0.0"" name=""name""/>" & vbCrLf _
        & "<description>Enter your Description Here</description>" & vbCrLf _
        & "<dependency>" & vbCrLf _
        & "   <dependentAssembly>" & vbCrLf _
        & "      <assemblyIdentity" & vbCrLf _
        & "           type=""win32""" & vbCrLf _
        & "           name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0""" & vbCrLf _
        & "           language=""*""" & vbCrLf _
        & "           processorArchitecture=""*""" & vbCrLf _
        & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf _
        & "      />" & vbCrLf _
        & "   </dependentAssembly>" & vbCrLf _
        & "</dependency>" & vbCrLf _
        & "</assembly>" & vbCrLf
    Open file$ For Binary Access Write Lock Write As #1 Len = 1
    Put #1, , qwe
    Close #1
    SetAttr file$, vbReadOnly Or vbHidden ' Or vbSystem
    file2$ = AppPath(App.Path) & App.EXEName & ".exe"
    Shell file2$, vbNormalFocus
    ExitProcess 1
  End If
End Sub

Public Sub InitControlsXP()
  MakeManifest
  InitCommonControls
End Sub
