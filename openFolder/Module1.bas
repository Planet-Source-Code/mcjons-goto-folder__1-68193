Attribute VB_Name = "Module1"
'--------- "Browse for Folder"
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Public Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Public Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
'---------- End of browse for folder ------------


Public strTujuan As String
Public openfolder As String

'for saving to notepad
Option Explicit
Public dirF, test, line As Variant
Public delfolder As Integer
Public Sub loadfiles()
     Open dirF For Input As #1
        Do Until EOF(1)
            Line Input #1, line
            If line <> "" Then
                
                
                    Form1.dirList.AddItem (line)
                
            End If
        Loop
    Close #1
End Sub
Public Sub savefiles()
    Dim i As Integer
    Dim data As Variant
    Open dirF For Output As #1
        data = ""
        Form1.dirList.ListIndex = -1
        For i = 0 To Form1.dirList.ListCount
            If Form1.dirList.List(i) <> "" Then
                data = Form1.dirList.List(i) & vbCrLf
                Print #1, data
            End If
        Next i
    Close #1
End Sub
