VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Goto Folder "
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   2880
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete List"
      Height          =   615
      Left            =   2040
      Picture         =   "Form1.frx":4A97
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse Folder"
      Height          =   615
      Left            =   480
      Picture         =   "Form1.frx":A279
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox dirList 
      Height          =   1230
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Folder"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3600
      Picture         =   "Form1.frx":A6BB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":AAFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   555
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


Private Const SW_SHOWNORMAL As Long = 1


Private Sub OpenBrowseForFolder()
Dim lpIDList As Long
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
  szTitle = "Choose destination folder/directory..."
  With tBrowseInfo
     .hWndOwner = Me.hwnd
     .lpszTitle = lstrcat(szTitle, "")
     .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
     strTujuan = Space(MAX_PATH)
     SHGetPathFromIDList lpIDList, strTujuan
         strTujuan = Left(strTujuan, InStr(strTujuan, vbNullChar) - 1)
         dirList.AddItem (strTujuan)
         Call savefiles
  End If
End Sub
Private Sub Command1_Click()
 ShellExecute Me.hwnd, "Open", Text1.Text, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub


Private Sub Command2_Click()
OpenBrowseForFolder
End Sub


Private Sub Command3_Click()
 Dim index As Integer
    index = dirList.ListIndex
    On Error Resume Next
    
    Form1.dirList.RemoveItem (index)
    
    Call savefiles
    
End Sub


Private Sub Command4_Click()

End Sub

Private Sub dirList_Click()
dirfolder = dirList.Text
Text1.Text = dirfolder
Text2.Text = dirList.ListIndex
Command1.Enabled = True
End Sub


Private Sub Form_Load()
 dirF = windir$ & "\mcjons.txt"

  test = Dir(dirF)
    If test = "" Then
        Open dirF For Output As #1
        Close #1
    End If
 Call loadfiles
End Sub


Private Sub Label1_Click()
End
End Sub


Private Sub Label2_Click()
Form1.WindowState = 1
End Sub


