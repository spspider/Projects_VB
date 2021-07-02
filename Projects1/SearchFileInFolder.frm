VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Search"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "SearchFileInFolder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d1 As String
Dim n As Long
Private Sub File1_Click()
Form1.Caption = "Search  " + File1.FileName
End Sub

Private Sub File1_DblClick()
On Error GoTo er:
Shell (File1.FileName), vbNormalFocus
er:
End Sub

Private Sub Form_Load()
d1 = FreeFile
Open ("C:\clickME.txt") For Append As d1
On Error GoTo er:
For n = 1 To File1.ListCount
'n = n + 1
File1.ListIndex = n - 1
    If File1.FileName <> (App.EXEName & ".exe") Then
    Print #d1, n; "  " & File1.FileName
    End If
Next
Close #d1
er:
End Sub

