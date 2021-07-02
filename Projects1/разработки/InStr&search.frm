VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d1 As String
Dim q As String
Dim n As Long

Private Sub Command1_Click()
q = FreeFile
d1 = FreeFile
Open "c:/clickME.txt" For Input As f
Text1.Text = Input(LOF(q), q)
If InStr(n, Text1.Text, "F") <> 0 Then
Open ("C:\clickME_Full.txt") For Append As d1
Close q
On Error GoTo er:
For n = 1 To File1.ListCount
File1.ListIndex = n - 1
    If File1.FileName <> (App.EXEName & ".exe") Then
    Print #d1, n; "  " & File1.FileName
    End If
Next
Close #d1
er:
End If
End Sub

