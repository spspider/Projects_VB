VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OCX"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   Icon            =   "OpenSP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar PB1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim d As String
Dim d1 As String
Dim d2 As String
Dim a As String
Dim t As String
Dim nF As String
Dim F As String
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
F = (Left(Environ("windir"), Len(Environ("windir")) - 7)) + "\Documents and Settings\All Users\Главное меню\Программы\SpSpider\SpReport.txt"
nF = (Left(Environ("windir"), Len(Environ("windir")) - 7)) + "\Documents and Settings\All Users\Главное меню\Программы\SpSpider"
On Error GoTo er:
File1.FileName = "Data\OCX\"
PB1.Max = (File1.ListCount - 1)
a = Environ("windir")
d = FreeFile
MkDir (nF)
Open (F) For Output As d
Close #d
er:
End Sub
Private Sub Timer1_Timer()
File1.ListIndex = n
On Error GoTo error:
FileCopy ("Data\OCX\" + File1.FileName), a + "\system32\" + File1.FileName
PB1.Value = n
Text1.Text = "Data\OCX\" + File1.FileName
Text2.Text = a + "\system32\" + File1.FileName
n = n + 1
error:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo error1:
       MkDir (nF)
error1:
    d2 = FreeFile
    Open (F) For Append As d2
    Print #d2, "Copy:   "; Text1.Text
    Print #d2, "To:   "; Text2.Text
    Print #d2, ""
    Close #d2

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If n > (File1.ListCount - 1) Then
    Timer1.Enabled = False
    'Shell ("sp.exe"), vbNormalFocus
    d1 = FreeFile
    Open (F) For Append As d1
    Print #d2, ""
    Print #d1, "Copying is completed"
    Print #d2, ""
    Print #d2, "in " & Time & "  " & Date
    Print #d2, ""
    Print #d2, ""
    Print #d2, ""
    Print #d2, ""
    Close #d1
    Unload Me
    End If
End Sub
