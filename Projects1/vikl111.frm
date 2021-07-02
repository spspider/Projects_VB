VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DTLL"
   ClientHeight    =   2685
   ClientLeft      =   6525
   ClientTop       =   5475
   ClientWidth     =   3090
   Icon            =   "vikl111.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   3090
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1320
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "15:09:00"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim d As Integer

Private Sub Form_Load() '''''''''''''''''
Form1.Visible = True
On Error GoTo er1:
f = FreeFile
Open "Data\SP.xxx" For Input As f
a = Input(LOF(f), f)
Close #f
er1:
If a = 0 Then
Else
Timer2.Enabled = True
End If
End Sub
Private Sub Timer1_Timer()
Text1.Text = Time
If Text1.Text = Text2.Text Then
Shell ("C:\windows\system32\shutdown.exe -r"), vbHide
End If
End Sub

