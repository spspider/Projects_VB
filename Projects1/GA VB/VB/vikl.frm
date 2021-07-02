VERSION 5.00
Object = "{59E060F0-4FD5-4C80-AF3C-D1B7E0ED65B2}#1.0#0"; "CompControls.ocx"
Begin VB.Form Form1 
   Caption         =   "DTLL"
   ClientHeight    =   2685
   ClientLeft      =   6525
   ClientTop       =   5475
   ClientWidth     =   3090
   Icon            =   "vikl.frx":0000
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
   Begin CompControler.CompControl CC1 
      Left            =   2280
      Top             =   120
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim d As Integer
Function ret()
On Error GoTo er2:
MkDir "data"
er2:
End Function
Private Sub Form_Load() '''''''''''''''''
ret
d = 10
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

Private Sub Text2_Click()
Text2.SetFocus
End Sub
Private Sub Timer1_Timer()
Text3.Text = a
Text1.Text = Time
If Text1.Text = Text2.Text Then
CC1.ShutDown
f = FreeFile
Open "Data\SP.xxx" For Append As f
Print #f, "1"
Close #f
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
d = d - 1
If d = 0 Then
CC1.ShutDown
Unload Me
d = 10
End If
End Sub
