VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   480
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4815
      Left            =   7800
      Max             =   4714
      Min             =   -1570
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Курс:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim course As Currency
Dim q As Boolean
Dim Xm, Ym As Integer
Private Sub Command1_Click()
course = ((360 * Rnd) / 56.9 - 1.57)
Text2.Text = course * 1000
VScroll1.Value = course * 1000
End Sub

Private Sub Form_Click()
If q = False Then
q = True
Else
q = False
End If
End Sub

Private Sub Form_Load()
Form1.AutoRedraw = True
End Sub

Private Sub Timer1_Timer()
Text1.Text = Cos(VScroll1.Value / 100)
Text3.Text = Sin(VScroll1.Value / 1000 * 56.9 + 90)
Text4.Text = Sin((Xm - Form1.Width / 2) / (Ym - Form1.Height / 2))
'Text5.Text = Math.Atn(0.23)
Form1.Cls
Line (Form1.Width / 2, Form1.Height / 2)-(Form1.Width / 2 + Cos(VScroll1.Value / 1000) * 3000, Form1.Height / 2 + Sin(VScroll1.Value / 1000) * 3000), RGB(0, 255, 100)
If q = True Then
Line (Form1.Width / 2, Form1.Height / 2)-(Xm, Ym), RGB(0, 255, 100)
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
