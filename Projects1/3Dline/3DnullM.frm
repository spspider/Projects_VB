VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   1440
   End
   Begin VB.OptionButton Op2 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Op1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim v1 As Integer
Dim Xm As Integer
Dim Ym As Integer
Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 - 1
v2 = v2 + 2
For a = 1 To 45
Op1.Top = Tan(((v1 / 10) + a) / 10) * 5000 + Form1.Height / 2
Op1.Left = -(1000 + Xm * 2)
Line (Op1.Left, Op1.Top)-(Op2.Left, Op2.Top), RGB(0, 250, 100)
Op2.Top = 2 * Ym - Op1.Top
Op2.Left = Xm * 2 - Op1.Left
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
