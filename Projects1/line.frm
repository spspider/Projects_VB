VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   120
      Max             =   32000
      TabIndex        =   4
      Top             =   6360
      Width           =   7695
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   6255
      Left            =   6840
      Max             =   3000
      Min             =   -3000
      TabIndex        =   3
      Top             =   120
      Value           =   1000
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   120
      Max             =   32000
      SmallChange     =   100
      TabIndex        =   2
      Top             =   6840
      Value           =   3000
      Width           =   7695
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6255
      Left            =   7320
      Max             =   3000
      Min             =   -3000
      TabIndex        =   1
      Top             =   120
      Value           =   1000
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   5520
   End
   Begin VB.OptionButton Op1 
      Caption         =   "Op1"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   5280
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   2160
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C000C0&
      X1              =   960
      X2              =   3840
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   2160
      Y1              =   1800
      Y2              =   4080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Private Sub Timer1_Timer()

X1 = 1500
X2 = 1500

a = a + 1

Op1.Top = Math.Sin(HScroll1.Value / 1500) * VScroll2.Value + X2
Op1.Left = Math.Cos(HScroll1.Value / 1500) * VScroll1.Value + X1

Line1.X1 = Op1.Left
Line1.Y1 = Op1.Top
Line1.X2 = 2 * X1 - Op1.Left
Line1.Y2 = 2 * X2 - Op1.Top

Line2.X1 = Line1.X1
Line2.Y1 = Line1.Y1 + X1 * 2
Line2.X2 = Line1.X2
Line2.Y2 = Line1.Y2 + X2 * 2

Line3.X1 = Line1.X1
Line3.Y1 = Line1.Y1
Line3.X2 = Line2.X1
Line3.Y2 = Line2.Y1

Line4.X1 = Line1.X2
Line4.Y1 = Line1.Y2
Line4.X2 = Line2.X2
Line4.Y2 = Line2.Y2


End Sub

