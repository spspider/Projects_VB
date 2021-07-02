VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Op1 
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   360
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4800
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Text            =   "12"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   30000
      TabIndex        =   10
      Top             =   4680
      Width           =   5535
   End
   Begin VB.VScrollBar VS1 
      Height          =   3615
      Left            =   6240
      Max             =   3600
      TabIndex        =   9
      Top             =   840
      Value           =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Text            =   "1800"
      Top             =   4560
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Line DD1 
      X1              =   3360
      X2              =   4320
      Y1              =   3720
      Y2              =   3240
   End
   Begin VB.Line CD 
      X1              =   3360
      X2              =   3360
      Y1              =   2640
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Caption         =   "C"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "D"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   135
   End
   Begin VB.Line CC1 
      X1              =   3360
      X2              =   4320
      Y1              =   2640
      Y2              =   2160
   End
   Begin VB.Line BC 
      X1              =   1200
      X2              =   3360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line AD 
      X1              =   1200
      X2              =   3360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line A1B1 
      BorderStyle     =   3  'Dot
      X1              =   2280
      X2              =   2280
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line C1D1 
      X1              =   4320
      X2              =   4320
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line A1D1 
      BorderStyle     =   3  'Dot
      X1              =   2280
      X2              =   4320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line B1C1 
      X1              =   2280
      X2              =   4320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line AB 
      DrawMode        =   1  'Blackness
      X1              =   1200
      X2              =   1200
      Y1              =   2640
      Y2              =   3720
   End
   Begin VB.Line AA1 
      BorderStyle     =   3  'Dot
      X1              =   1200
      X2              =   2280
      Y1              =   3720
      Y2              =   3240
   End
   Begin VB.Line BB1 
      X1              =   1200
      X2              =   2280
      Y1              =   2640
      Y2              =   2160
   End
   Begin VB.Label Label8 
      Caption         =   "D1"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "C1"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "B1"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "A1"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Timer1_Timer()
If AB.X1 <> BC.X1 Or AB.Y1 <> BC.Y1 Then
AB.X1 = BC.X1
AB.Y1 = BC.Y1
End If
If AB.X1 <> BB1.X1 Or AB.Y1 <> BB1.Y1 Then
AB.X1 = BB1.X1
AB.Y1 = BB1.Y1
End If
If AB.X2 <> AD.X1 Or AB.Y2 <> AD.Y1 Then
AB.X2 = AD.X1
AB.Y2 = AD.Y1
End If
If AB.X2 <> AA1.X1 Or AB.Y2 <> AA1.Y1 Then
AB.X2 = AA1.X1
AB.Y2 = AA1.Y1
End If

AD.X1 = AB.X2
AD.Y1 = AB.Y2
AD.X1 = AA1.X1
AD.Y1 = AA1.Y1
AD.X2 = CD.X2
AD.Y2 = CD.Y2
AD.X2 = DD1.X1
AD.Y2 = DD1.Y1

BC.X1 = AB.X1
BC.Y1 = AB.Y1
BC.X2 = CD.X2
BC.Y2 = CD.Y2
BC.X1 = BB1.X1
BC.Y1 = BB1.Y1
BC.X2 = CC1.X1
BC.Y2 = CC1.Y1

CD.X1 = BC.X1
CD.Y1 = BC.Y1
CD.X1 = CC1.X1
CD.Y1 = CC1.Y1
CD.X2 = AD.X2
CD.Y2 = AD.Y2
CD.X2 = DD1.X1
CD.Y2 = DD1.Y1

A1B1.X1 = BB1.X2
A1B1.Y1 = BB1.Y2
A1B1.X1 = B1C1.X1
A1B1.Y1 = B1C1.Y1
A1B1.X2 = AA1.X2
A1B1.Y2 = AA1.Y2
A1B1.X2 = A1D1.X1
A1B1.Y2 = A1D1.Y1

B1C1.X1 = BB1.X1
B1C1.Y1 = BB1.Y1
B1C1.X1 = A1B1.X1
B1C1.Y1 = A1B1.Y1
B1C1.X2 = CC1.X2
B1C1.Y2 = CC1.Y2
B1C1.X2 = C1D1.X1
B1C1.Y2 = C1D1.Y1

C1D1.X1 = B1C1.X2
C1D1.Y1 = B1C1.Y2
C1D1.X1 = CC1.X2
C1D1.Y1 = CC1.Y2
C1D1.X2 = A1D1.X2
C1D1.Y2 = A1D1.Y2
C1D1.X2 = DD1.X2
C1D1.Y2 = DD1.Y2

A1D1.X1 = AA1.X2
A1D1.Y1 = AA1.Y2
A1D1.X1 = A1B1.X2
A1D1.Y1 = A1B1.Y2
A1D1.X2 = DD1.X2
A1D1.Y2 = DD1.Y2
A1D1.X2 = C1D1.X2
A1D1.Y2 = C1D1.Y2

AA1.X1 = AD.X1
AA1.Y1 = AD.Y1
AA1.X1 = AB.X2
AA1.Y1 = AB.Y2
AA1.X2 = A1B1.X2
AA1.Y2 = A1B1.Y2
AA1.X2 = A1D1.X1
AA1.Y2 = A1D1.Y1

BB1.X1 = AB.X1
BB1.Y1 = AB.Y1
BB1.X1 = BC.X1
BB1.Y1 = BC.Y1
BB1.X2 = B1C1.X2
BB1.Y2 = B1C1.Y2
BB1.X2 = A1B1.X1
BB1.Y2 = A1B1.Y1

CC1.X1 = BC.X2
CC1.Y1 = BC.Y2
CC1.X1 = CD.X1
CC1.Y1 = CD.Y1
CC1.X2 = B1C1.X2
CC1.Y2 = B1C1.Y2
CC1.X2 = C1D1.X1
CC1.Y2 = C1D1.Y1

AA1.X1 = AD.X1
AA1.Y1 = AD.Y1
AA1.X1 = AB.X2
AA1.Y1 = AB.Y2
AA1.X2 = A1D1.X1
AA1.Y2 = A1D1.Y1
AA1.X2 = A1B1.X2
AA1.Y2 = A1B1.Y2

DD1.X1 = AD.X2
DD1.Y1 = AD.Y2
DD1.X1 = CD.X2
DD1.Y1 = CD.Y2
DD1.X2 = A1D1.X2
DD1.Y2 = A1D1.Y2
DD1.X2 = C1D1.X2
DD1.Y2 = C1D1.Y2



Text1.Text = VS1.Value
BB1.X1 = VS1.Value
CC1.X1 = 5000 - VS1.Value

a = a + 1
Text2.Text = a

Op1.Left = Math.Sin(Text2.Text / 30) * 400 + 1000
Op1.Top = Math.Cos(Text2.Text / 30) * 400 + 1000
End Sub

