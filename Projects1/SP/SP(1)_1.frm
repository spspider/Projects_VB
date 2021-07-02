VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   720
   End
   Begin VB.Line Line12 
      X1              =   4560
      X2              =   5040
      Y1              =   3600
      Y2              =   3120
   End
   Begin VB.Line Line11 
      X1              =   2640
      X2              =   3360
      Y1              =   3600
      Y2              =   3120
   End
   Begin VB.Line Line10 
      X1              =   4560
      X2              =   5160
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line9 
      X1              =   2760
      X2              =   3360
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line8 
      X1              =   4560
      X2              =   4560
      Y1              =   1800
      Y2              =   3600
   End
   Begin VB.Line Line7 
      X1              =   5160
      X2              =   5040
      Y1              =   1200
      Y2              =   3000
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   3360
      Y1              =   1200
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   2760
      X2              =   2640
      Y1              =   1800
      Y2              =   3480
   End
   Begin VB.Line Line4 
      X1              =   2760
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   5040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   4920
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   4440
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim e As Integer
Dim b As Integer


Private Sub Form_Load()
e = 900
b = 3000
End Sub

Private Sub Timer1_Timer()

a = a + 1

Line1.X1 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line1.Y1 = Math.Cos(a / 8) * 200 + (e + 1600)
Line1.X2 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line1.Y2 = Math.Cos(a / 10) * 100 + (e + 1600)

Line2.X1 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line2.Y1 = Math.Cos(a / 8) * 200 + (e + 1000)
Line2.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line2.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)

Line3.X1 = Math.Sin(a / 7) * 1500 + (b + 400)
Line3.Y1 = Math.Cos(a / 7) * 100 + e
Line3.X2 = Math.Sin(a / 7) * 1500 + (b + 1900)
Line3.Y2 = Math.Cos(a / 12) * 100 + e

Line4.X1 = Math.Sin(a / 7) * 1500 + b
Line4.Y1 = Math.Cos(a / 7) * 100 + (e + 600)
Line4.X2 = Math.Sin(a / 7) * 1500 + (b + 1500)
Line4.Y2 = Math.Cos(a / 12) * 100 + (e + 600)

Line5.X1 = Math.Sin(a / 7) * 1500 + b
Line5.Y1 = Math.Cos(a / 7) * 100 + (e + 600)
Line5.X2 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line5.Y2 = Math.Cos(a / 8) * 200 + (e + 1600)

Line6.X1 = Math.Sin(a / 7) * 1500 + (b + 400)
Line6.Y1 = Math.Cos(a / 7) * 100 + e
Line6.X2 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line6.Y2 = Math.Cos(a / 8) * 200 + (e + 1000)

Line7.X1 = Math.Sin(a / 7) * 1500 + (b + 1900)
Line7.Y1 = Math.Cos(a / 12) * 100 + e
Line7.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line7.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)

Line8.X1 = Math.Sin(a / 7) * 1500 + (b + 1500)
Line8.Y1 = Math.Cos(a / 12) * 100 + (e + 600)
Line8.X2 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line8.Y2 = Math.Cos(a / 10) * 100 + (e + 1600)

Line9.X1 = Math.Sin(a / 7) * 1500 + b
Line9.Y1 = Math.Cos(a / 7) * 100 + (e + 600)
Line9.X2 = Math.Sin(a / 7) * 1500 + (b + 400)
Line9.Y2 = Math.Cos(a / 7) * 100 + e

Line10.X1 = Math.Sin(a / 7) * 1500 + (b + 1500)
Line10.Y1 = Math.Cos(a / 12) * 100 + (e + 600)
Line10.X2 = Math.Sin(a / 7) * 1500 + (b + 1900)
Line10.Y2 = Math.Cos(a / 12) * 100 + e

Line11.X1 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line11.Y1 = Math.Cos(a / 8) * 200 + (e + 1600)
Line11.X2 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line11.Y2 = Math.Cos(a / 8) * 200 + (e + 1000)

Line12.X1 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line12.Y1 = Math.Cos(a / 10) * 100 + (e + 1600)
Line12.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line12.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)
End Sub

