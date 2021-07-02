VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Op4 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op3 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Long
Dim h1 As Long
Dim h2 As Long
Dim w As Long
Dim i As Long
Dim k As Currency
Dim a As Long
Dim er As Long
Private Sub Timer1_Timer()
If h1 + 30 > Width Then
Timer2.Enabled = True
a = a + 1
End If

If a > 5000 Then
a = 1
Timer2.Enabled = False
k = Math.Rnd * 100
h = 0
h1 = 0
End If

If k > 55 Then
'''''''''''''''''''''''''''эффект шторки'''''
Timer3.Enabled = True
Timer4.Enabled = False
Else
'''''''''''''''''''''''''''эффект крест''''''
Timer4.Enabled = True
Timer3.Enabled = False
End If

End Sub

Private Sub Timer2_Timer() '''''''''''незнаю, надо разобраться'''''
er = a
If er > 2000 Then
er = 2000
End If

Op1.Top = Math.Sin(a / 26) * (1000 + er + Math.Rnd * 10) + 2900
Op1.Left = Math.Cos(a / 23) * (1000 + er + Math.Rnd * 10) + 2900

Op2.Top = Math.Sin(a / 25) * (1000 + er + Math.Rnd * 10) + 3900
Op2.Left = Math.Cos(a / 28) * (1000 + er + Math.Rnd * 10) + 3900

Op3.Top = Math.Sin(a / 30) * (1000 + er + Math.Rnd * 10) + 4900
Op3.Left = Math.Cos(a / 24) * (1000 + er + Math.Rnd * 10) + 4900

Op4.Top = Math.Sin(a / 25) * (1000 + Math.Rnd * 10) + 5700
Op4.Left = Math.Cos(a / (22)) * (1000 + Math.Rnd * 10) + 5600
''''''наверно координаты отрезков''''''
For i = 0 To 255 '''''''''''рисование линии и её изменение цвета в цикле(255) на 1 миллисекунду''''
Line (Op1.Left, Op1.Top)-(Op3.Left, Op3.Top), RGB(0, 255 - i, 355 - i)
Line (Op1.Left, Op1.Top)-(Op2.Left, Op2.Top), RGB(0, 255 - i, 355 - i)
Line (Op1.Left, Op1.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
Line (Op2.Left, Op2.Top)-(Op3.Left, Op3.Top), RGB(0, 255 - i, 355 - i)
Line (Op2.Left, Op2.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
Line (Op3.Left, Op3.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
Next
End Sub

Private Sub Timer3_Timer() '''''''''''''шторки
If h < Height Then
h = h + 15
Line (h1 - h, 0)-(Width, h), RGB(0, 0, h / 13)
End If
If h1 < Width Then
h1 = h1 + 15
Line (0, h - h1)-(h1, Height), RGB(0, 0, h1 / 13)
End If
End Sub

Private Sub Timer4_Timer() ''''''''''''крест
If h < Height Then
h = h + 15
Line (0, h)-(Width, h), RGB(0, 0, h / 13)
End If
If h1 < Width Then
h1 = h1 + 15
Line (h1, 0)-(h1, Height), RGB(0, 0, h1 / 13)
End If
End Sub
