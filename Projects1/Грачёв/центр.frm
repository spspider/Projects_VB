VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Variant
Dim b As Variant
Dim X1 As Variant
Dim X2 As Variant
Dim Y1 As Variant
Dim Y2 As Variant
Dim R1 As Variant
Private Sub Timer1_Timer()
'Form1.Cls
a = a + 12
b = a + 1
Line (X1, Y1)-(X2, Y2), RGB(b + a * 2, 500, a + b * 2)
'''''''''''''''''''''''''''''''''''''''''''''

X1 = (Form1.Width \ 2) + b - a
X2 = 600 - a
Y1 = b + 20
Y2 = 100
Circle (X1, Y1), b + 1
Circle (X2, Y2), a + 10
''''''''''''''''''''''''''''''''''''''''''''''
Circle (Form1.Width \ 2, Form1.Height \ 2), 400 + a, RGB(10 + (a - 5), 0 + (a * 2) - b, b + 877)
If b >= 100 Then
b = 0
End If
If a >= 8000 Then
a = 0
End If
End Sub
