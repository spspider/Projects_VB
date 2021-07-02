VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "\'/"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   Icon            =   "Heart2lov2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Integer
Dim v2 As Integer
Dim a As Integer

Dim op1T As Integer
Dim op1L As Integer
Dim op2T As Integer
Dim op2L As Integer

Dim op1T1 As Integer
Dim op1L1 As Integer
Dim op2T1 As Integer
Dim op2L1 As Integer


Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
Form1.BackColor = RGB(216, 228, 248)
v1 = 400
On Error GoTo er:
Call mciExecute("play data\menu.mid")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = 625
v2 = v2 + 1
For a = 1 To 63
op1T = Cos(v1 / 90 + a / 20) * 2000 * Abs(Sin(v2 / 30) + 1 / 8) + Form1.Height / 2
op1L = Abs(Sin(v1 / 100 + a / 20)) * Abs(Sin(v2 / 30) + 1 / 8) * 2000 * Abs(Sin(v2 / 12)) + Form1.Width / 2

op2T = Cos(v1 / 90 + a / 20) * 2000 * Abs(Sin(v2 / 30) + 1 / 8) + Form1.Height / 2
op2L = Abs(Sin(v1 / 100 + a / 20)) * Abs(Sin(v2 / 30) + 1 / 8) * -1 * 2000 * Abs(Sin(v2 / 12)) + Form1.Width / 2

Circle (op1L, op1T), (Abs(Sin(a / 20) * 80)), RGB(255, 100, 255)
Circle (op2L, op2T), (Abs(Sin(a / 20) * 80)), RGB(255, 100, 255)
Next
''''''''''''''''''''''''''''''''''''''''''
For a = 1 To 100
op1T1 = Cos(v2 / 30 + a / 35) * 4000 + Form1.Height / 2
op1L1 = Cos(v2 / 30 + a / 20) * 4000 + Form1.Width / 2

op2T1 = Cos(v2 / 30 + a / 20) * 4000 + Form1.Height / 2
op2L1 = Cos(v2 / 30 + a / 35) * 4000 + Form1.Width / 2

'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
Circle (op1L1, op1T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)
Circle (op2L1, op2T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)

Next
''''''''''''''''''''''''''''''''''''''''''
If v1 > 20000 Then
v1 = 0
End If
End Sub
Private Sub Timer2_Timer()
v1 = v1 + 1
End Sub
Private Sub Timer3_Timer()
v1 = v1 - 1
End Sub


