VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Long
Dim v2 As Long
Dim a As Integer

Dim op1T As Long
Dim op1L As Long
Dim op2T As Long
Dim op2L As Long


Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
v1 = 400
On Error GoTo er:
'Call mciExecute("play data\menu.mid")
er:
End Sub

Private Sub Timer1_Timer()
'If v1 < 561 Then
'Timer3.Enabled = False
'Timer2.Enabled = True
'End If
'If v1 > 954 Then
'Timer2.Enabled = False
'Timer3.Enabled = True
'End If
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

