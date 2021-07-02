VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   120
   End
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
Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer

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
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
If v1 < 561 Then
Timer3.Enabled = False
Timer2.Enabled = True
End If
If v1 > 954 Then
Timer2.Enabled = False
Timer3.Enabled = True
End If
Form1.Cls

For a = 1 To 65
op1T = Cos(v1 / 90 + a / 20) * 2000 + Form1.Height / 2
op1L = Abs(Sin(v1 / 100 + a / 20)) * 2000 + Form1.Width / 2

op2T = Cos(v1 / 90 + a / 20) * 2000 + Form1.Height / 2
op2L = Abs(Sin(v1 / 100 + a / 20)) * -1 * 2000 + Form1.Width / 2

'Line (op1L, op1T)-(op1L * 750 / 500, op1T ), RGB(a * 2, 100, a * 2)
Circle (op1L, op1T), 100, RGB(255, 100, 255)
Circle (op2L, op2T), 100, RGB(255, 100, 255)

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
