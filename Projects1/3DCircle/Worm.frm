VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   2130
   ClientTop       =   1515
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   240
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
'Form1.BackColor = RGB(216, 228, 248)
Form1.BackColor = RGB(0, 0, 0)
On Error GoTo er:
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a = 1 To 100
op1T = Sin(v1 / 20 + a / 4) * 100 + Sin(v1 / 10 + a / 10) * 500 + Form1.Height / 2
op1L = Cos(-v1 / 10 + a / 10) * 100 + Cos(v1 / 50 + a / 30) * 3000 + Form1.Width / 2
'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
Circle (op1L, op1T), Abs(Sin(a / 33)) * 150, RGB(a * 2, a * 2, 255 - a)
op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
Next
If v1 > 20000 Then
v1 = 0
End If
End Sub



