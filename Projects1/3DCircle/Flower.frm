VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
'Form1.BackColor = RGB(216, 228, 248)
Form1.BackColor = RGB(0, 0, 0)
On Error GoTo er:
'Call mciExecute("play data\menu.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a = 1 To 100
op1T = Sin(v1 / 50 + a / 20) * Sin(v1 / 50 + a / 11) * 5000 + Form1.Height / 2
op1L = Cos(v1 / 50 + a / 20) * Sin(v1 / 50 + a / 11) * 5000 + Form1.Width / 2

'Line (op1L, op1T)-(op2L, op2T), RGB(255 - a * 2, 255 - a * 2, 255 - a)
Circle (op1L, op1T), Abs(Sin(a / 33)) * 150, RGB(a * 2, a * 2, 255 - a)
op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
Next

If v1 > 20000 Then
v1 = 0
End If
End Sub






