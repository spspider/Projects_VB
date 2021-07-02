VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
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
On Error GoTo er:
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a = 1 To 126
op1T = Sin(v1 / 100 + a / 10) * 3000 + Sin(v1 / 20 + a / 20) * 3000 + Form1.Height / 2
op1L = Cos(v1 / 100 + a / 10) * 3000 + Cos(v1 / 20 + a / 20) * 3000 + Form1.Width / 2
'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
Circle (op1L, op1T), 50, RGB(Abs(Sin(a / 5)) * 200, Abs(Sin(a / 5)) * 200, Abs(Sin(a / 5)) * 200)
Circle (op2L, op2T), 50, RGB(Abs(Sin(a / 5)) * 200, Abs(Sin(a / 5)) * 200, Abs(Sin(a / 5)) * 200)

op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
Next
If v1 > 20000 Then
v1 = 0
End If
End Sub


