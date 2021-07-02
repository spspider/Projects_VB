VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
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
On Error GoTo er:
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a = 1 To 100
op1T = Sin(v1 / 100 + a / 10) * 3000 + Sin(v1 / 20 + a / 20) * 3000 + Form1.Height / 2
op1L = Cos(v1 / 100 + a / 10) * 3000 + Cos(v1 / 30 + a / 30) * 3000 + Form1.Width / 2
'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
Circle (op1L, op1T), 100, RGB(a * 2, a * 2, 255 - a)
op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
Next
If v1 > 20000 Then
v1 = 0
End If
End Sub


