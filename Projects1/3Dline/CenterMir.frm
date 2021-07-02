VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
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
Dim v1 As Currency
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

Private Sub Form_Load()
On Error GoTo er:
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
If v1 > 1000 Then v1 = 1
Form1.Cls
v1 = v1 + 1 / 2
'v2 = v2 + 2
For a = 1 To 80
op1T = Sin(a + v1 / 80) * Sin(v1 / 100) * 5000 + Form1.Height / 2
op1L = Cos(a * 8 / 7 + v1 / 100) * Sin(v1 / 100) * 5000 + Form1.Width / 2
Line (op1L, op1T)-(op2L, op2T), RGB(a, a * 2, a * 2) ', B
'op2T = Form1.Height / 2
'op2L = Form1.Width / 2

op2T = op1T
op2L = op1L
Next
End Sub
