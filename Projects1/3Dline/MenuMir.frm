VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
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

Dim op1T1 As Long
Dim op1L1 As Long
Dim op2T1 As Long
Dim op2L1 As Long
Dim a1 As Integer
Dim v11 As Integer

Dim a2 As Integer
Dim op1L2 As Long
Dim op1T2 As Long
Dim op2L2 As Long
Dim op2T2 As Long
Dim v12 As Long

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
v2 = v2 + 2
For a = 1 To 30
op1T = Cos(((Cos(v1 / 40) / Cos(a / 5)) - a) / 10) * 3000 + Form1.Height / 3 + Sin(v1 / 20) * 100
op1L = Cos(((Sin(v1 / 40) / Sin(a / 5)) - a) / 10) * 3000 + Form1.Height / 3 + Cos(v1 / 80) * 1000
Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 4, a * 8), B
op2T = 2 * op1T
op2L = 2 * op1L
Next

v11 = v11 + 1
For a1 = 1 To 30
op1T1 = -Form1.Width
op1L1 = Cos(((v11 / -20) + a1) / 2) * 3000 + Form1.Width / 4
Line (op1L1, op1T1)-(op2L1, op2T1), RGB(a1 * 8, a1 * 4, a1 * 5)
op2T1 = Form1.Height - op1T1
op2L1 = (Form1.Width - Form1.Width / (15 / 10)) - op1L1
Next
''''''''''''''''''''''''''''''''''''''''
v12 = v12 - 1
For a2 = 1 To 50
op1T2 = Tan(((v12 / 10) + a2) / 10) * 5000 + Form1.Height / 2
op1L2 = -30000
Line (op1L2, op1T2)-(op2L2, op2T2), RGB(0, a2 * 4, 0)
op2T2 = Height - op1T2
op2L2 = Width - op1L2
Next
'''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub

