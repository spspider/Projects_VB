VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   360
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
v2 = v2 + 2
For a = 1 To 30
op1T = Cos(((Cos(v1 / 40) / Cos(a / 5)) - a) / 10) * 3000 + Form1.Height / 3 + Sin(v1 / 20) * 100
op1L = Cos(((Sin(v1 / 40) / Sin(a / 5)) - a) / 10) * 3000 + Form1.Height / 3 + Cos(v1 / 80) * 1000
Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 4, a * 8), B
op2T = 2 * op1T
op2L = 2 * op1L
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
