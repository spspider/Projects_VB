VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3075
   ClientTop       =   2595
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   1440
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
op1T = Sin(a / 10 + v1 / 10) * 1500 + Sin(v1 / -30) * 2000 + 4000
op1L = Sin(a / 10 + v1 / 10) * 1500 + Cos(v1 / -30) * 2000 + 4000
Line (op1L, op1T)-(op2L, op2T), RGB(a, a * 8, a * 8), B
op2T = 2 * op1T
op2L = 2 * op1L
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
