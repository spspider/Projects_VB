VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   1200
   End
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

Dim h1 As Currency
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
h1 = 1
On Error GoTo er:
'Call mciExecute("play data\Junkie_XL_Underachivers.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a = 1 To 126
op1T = Sin(v1 / 100 + a / 10) * 1000 / 2 + Sin(v1 / 20 + a / 20) * 1500 / 2 + Form1.Height / 2
op1L = Cos(v1 / 100 + a / 20) * 1000 / 2 + Cos(v1 / 20 + a / 20) * 3000 / 2 + Form1.Width / 2
'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
Circle (op1L, op1T), 100, RGB(255, 100, 100)
op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
Next
If v1 > 20000 Then
v1 = 0
End If
End Sub


