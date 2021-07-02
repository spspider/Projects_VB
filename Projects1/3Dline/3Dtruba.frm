VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
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

Dim v12 As Integer
Dim Xm As Integer
Dim Ym As Integer

Dim a2 As Integer
Dim op1L2 As Long
Dim op1T2 As Long
Dim op2L2 As Long
Dim op2T2 As Long

Private Sub Timer1_Timer()
Form1.Cls
v12 = v12 - 1
For a2 = 1 To 50
op1T2 = Tan(((v12 / 10) + a2) / 10) * 5000 + Form1.Height / 2
op1L2 = -30000
Line (op1L2, op1T2)-(op2L2, op2T2), RGB(0, a2 * 4, 0)
op2T2 = Height - op1T2
op2L2 = Width - op1L2
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
