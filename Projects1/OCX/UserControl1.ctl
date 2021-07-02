VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00000000&
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   ScaleHeight     =   6600
   ScaleWidth      =   9855
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Dim v12 As Integer
Dim Xm As Integer
Dim Ym As Integer
Dim a2 As Integer
Dim op1L2 As Long
Dim op1T2 As Long
Dim op2L2 As Long
Dim op2T2 As Long

Private Sub Timer1_Timer()
UserControl.Cls
v12 = v12 - 1
For a2 = 1 To 50
op1T2 = Tan(((v12 / 10) + a2) / 10) * 5000 + Height / 2
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

