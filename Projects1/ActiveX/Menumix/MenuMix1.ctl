VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13470
   ScaleHeight     =   8400
   ScaleWidth      =   13470
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
Dim v1 As Long
Dim v2 As Long
Dim a As Integer
Dim step As Integer
Dim X As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer

Dim wer As Long

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

Dim a3 As Integer
Dim v13 As Long
Dim op1L3 As Long
Dim op1T3 As Long
Dim op2L3 As Long
Dim op2T3 As Long

Dim Xm As Long
Dim Ym As Long

Dim Nmus As Integer

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub UserControl_Click()
'Unload Me
End Sub

Private Sub UserControl_Load()
Nmus = 1
End Sub
Private Sub Timer1_Timer()
''''''''''''''''''''''''''''''''''''''''''''''''''''
Nmus = Nmus - 1
If Nmus = 0 Then
On Error GoTo er:
Call mciExecute("play data\menu.mp3")
Nmus = FileLen("Data\Menu.mp3") / 860
End If
er:
''''''''''''''''''''''''''''''''''''''''''''''''''''
If v1 > 30000 Then
v1 = 1
v11 = 1
v12 = 1
v13 = 1
End If
UserControl.Cls
v1 = v1 + 1
v2 = v2 + 2
For a = 1 To 50
op1T = Cos(((Cos(v1 / 40) / Cos(a / 5)) - a) / 10) * 3000 + Height / 3 + Sin(v1 / 20) * 100
op1L = Cos(((Sin(v1 / 40) / Sin(a / 5)) - a) / 10) * 3000 + Height / 3 + Cos(v1 / 80) * 1000
Line (op1L, op1T)-(op2L, op2T), RGB(a * 3 / 5, a * 6 / 5, a * 12 / 5), B
op2T = 2 * op1T
op2L = 2 * op1L
Next
'''''''''''''''''''''''''''''''''''''''
v11 = v11 + 1
For a1 = 1 To 30
op1T1 = -Width
op1L1 = Cos(((v11 / 20) + a1) / 2) * 3000 + Width / 4 - 1000
Line (op1L1, op1T1)-(op2L1, op2T1), RGB(a1 * 8, a1 * 4, a1 * 5)
op2T1 = Height - op1T1
op2L1 = (Width - Width / (15 / 10)) - op1L1 - 1000
Next
''''''''''''''''''''''''''''''''''''''''
v12 = v12 - 1
For a2 = 1 To 50
op1T2 = Tan(((v12 / 10) + a2) / 10) * 5000 + Height / 2 + 2000
op1L2 = -30000
Line (op1L2, op1T2)-(op2L2, op2T2), RGB(0, a2 * 3, 0)
op2T2 = (Height + 3000) - op1T2
op2L2 = Width - op1L2
Next
'''''''''''''''''''''''''''''''''''''''''
v13 = v13 + 1
For a3 = 1 To 150
op1T3 = Sin(((v13 / 4) + a3 / 20) / 2 + Xm / 3000) * 1000 + 1200
op1L3 = Cos(((v13 / 4) + a3 / 20) / 2 + Ym / 3000) * 1000 + Width - 1200
Line (op1L3, op1T3)-(op2L3, op2T3), RGB(a3 * 8 / 6, a3 * 8 / 6, 0) ', B
op2T3 = 1200 '- op1T3
op2L3 = Width - 1200 '- op1L3

Next
For wer = 0 To 10
Line (0, Ym + wer * 2)-(Width, Ym + wer * 2), RGB(Xm / 60, 0, 0) ', B
Line (Xm + wer * 2, 0)-(Xm + wer * 2, Height), RGB((Xm + Ym) / 80, 0, 0) ', B
Next
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub


