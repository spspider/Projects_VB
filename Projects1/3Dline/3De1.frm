VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   3075
   ClientTop       =   2595
   ClientWidth     =   10680
   Icon            =   "3De1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "123"
      Top             =   7080
      Width           =   10695
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public b As String
Dim tex As String
Dim v1 As Long
Dim v As Integer
Dim v2 As Long
Dim v13 As Long
Dim a As Integer
Dim a2 As Integer
Dim a3 As Long
Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim Xm As Long
Dim Ym As Long
Dim d As Integer

Dim op1T As Long
Dim op1L As Long
Dim op2T As Long
Dim op2L As Long
Dim op1T1 As Long
Dim op1L1 As Long
Dim op2T1 As Long
Dim op2L1 As Long
Dim op1T3 As Long
Dim op1L3 As Long
Dim op2T3 As Long
Dim op2L3 As Long
Dim op1T2 As Long
Dim op1L2 As Long
Dim op2T2 As Long
Dim op2L2 As Long

Dim TxT1 As String
Dim TxT2 As String
Dim TxT3 As String
Dim TxT4 As String
Dim TxTEnd As String

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 'событие отпускани€ клавиши
Const VK_CONTROL = &H11 'клавиша Ctrl
Const VK_ESCAPE = &H1B 'клавиша Escape
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
TxTEnd = "by Sergey Paukov"
n
Call mciExecute("play data\menu.mp3")
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload (b)
End Sub

Private Sub Timer1_Timer()
'Call SetFormPosition(Me.hWnd, True)
Form1.Cls
v1 = v1 + 1
v2 = v2 + 2
For a = 1 To 30
op1T = Sin(a / 10 + Ym / 5000) * Sin(v1 / 23) * 2000 + Form1.Height / 2 + Sin(v1 / 60) * 1000
op1L = Tan(a / 10 + Xm / 5000) * 2000 * Sin(v1 / 15) + Form1.Height / 2 + Cos(v1 / 60) * 1000

Circle (op1L, op1T), 100 * Sin(a / 10), RGB(a * 8, a * 8, a * 8)
Circle (op2L, op2T), 100 * Sin(a / 10), RGB(a * 8, a * 8, a * 8)

Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 4, a * 8), B
'Line (op1L, op1T)-(op2L, op2T), RGB(a * 1, a * 2, a * 4) ', B

op2T = Form1.Height - op1T
op2L = Form1.Width - op1L
'op2T = op1T
'op2L = op1L
Next

v13 = v13 + 2
For a3 = 1 To 40
op1T3 = Sin(((-v13 / 20) - a3 * 400)) * 1000 + Height - 2100
op1L3 = Cos(((-v13 / 20) - a3 * 400)) * 1000 + 2100
Line (op1L3, op1T3)-(op2L3, op2T3), RGB(a3 * 6, a3 * 3, a3 * 6), B
'Circle (op1L, op1T), RGB(15, 5, 0), RGB(255, 255 - a1 * 10, 255)
op2T3 = Height - 2100 '- op1T3
op2L3 = 2100 '- op1L3
Next
x = x + 2

If x = 10 Then n
If x = 20 Then n

If x = 40 Then c
If x = 50 Then c

If x = 10 Then s
If x = 20 Then s

If x = 60 Then x = 0

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Xm = x
Ym = Y
End Sub


Private Sub Timer2_Timer()
''''''''''''''''''''''''''''''''''''''
v = Len(Text1.Text) - 1
Text1.Text = Mid((Text1.Text), 2, v)
If Text1.Text = "" Then
For l = 0 To 230
Text1.Text = Text1.Text + " "
Next
Text1.Text = Text1.Text + TxTEnd
End If
''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub n()
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End Sub
Private Sub c()
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub s()
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End Sub
