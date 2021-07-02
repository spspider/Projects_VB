VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
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
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Dim x As Long
Const KEYEVENTF_KEYUP = &H2 'событие отпускани€ клавиши
Const VK_CONTROL = &H11 'клавиша Ctrl
Const VK_ESCAPE = &H1B 'клавиша Escape

'Const vbKeyCapital = &H1B 'клавиша Ctrl
'Ёмулирующа€ нажатие кнопки ѕ”— 
Private Sub Timer1_Timer()
x = x + 2

If x = 10 Then n
If x = 20 Then n

If x = 40 Then c
If x = 50 Then c

If x = 10 Then s
If x = 20 Then s

If x = 60 Then x = 0
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

Private Sub Form_Load()
n
Form1.Visible = False
End Sub

