VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W"
   ClientHeight    =   135
   ClientLeft      =   3885
   ClientTop       =   4350
   ClientWidth     =   1110
   Icon            =   "CraizyKeyboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleWidth      =   1110
   Begin VB.PictureBox ImOn 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3360
      Picture         =   "CraizyKeyboard.frx":5A4A
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Num 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      Picture         =   "CraizyKeyboard.frx":5B8A
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox Scroll 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   960
      Picture         =   "CraizyKeyboard.frx":5CCA
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox Caps 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      Picture         =   "CraizyKeyboard.frx":5E0A
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox ImOff 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3120
      Picture         =   "CraizyKeyboard.frx":5F4A
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2520
      Top             =   1200
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
'''''''''''
If x = 10 Then
Num.Picture = ImOn.Picture
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 20 Then
Num.Picture = ImOff.Picture
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
''''''''''''''
If x = 30 Then
Caps.Picture = ImOn.Picture
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 40 Then
Caps.Picture = ImOff.Picture
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''
If x = 50 Then
Scroll.Picture = ImOn.Picture
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 60 Then
Scroll.Picture = ImOff.Picture
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If x = 14 Then
Num.Picture = ImOn.Picture
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 24 Then
Num.Picture = ImOff.Picture
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
''''''''''''''
If x = 34 Then
Caps.Picture = ImOn.Picture
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 44 Then
Caps.Picture = ImOff.Picture
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''
If x = 54 Then
Scroll.Picture = ImOn.Picture
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 64 Then
Scroll.Picture = ImOff.Picture
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If x = 100 Then x = 0
End Sub

Private Sub Form_Load()
Form1.Visible = False
End Sub
