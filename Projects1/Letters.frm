VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2055
   Icon            =   "Letters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   2055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Скрыть"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "стоп"
      Height          =   975
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "пуск"
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Dim x As Long
Const KEYEVENTF_KEYUP = &H2 'событие отпускания клавиши
Const VK_CONTROL = &H11 'клавиша Ctrl
Const VK_ESCAPE = &H1B 'клавиша Escape
'Const vbKeyCapital = &H1B 'клавиша Ctrl
'Эмулирующая нажатие кнопки ПУСК

Private Sub Command4_Click()
'Функция эмулирует нажатие Ctrl + Esc
Call keybd_event(VK_CONTROL, 0, 0, 0) 'Hажимаем Ctrl
Call keybd_event(VK_ESCAPE, 0, 0, 0) 'Hажимаем Esc
Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0) 'Отпускаем Esc
Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0) 'Отпускаем Ctrl
End Sub
Private Sub Command1_Click()
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End Sub
Private Sub Command2_Click()
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End Sub
Private Sub Command3_Click()
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End Sub


Private Sub Command6_Click()
Timer1.Enabled = False
End Sub
Private Sub Command5_Click()
Timer1.Enabled = True
End Sub
Private Sub Command7_Click()
Form1.Visible = False
End Sub

Private Sub Form_Load()
Form1.Visible = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
'''''''''''
If x = 10 Then
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 20 Then
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
''''''''''''''
If x = 30 Then
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 40 Then
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''
If x = 50 Then
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 60 Then
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If x = 12 Then
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 22 Then
Call keybd_event(vbKeyNumlock, 0, 0, 0)
Call keybd_event(vbKeyNumlock, 0, KEYEVENTF_KEYUP, 0)
End If
''''''''''''''
If x = 32 Then
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 42 Then
Call keybd_event(vbKeyCapital, 0, 0, 0)
Call keybd_event(vbKeyCapital, 0, KEYEVENTF_KEYUP, 0)
End If
'''''''''''''
If x = 52 Then
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If
If x = 62 Then
Call keybd_event(vbKeyScrollLock, 0, 0, 0)
Call keybd_event(vbKeyScrollLock, 0, KEYEVENTF_KEYUP, 0)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If x = 100 Then x = 0
End Sub
