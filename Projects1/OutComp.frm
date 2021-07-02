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
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   720
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Dim x As Long
Const KEYEVENTF_KEYUP = &H2
Const VK_CONTROL = &H11
Const VK_ESCAPE = &H1B

Private Sub Command1_Click()
Call keybd_event(VK_CONTROL, 0, 0, 0)
Call keybd_event(VK_ESCAPE, 0, 0, 0)
Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
Call keybd_event(vbKeyN, 0, 0, 0)
Call keybd_event(vbKeyN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
Text1.Text = KeyAscii
End Sub

