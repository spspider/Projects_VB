VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Enum ButtonClick
btcLeft
btcRight
btcMiddle
End Enum

Private Function MouseClick(ByVal MBClick As ButtonClick) As Boolean
Dim cbuttons As Long
Dim dwExtraInfo As Long
Dim mevent As Long

Select Case MBClick
Case ButtonLeft
mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP
Case ButtonRight
mevent = MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP
Case ButtonMiddle
mevent = MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP
Case Else
MouseClick = False
Exit Function
End Select

mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo
MouseClick = True
End Function

Private Sub Timer1_Timer()
Call MouseClick(ButtonLeft)
End Sub
