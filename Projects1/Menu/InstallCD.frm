VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "InstallCD.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "InstallCD.frx":5A4A
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   1560
      Top             =   960
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1200
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   840
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   960
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Visual Basic 6.0"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Projects"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "HELP"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "кодек"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8760
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   30
      Left            =   480
      Picture         =   "InstallCD.frx":5A754
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TxTEnd As String
Dim v As Integer
Dim v1 As Integer
Dim x As Long
Dim y As Long
Dim a2 As Long
Dim q As Long

Dim v2 As Integer
Dim v3 As Integer
Dim v4 As Integer
Dim v5 As Integer
 
Dim op1T As Long
Dim op1l As Long
Dim op2T As Long
Dim op2L As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 'событие отпускания клавиши
Const VK_CONTROL = &H11 'клавиша Ctrl
Const VK_ESCAPE = &H1B 'клавиша Escape
Private Sub Command4_Click()
Shell ("кодеки\codecs.exe"), vbNormalFocus
End Sub
Private Sub Command3_Click()
Shell ("help.exe"), vbNormalFocus
End Sub
Private Sub Command2_Click()
Shell ("projects.exe"), vbNormalFocus
End Sub
Private Sub Command1_Click()
Shell ("VB\Setup.exe"), vbNormalFocus
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call mciExecute("play data\music.mp3")
Command1.Left = -30000
Command2.Left = -30000
Command3.Left = -30000
Command4.Left = -30000
TxTEnd = "by Sergey Paukov"
n
End Sub

Private Sub Timer1_Timer()
v = Len(Text1.Text) - 1
Text1.Text = Mid((Text1.Text), 2, v)
If Text1.Text = "" Then
For l = 0 To 230
Text1.Text = Text1.Text + " "
Next
Text1.Text = Text1.Text + TxTEnd
End If

x = x + 1

If x = 1 Then n
If x = 2 Then n

If x = 1 Then c
If x = 2 Then c

If x = 1 Then s
If x = 2 Then s

If x = 3 Then x = 1

'Form1.Cls
'a2 = a2 + 1
'For q = 0 To 40
'x = Tan(a2 / 20 + q / 8) * 3000 + Width / 2
'y = Cos(a2 / 60 + q / 4) * 3000 + Height / 2 * q / 20
'PaintPicture Image1, x, y
'Next
'For a = 1 To 15
'op1T = -Form1.Width
'op1l = Sin(((-v1 / 10) + a) / 2) * 3000 + Form1.Height / 2
'Line (op1l, op1T)-(op2L, op2T), RGB(a * 5, a * 5, a * 5)
'op2T = Form1.Width / 2 - op1T
'op2L = Form1.Height / 2 - op1l
'Next
End Sub

Private Sub Timer2_Timer()
v2 = v2 + 1
Command1.Left = Sin(v2 / 50) * Form1.Width
If Command1.Left > 10200 Then
Timer2.Enabled = False
End If
End Sub
Private Sub Timer3_Timer()
v3 = v3 + 1
Command2.Left = Sin(v3 / 100) * Form1.Width
If Command2.Left > 10200 Then
Timer3.Enabled = False
End If
End Sub
Private Sub Timer4_Timer()
v4 = v4 + 1
Command3.Left = Sin(v4 / 150) * Form1.Width
If Command3.Left > 10200 Then
Timer4.Enabled = False
End If
End Sub
Private Sub Timer5_Timer()
v5 = v5 + 1
Command4.Left = Sin(v5 / 200) * Form1.Width
If Command4.Left > 10200 Then
Timer5.Enabled = False
End If
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
