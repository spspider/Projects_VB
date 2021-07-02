VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Любименькой Машеньке от Серёжи:)"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "Heart2lov.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   12690
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      Picture         =   "Heart2lov.frx":5A4A
      ScaleHeight     =   1335
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   7560
      Width           =   4575
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   480
      Picture         =   "Heart2lov.frx":EA05
      ScaleHeight     =   1695
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7680
      Picture         =   "Heart2lov.frx":15623
      ScaleHeight     =   855
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Long
Dim a As Integer

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

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Click()
Unload Me
End Sub
Private Sub Form_Load()
'Call mciExecute("play data\menu.mp3")

Form1.BackColor = RGB(216, 228, 248)
Picture1.BackColor = RGB(216, 228, 248)
Picture2.BackColor = RGB(216, 228, 248)
End Sub

Private Sub Form_Resize()
Picture1.Left = Width - (Picture1.Width + 300)
Picture1.Top = Height / 2
Picture3.Left = 0
Picture3.Top = Height - (Picture3.Height + 500)
End Sub

Private Sub Timer1_Timer()
Form1.Cls
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
v1 = v1 + 1
For a = 1 To 63
op1T = Cos(625 / 90 + a / 20) * 2500 * Abs(Sin(v1 / 60 + 625 / 20)) + Form1.Height / 2
op1L = Abs(Sin(625 / 100 + a / 20)) * Abs(Sin(v1 / 60 + 625 / 20)) * 2500 * Abs(Cos(v1 / 20)) + Form1.Width / 2

op2T = Cos(625 / 90 + a / 20) * 2500 * Abs(Sin(v1 / 60 + 625 / 20)) + Form1.Height / 2
op2L = Abs(Sin(625 / 100 + a / 20)) * Abs(Sin(v1 / 60 + 625 / 20)) * -1 * 2500 * Abs(Cos(v1 / 20)) + Form1.Width / 2

'line (op1L, op1T)-(op2L, op2T), RGB(255, 0, 0)
Circle (op1L, op1T), (Abs(Sin(a / 20) * 80)), RGB(255, 0, 0)
Circle (op2L, op2T), (Abs(Sin(a / 20) * 80)), RGB(255, 0, 0)
Next
If v2 > 20000 Then
v2 = 0
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For a = 1 To 100
op1T1 = Cos(v1 / 40 + a / 35) * 4000 + Form1.Height / 2
op1L1 = Cos(v1 / 40 + a / 20) * 4000 + Form1.Width / 2

op2T1 = Cos(v1 / 40 + a / 20) * 4000 + Form1.Height / 2
op2L1 = Cos(v1 / 40 + a / 35) * 4000 + Form1.Width / 2

'op1T1 = Cos(v1 / 40 + a / 35) * 4000 + Form1.Height / 2
'op1L1 = Cos(v1 / 40 + a / 20) * 4000 + Form1.Width / 2

'op2T1 = Cos(v1 / 40 + a / 20) * 4000 + Form1.Height / 2
'op2L1 = Cos(v1 / 40 + a / 35) * 4000 + Form1.Width / 2

'Line (op1L, op1T)-(op2L, op2T), RGB(a * 2, a * 2, 255 - a)
'Circle (op1L1, op1T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)
'Circle (op2L1, op2T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)

Circle (op1L1, op1T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)
Circle (op2L1, op2T1), a * 2, RGB(a * 2, 255 - a * 2, a * 3)

Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Picture2.Top = Cos(v1 / 30) * 500 + Height / 2 - 1500
Picture1.Top = Sin(v1 / 30) * 500 + Height / 2 + 1000
Picture3.Left = Sin(v1 / 30) * 500 + 300
End Sub



