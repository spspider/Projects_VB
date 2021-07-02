VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1560
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   360
      Picture         =   "Pic4.frx":0000
      Top             =   240
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xm As Long
Dim Ym As Long

Dim q1X As Long
Dim q1Y As Long
Dim q2X As Long
Dim q2Y As Long
Dim qX As Long
Dim qY As Long

Dim X As Long
Dim Y As Long
Dim a2 As Long
Dim q As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub Timer1_Timer()
Text1.Text = (Sqr(q2X * q2X + q1X * q1X) + Sqr(q2Y * q2Y + q1Y * q1Y)) / 2
Text1.Text = (((q2X + q2X) - (q1X + q1X)) / 2 + ((q2Y + q2Y) - (q1Y + q1Y)) / 2) / 2
''''''''''''''''''''''''''''''''''''''''''''''
qX = q2X - q1X
qY = q2Y - q1Y
''''''''''''''''''''''''''''''''''''''''''''''
Form1.Cls
a2 = a2 + 1
For q = 0 To 20
Y = Cos(-qY / 1000) * 1000 * q / 6 + 3000
X = Sin(-qX / 1000) * 1000 * q / 6 + 3000
PaintPicture Image1, X, Y
Next
For q = 0 To 20
Y = Sin(-qY / 1000) * 1000 * q / 6 + 3000
X = Cos(-qX / 1000) * 1000 * q / 6 + 6000
PaintPicture Image1, X, Y
Next
For q = 0 To 20
Y = Sin(-qY / 1000) * 1000 * q / 6 + 8000
X = Sin(-qX / 1000) * 1000 * q / 6 + 8000
PaintPicture Image1, X, Y
Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub

Private Sub Timer2_Timer()
q1X = Xm
q1Y = Ym
End Sub

Private Sub Timer3_Timer()
q2X = Xm
q2Y = Ym
End Sub


