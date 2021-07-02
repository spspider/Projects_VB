VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   Icon            =   "N-G1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2280
      Picture         =   "N-G1.frx":5A4A
      ScaleHeight     =   3465
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   1920
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BEST OF THE BEST!!!!!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Long
Dim v2 As Long
Dim a As Integer
Dim step As Integer
Dim X As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer

Dim Xm As Integer
Dim Xn As Integer
Dim Ym As Integer

Dim op1T As Long
Dim op1L As Long
Dim op2T As Long
Dim op2L As Long
Dim q As Long
Dim op1T1 As Long
Dim op1L1 As Long

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Form1.BackColor = RGB(216, 228, 248)
Form1.BackColor = RGB(0, 0, 0)
On Error GoTo er:
Call mciExecute("play music\menu.mp3")
er:
End Sub

Private Sub Timer1_Timer()
Form1.Cls
q = q + 1
If q = 10 Then
Label1.Visible = True
End If
If q = 13 Then
Label1.Visible = False
End If
If q = 16 Then
Label1.Visible = True
End If
If q = 19 Then
Label1.Visible = False
End If

If q = 100 Then q = 0

'op1T = Sin(v1 / 20 + a / 4) * 100 + Sin(v1 / 10 + a / 10) * 500 + Form1.Height / 2
'op1L = Cos(-v1 / 10 + a / 10) * 100 + Cos(v1 / 50 + a / 30) * 3000 + Form1.Width / 2
'Circle (op1L, op1T), Abs(Sin(a / 33)) * 150, RGB(a * 2, a * 2, 255 - a)
'op2T = Form1.Height - op1T
'op2L = Form1.Width - op1L

v1 = v1 + 1
For a = 1 To 20 ''''''''''змейка
op1T1 = Sin(v1 / 20 + a / 10) * Form1.Height / 3 + Form1.Height / 2  'Y
op1L1 = Cos(a / 1000) + Sin(v1 / 10 + a / 5) * 150 + 1200 'X
Circle (op1L1, op1T1), Abs(Sin(a / 20)) * 150, RGB(a * 12, a * 12, (Abs(Sin(v1 / 20)) * 255))
'''''''''''''''''''''''''''''''''''''''''''''''''
op1T2 = Sin(v1 / 20 + a / 10) * Form1.Height / 3 + Form1.Height / 2  'Y
op1L2 = Cos(a / 1000) + Sin(-v1 / 10 - a / 5) * 150 + Form1.Width - 1200 'X
Circle (op1L2, op1T2), Abs(Sin(a / 20)) * 150, RGB(a * 12, a * 12, (Abs(Sin(v1 / 20)) * 255))
'''''''''''''''''''''''''''WORM
op1T5 = Sin(v1 / 2 + a / 4) * 100 + Sin(v1 / 10 + a / 10) * 500 + Form1.Height - 1500
op1L5 = Cos(-v1 / 10 + a / 10) * 100 + Cos(v1 / 60 + a / 30) * 4000 + Form1.Width / 2
Circle (op1L5, op1T5), Abs(Sin(a / 7)) * 50, RGB(a * 12, a * 2, 255 - a)
'''''''''''''''''''''''''''''''
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For a = 1 To 30 '''''''''''взрыв
op1T = Sin(v1 / 50 + a * 51) * (Cos(v1 / 20)) * 2000 + Form1.Height / 5
op1L = Cos(v1 / 50 + a * 70) * (Cos(v1 / 20)) * 2000 + 1200
Circle (op1L, op1T), Abs(Sin(a * 100)) * 50, RGB(a * 12, Abs(Sin(v1 / 100)) * 254, 0)
''''''''''''''''''''''''''''''''
op1T3 = Sin(v1 / 50 + a * 51) * (Cos(v1 / 20)) * 2000 + Form1.Height / 5
op1L3 = Cos(v1 / 50 + a * 70) * (Cos(v1 / 20)) * 2000 + Form1.Width - 1200
Circle (op1L3, op1T3), Abs(Sin(a * 100)) * 50, RGB(a * 12, Abs(Sin(v1 / 100)) * 254, 0)
Next

For a = 1 To 50
op1T4 = Tan(((Xm / 1000) + a) / 10) * 90000 + Form1.Height / 2
'op1T2 = Tan(((v12 / 10) + a) / 10) * 5000 + Form1.Height / 2
op1L4 = -100000
Line (op1L4, op1T4)-(op2L4, op2T4), RGB(0, a * 3, 0)
op2T4 = Height - op1T4
op2L4 = Width - op1L4
Next

For a = 1 To 100
op1T6 = Sin(Ym / 2000 + a / 20 + v1 / 100) * Sin(Ym / 2000 + a / 11) * 1400 + Form1.Height / 5
op1L6 = Cos(Xm / 2000 + a / 20 + v1 / 100) * Sin(Xm / 2000 + a / 11) * 1400 + Form1.Width / 2

'Line (op1L, op1T)-(op2L, op2T), RGB(255 - a * 2, 255 - a * 2, 255 - a)
Circle (op1L6, op1T6), Abs(Sin(a / 33)) * 150, RGB(a * 2, a * 2, 255 - a)
Next
'''''''''''''''''''''''
If v1 > 2000 Then
v1 = 0
End If
Picture1.Left = Form1.Width / 2 - Picture1.Width / 2
Picture1.Top = Form1.Height / 2 - Picture1.Height / 2
Label1.Left = Form1.Width / 2 - Label1.Width / 2
Label1.Top = Form1.Height - 3000
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub



