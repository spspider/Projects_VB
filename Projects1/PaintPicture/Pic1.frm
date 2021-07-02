VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   7650
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   600
      Picture         =   "Pic1.frx":0000
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xm As Long
Dim Ym As Long

Dim X As Long
Dim Y As Long
Dim a As Long
Dim q As Long
Private Sub Timer1_Timer()
Form1.Cls
a = a + 1
For q = 0 To 20
X = Sin(a / 10 + q / 20) * 1000 + Width / 2
Y = Cos(a / 20 + q / 1) * 1000 + Height / 2
PaintPicture Image1, Xm * q / 20, Y + (Ym - 6000)
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
