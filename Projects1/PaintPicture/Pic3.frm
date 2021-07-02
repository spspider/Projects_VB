VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   600
      Picture         =   "Pic3.frx":0000
      Top             =   1440
      Width           =   135
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
Dim a2 As Long
Dim q As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub Timer1_Timer()
Form1.Cls
a2 = a2 + 1
For q = 0 To 24
X = (Sin(Xm / 2000 + a2 / 40 + q / 4) * 1000 + Width / 2) * Abs(Sin(Sin(Xm / Width) * 10000 / Width)) * 7 / 2 - 1000
Y = (Cos(Ym / 2000 + a2 / 40 + q / 4) * 1000 + Height / 2) * Abs(Sin(Sin(Ym / 4000) * 10000 / Height)) * 2 - 1000 '* q / 20
PaintPicture Image1, X, Y
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
