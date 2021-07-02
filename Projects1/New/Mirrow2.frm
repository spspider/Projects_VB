VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   9705
   ClientTop       =   8070
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5475
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenW As Integer
Dim ScreenH As Integer
Dim W As Currency
Dim H As Currency


Private Sub Timer1_Timer()
On Error GoTo er:
ScreenW = 800
ScreenH = 600
StretchBlt Form1.hdc, 0, 0, Form1.Width / 15, Form1.Height / 15, Form1.Picture1.hdc, -Form1.Left / 15 + 1024, -Form1.Top / 15, ScreenW, ScreenH, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.

W = ScreenW / (Form1.Width / 15)
H = ScreenH / (Form1.Height / 15)
Form1.Caption = Left(W, 4) & ":" & Left(H, 4)

Form1.Width = ScreenW * 15 / Left(W, 4)
Form1.Height = ScreenH * 15 / Left(W, 4)
er:
End Sub
