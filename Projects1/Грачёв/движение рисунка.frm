VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   2640
   ClientTop       =   2670
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   7560
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   0
      Picture         =   "движение рисунка.frx":0000
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   -600
      Visible         =   0   'False
      Width           =   12060
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Form1.Cls
StretchBlt Form1.hdc, auto.x, auto.y, 100, 50, Form1.Picture1.hdc, 300, 100, 100, 50, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.
auto.x = auto.x + 1
auto.y = auto.y + 1
If auto.x >= 200 Then
auto.x = 20
End If
If auto.y >= 198 Then
auto.y = 25
End If

End Sub
