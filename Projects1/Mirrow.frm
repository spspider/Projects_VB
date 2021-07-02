VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll4 
      Height          =   2895
      Left            =   8880
      Max             =   500
      Min             =   -500
      TabIndex        =   4
      Top             =   2640
      Width           =   255
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   2895
      Left            =   9120
      Max             =   500
      Min             =   -500
      TabIndex        =   3
      Top             =   2640
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   2895
      Left            =   9600
      Max             =   1000
      Min             =   10
      TabIndex        =   2
      Top             =   2640
      Value           =   10
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   360
      Picture         =   "Mirrow.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
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
StretchBlt Form1.hdc, 0, 0, 300, 300, Form1.Picture1.hdc, VScroll4.Value, VScroll3.Value, VScroll2.Value, VScroll2.Value, SRCCOPY     'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.

End Sub


