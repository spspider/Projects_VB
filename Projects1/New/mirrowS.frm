VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   10440
   ClientTop       =   7485
   ClientWidth     =   4620
   Icon            =   "mirrowS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   720
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Long
'Dim y As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Form_Load()
Call mciExecute("play data/music.mp3")
Winsock1.RemoteHost = "192.168.1.36" ' указываем адресс сервера (здесь ты указал свой же адресс)
Winsock1.RemotePort = 1000 ' указываем порт сервера (в нашем случае 1000)
Winsock1.Connect ' вызываем коннект
End Sub

Private Sub Timer1_Timer()
'y = y + 1
'Label1.Left = Sin(y / 3) * 1000 + 1800
Form1.Cls
StretchBlt Form1.hdc, 0, 0, 300, 300, Form1.Picture1.hdc, -739, -521, 1000, 1000, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.
Winsock1.SendData Form1.hdc
End Sub
