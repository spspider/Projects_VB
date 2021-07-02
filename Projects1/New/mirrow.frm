VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   10500
   ClientTop       =   7815
   ClientWidth     =   4620
   Icon            =   "mirrow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "1000"
      RemotePort      =   1000
      LocalPort       =   1000
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1200
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Long
Dim y As Long
Dim der As Variant
Dim d1 As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Sub Form_Load()
'Call mciExecute("play data/music.mp3")
'Winsock1.RemoteHost = "192.168.1.36" ' указываем адресс сервера (здесь ты указал свой же адресс)
'Winsock1.RemotePort = 1000 ' указываем порт сервера (в нашем случае 1000)
'Winsock1.Connect ' вызываем коннект
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
StretchBlt Picture1.Picture, 0, 0, Form1.Width / 15, Form1.Height / 15, Form1.Picture1.hdc, -Form1.Left / 15, -Form1.Top / 15, 1040, 822, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.
End Sub

Private Sub Timer1_Timer()
Picture1.Width = Form1.Width
Picture1.Height = Form1.Height
y = y + 1
'Label1.Left = Sin(y / 3) * 1000 + 1800
Form1.Cls
'StretchBlt Form1.hdc, 0, 0, Form1.Width / 15, Form1.Height / 15, Form1.Picture1.hdc, -Form1.Left / 15, -Form1.Top / 15, Form1.Width / 15, Form1.Height / 15, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.
StretchBlt Form1.hdc, 0, 0, Form1.Width / 15, Form1.Height / 15, Form1.Picture1.hdc, -Form1.Left / 15, -Form1.Top / 15, 1040, 822, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.

'StretchBlt Form1.hdc, -200, -200, Form1.Width / 15 + 200, Form1.Height / 15 + 200, Form1.Picture1.hdc, -100, -100, 100, 100, SRCCOPY 'wwww.hdc - куда копировать (auto.x,auto.y) координаты тоесть где будет находиться объект;zzzzz.hdc - откуда копировать четыре цифры(0,0,23,31) означают координаты объекта и длину,ширину.

'Winsock1.SendData "qwer"
'Winsock1.SendData (Picture2.Picture)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
''''''''''''''''''''''''''''''''''


