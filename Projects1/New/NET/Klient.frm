VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   480
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "1000"
      RemotePort      =   1000
      LocalPort       =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Winsock1.LocalPort = 1000 ' указываем какой порт должен слушать сервер
Winsock1.Listen ' заставляем его слушать
End Sub
Private Sub winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close ' если сервер не закрыт, то закрываем его
Winsock1.Accept requestID ' принимаем запрос
End Sub
Private Sub winsock1_dataArrival(ByVal BytesTotal As Long)
Dim Data As Variant  'обявляем переменную для хранения данных
Winsock1.GetData Picture1.Picture 'присваиваем переменной Data все принятые данные
End Sub

