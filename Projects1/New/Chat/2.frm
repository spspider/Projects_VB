VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Text            =   "192.168.1.36"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.RemoteHost = Text3.Text ' указываем адресс сервера (здесь ты указал свой же адресс)
Winsock1.RemotePort = 1000 ' указываем порт сервера (в нашем случае 1000)
Winsock1.Connect ' вызываем коннект
Winsock1.SendData (Text1.Text)
End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 2000 ' указываем какой порт должен слушать сервер
Winsock1.Listen ' заставляем его слушать
End Sub
Private Sub winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close ' если сервер не закрыт, то закрываем его
Winsock1.Accept requestID ' принимаем запрос
End Sub
Private Sub winsock1_dataArrival(ByVal BytesTotal As Long)
Dim Data As Variant  'обявляем переменную для хранения данных
Winsock1.GetData Data 'присваиваем переменной Data все принятые данные
Text2.Text = Text2.Text + Data
End Sub

