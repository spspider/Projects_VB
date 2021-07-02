VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
Winsock1.RemoteHost = "127.0.0.1" ' указываем адресс сервера (здесь ты указал свой же адресс)
Winsock1.RemotePort = 1000 ' указываем порт сервера (в нашем случае 1000)
Winsock1.Connect ' вызываем коннект
End Sub
Private Sub Command1_Click()
Winsock1.SendData "open"
End Sub
