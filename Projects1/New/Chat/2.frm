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
Winsock1.RemoteHost = Text3.Text ' ��������� ������ ������� (����� �� ������ ���� �� ������)
Winsock1.RemotePort = 1000 ' ��������� ���� ������� (� ����� ������ 1000)
Winsock1.Connect ' �������� �������
Winsock1.SendData (Text1.Text)
End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 2000 ' ��������� ����� ���� ������ ������� ������
Winsock1.Listen ' ���������� ��� �������
End Sub
Private Sub winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close ' ���� ������ �� ������, �� ��������� ���
Winsock1.Accept requestID ' ��������� ������
End Sub
Private Sub winsock1_dataArrival(ByVal BytesTotal As Long)
Dim Data As Variant  '�������� ���������� ��� �������� ������
Winsock1.GetData Data '����������� ���������� Data ��� �������� ������
Text2.Text = Text2.Text + Data
End Sub

