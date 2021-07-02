VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTicTacToe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Chat 
      Left            =   3240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Game 
      Left            =   3720
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picGame 
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   720
      Width           =   4095
      Begin VB.CommandButton cmdOption 
         Caption         =   "&Game Options"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdOkay 
         Caption         =   "!"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtChat 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmTicTacToe.frx":0000
         Top             =   2040
         Width           =   3735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   7
         Left            =   840
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   5
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   4
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmbBoard 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect/Listen"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Port"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "IP Address"
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optClient 
      Caption         =   "Client"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optServer 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yourTurn As Boolean

Public Sub checkTriplets(a As Integer, b As Integer, c As Integer)
    Dim f As Integer
    If cmbBoard(a).Caption = cmbBoard(b).Caption And cmbBoard(a).Caption = cmbBoard(c).Caption Then
    If cmbBoard(a).Caption <> "" Then
        MsgBox cmbBoard(a).Caption & " wins!"
        For f = 0 To 8
            cmbBoard(f).Caption = ""
        Next f
        yourTurn = yourTurn
    End If
    End If
End Sub

Public Sub CheckWinner()
    Dim i As Integer
    Dim Check As Boolean
    
    Call checkTriplets(0, 1, 2)
    Call checkTriplets(0, 4, 8)
    Call checkTriplets(0, 3, 6)
    Call checkTriplets(1, 4, 7)
    Call checkTriplets(5, 2, 8)
    Call checkTriplets(2, 4, 6)
    Call checkTriplets(3, 4, 5)
    Call checkTriplets(6, 7, 8)
    
    Check = False
    For i = 0 To 8
        If cmbBoard(i).Caption = "" Then Check = True
    Next i
    If Check = False Then
        MsgBox "Draw!"
        For i = 0 To 8
            cmbBoard(i).Caption = ""
        Next i
        yourTurn = yourTurn
    End If
End Sub

Private Sub Chat_ConnectionRequest(ByVal requestID As Long)
    Chat.Close
    Chat.Accept requestID
    txtChat.Text = txtChat.Text & "<Connected>" & vbCrLf
End Sub

Private Sub Chat_DataArrival(ByVal bytesTotal As Long)
    Dim NewChatStuff As String
    Chat.GetData NewChatStuff
    txtChat.Text = txtChat.Text & "Opponent: " & NewChatStuff & vbCrLf
End Sub

Private Sub cmbBoard_Click(Index As Integer)
    If optServer.Value = True Then
        If yourTurn = True Then
            If cmbBoard(Index).Caption = "" Then
                cmbBoard(Index).Caption = "X"
                Game.SendData Index
                Call CheckWinner
            End If
        End If
    ElseIf optServer.Value = False Then
        If yourTurn = True Then
            If cmbBoard(Index).Caption = "" Then
                cmbBoard(Index).Caption = "O"
                Game.SendData Index
                Call CheckWinner
            End If
        End If
    End If
    yourTurn = False
End Sub

Private Sub cmdConnect_Click()
    If optServer.Value = True Then
        yourTurn = False
        Game.LocalPort = txtPort.Text
        Game.Listen
        Chat.LocalPort = txtPort.Text + 1
        Chat.Listen
    Else:
        yourTurn = True
        Game.RemoteHost = txtIP.Text
        Game.RemotePort = txtPort.Text
        Game.Connect
        Chat.RemoteHost = txtIP.Text
        Chat.RemotePort = txtPort.Text + 1
        Chat.Connect
    End If
End Sub

Private Sub cmdOkay_Click()
    Chat.SendData txtMsg.Text
    txtChat.Text = txtChat.Text & "You: " & txtMsg.Text & vbCrLf
    txtMsg.Text = ""
    
End Sub

Private Sub cmdOption_Click()
    frmOptions.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Game.Close
End Sub

Private Sub Game_ConnectionRequest(ByVal requestID As Long)
    Game.Close
    Game.Accept requestID
    MsgBox "Connected!"
End Sub

Private Sub Game_DataArrival(ByVal bytesTotal As Long)
    Dim EnemyCoord As Integer
    Game.GetData EnemyCoord
    If optServer.Value = True Then
        If yourTurn = False Then
            If cmbBoard(EnemyCoord).Caption = "" Then
                 cmbBoard(EnemyCoord).Caption = "O"
                 yourTurn = True
            End If
        End If
    ElseIf optServer.Value = False Then
        If yourTurn = False Then
            If cmbBoard(EnemyCoord).Caption = "" Then
                 cmbBoard(EnemyCoord).Caption = "X"
                 yourTurn = True
            End If
        End If
    End If
    Call CheckWinner
End Sub

Private Sub lblLost_Click()
End Sub

Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Chat.SendData txtMsg.Text
        txtChat.Text = txtChat.Text & "You: " & txtMsg.Text & vbCrLf
        txtMsg.Text = ""
    End If
End Sub
