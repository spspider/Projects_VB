VERSION 5.00
Begin VB.Form frmPong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pong"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Help"
      Height          =   435
      Left            =   2640
      TabIndex        =   16
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   435
      Left            =   1800
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.Timer Time 
      Interval        =   100
      Left            =   4680
      Top             =   4680
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   3480
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   435
      Left            =   960
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox picBall 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7440
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Start!"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox Front 
      BackColor       =   &H00FFFFFF&
      Height          =   3960
      Left            =   120
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   3
      Top             =   480
      Width           =   6195
   End
   Begin VB.PictureBox Back 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3960
      Left            =   120
      ScaleHeight     =   264
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   7440
      Picture         =   "frmMain.frx":0642
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   7440
      Picture         =   "frmMain.frx":0984
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "0:00:0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5400
      TabIndex        =   14
      Top             =   0
      Width           =   960
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   12
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Difficulty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   1425
   End
   Begin VB.Label lblDifficulty 
      AutoSize        =   -1  'True
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lives:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      TabIndex        =   8
      Top             =   0
      Width           =   930
   End
   Begin VB.Label lblLife 
      AutoSize        =   -1  'True
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "frmPong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Vars
    Dim Active As Boolean
    Dim KD(255) As Boolean
    Dim KU(255) As Boolean
    
    Dim TimeOn As Boolean
    Dim Milli As Integer
    Dim Secs As Integer
    Dim Mins As Integer
    
Public Sub GetScore()
    lblScore.Caption = Score
End Sub

Public Sub MoveBall()
    Ball.y = Ball.y + DY
    Ball.x = Ball.x + DX
    
    If Ball.x <= 0 Then
        DX = DX * -1
        Life = Life - 1
    End If
    If Ball.x >= (Front.Width - Ball.w) Then
        DX = DX * -1
        Life = Life - 1
    End If
    
    If Ball.y <= 0 Then DY = DY * -1
    If Ball.y + Ball.h >= Front.Height Then DY = DY * -1
    
    lblLife.Caption = Life
'PADDLE 1
    If Ball.x <= (Player.x + Player.w) Then
        If Ball.y >= Player.y Then
            If Ball.y <= (Player.y + Player.h) Then
                DX = DX * -1
                If lblDifficulty.Caption = "Easy" Then
                    Score = Score + 100
                ElseIf lblDifficulty.Caption = "Hard" Then
                    Score = Score + 200
                ElseIf lblDifficulty.Caption = "Hell" Then
                    Score = Score + 300
                ElseIf lblDifficulty.Caption = "Uh Oh" Then
                    Score = Score + 400
                End If
                Call GetScore
            End If
        End If
    End If
    
    'Paddle2
    If (Ball.x + Ball.w) >= Player2.x Then
        If Ball.y >= Player2.y Then
            If Ball.y <= (Player2.y + Player2.h) Then
                DX = DX * -1
                If lblDifficulty.Caption = "Easy" Then
                    Score = Score + 100
                ElseIf lblDifficulty.Caption = "Hard" Then
                    Score = Score + 200
                ElseIf lblDifficulty.Caption = "Hell" Then
                    Score = Score + 300
                ElseIf lblDifficulty.Caption = "Uh Oh" Then
                    Score = Score + 400
                End If
                Call GetScore
            End If
        End If
    End If
    If Life = 0 Then
        Active = False
        cmdOptions.Enabled = True
        TimeOn = False
        MsgBox "Congratulations, you recieved a score of " & Score & " at the level of " & lblDifficulty.Caption & "."
        Life = 10
    End If
End Sub

Public Sub CheckKeys()
    If KD(vbKeyEscape) = True Then
        Active = False
        cmdOptions.Enabled = True
        TimeOn = False
    End If
    
    'Check if any key's pressed and move the player if in viewport
    If KD(vbKeyF12) = True Then
        Score = Score + 1000
        Call GetScore
    End If

    If KD(vbKeyDown) = True Then
        Player.y = Player.y + 3
        Player2.y = Player.y
        If (Player.y + Player.h) >= Front.Height Then Player.y = (Front.Height - Player.h)
        End If
    
    If KD(vbKeyUp) = True Then
        Player.y = Player.y - 3
        Player2.y = Player.y
        If Player.y <= 0 Then Player.y = 0
    End If

End Sub

Public Sub Draw()
    Dim A As Long
    Dim B As Long
    
    'Draw the background tiles (grass)
    For A = 0 To Front.Width / Tile.w
        For B = 0 To Front.Height / Tile.h
            BitBlt BackDC, A * Tile.w, B * Tile.h, Tile.w, Tile.h, Tile.DC, 0, 0, vbSrcCopy
        Next
    Next
    
    'And now the player
    BitBlt BackDC, Player.x, Player.y, Player.w, Player.h, Player.DC, Player.w, 0, vbSrcPaint 'The mask
    BitBlt BackDC, Player.x, Player.y, Player.w, Player.h, Player.DC, 0, 0, vbSrcAnd 'The player himself
    'And now the player
    BitBlt BackDC, Player2.x, Player2.y, Player2.w, Player2.h, Player2.DC, Player2.w, 0, vbSrcPaint 'The mask
    BitBlt BackDC, Player2.x, Player2.y, Player2.w, Player2.h, Player2.DC, 0, 0, vbSrcAnd 'The player himself
    'And now the picBall
    BitBlt BackDC, Ball.x, Ball.y, Ball.w, Ball.h, Ball.DC, Ball.w, 0, vbSrcPaint 'The mask
    BitBlt BackDC, Ball.x, Ball.y, Ball.w, Ball.h, Ball.DC, 0, 0, vbSrcAnd 'The ball itself
    'Flip
    BitBlt FrontDC, 0, 0, Front.Width, Front.Height, BackDC, 0, 0, vbSrcCopy
End Sub


Public Sub Main()
    While Active
        'Check the keys we set in Front_KeyDown
        CheckKeys
        
        'Draw scene
        Draw
        
        'Move ball
        MoveBall
        
        'And get new windows messages
        DoEvents
    Wend
    
    cmdRun.Enabled = True
End Sub


Private Sub cmdAbout_Click()
    MsgBox "Survival Pong" & vbCrLf & "-----------------" & vbCrLf & "Created By Steve Mack"
End Sub

Private Sub cmdExit_Click()
    MsgBox "Thank you for trying my Survival Pong demo." & vbCrLf & vbTab & "- Steve Mack"
    Active = False
    End
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show
End Sub

Private Sub cmdRun_Click()
    Score = 0
    If lblDifficulty.Caption = "Easy" Then
        DX = 1
        DY = 1
    ElseIf lblDifficulty.Caption = "Hard" Then
        DX = 2
        DY = 2
    ElseIf lblDifficulty.Caption = "Hell" Then
        DX = 3
        DY = 3
    End If
    Mins = 0
    Milli = 0
    Secs = 0
    Active = True
    Player.y = (Front.Height / 2) - Player.h
    Player.x = -5
    Player2.y = Player.y
    Player2.x = Front.Width - 15
    Ball.y = 100
    Ball.x = 100
    TimeOn = True
    cmdRun.Enabled = False
    cmdOptions.Enabled = False
    Front.SetFocus
    Main

End Sub


Private Sub Command1_Click()
    MsgBox "To Move: Use the up and down arrows to move your paddles." & vbCrLf & "Rules: You lose a life every time the ball hits the side, use your paddles to stop it!"
End Sub

Private Sub Form_Load()
    Milli = 0
    Secs = 0
    Mins = 0
    Score = 0
    Call GetScore
    DX = 1
    DY = 1
    Life = 10
    'Setup back buffer
    Back.Move 0, 0, Front.Width, Front.Height
    
    'Get DCs of the picture boxes
    BackDC = Back.hDC
    FrontDC = Front.hDC
    

    Player.DC = picPlayer.hDC
    Player2.DC = picPlayer.hDC
    Tile.DC = picTile(0).hDC
    Ball.DC = picBall.hDC

    'Get size of pictures

    Player.w = picPlayer.Width / 2 'Because of the mask
    Player.h = picPlayer.Height
    Player2.w = picPlayer.Width / 2 'Because of the mask
    Player2.h = picPlayer.Height
    
    Ball.w = picBall.Width / 2 'Because of the mask
    Ball.h = picBall.Height

    Tile.w = picTile(0).Width
    Tile.h = picTile(0).Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Thank you for trying my Survival Pong demo." & vbCrLf & vbTab & "- Steve Mack"
    Active = False
    End

End Sub
Private Sub Front_KeyDown(KeyCode As Integer, Shift As Integer)
    'Set state so wen can check it later
    KD(KeyCode) = True
    KU(KeyCode) = False
End Sub


Private Sub Front_KeyUp(KeyCode As Integer, Shift As Integer)
    KD(KeyCode) = False
    KU(KeyCode) = True
End Sub

Private Sub Time_Timer()
    If TimeOn = True Then
        Milli = Milli + 1
        If Milli > 9 Then
            Secs = Secs + 1
            Milli = 0
        End If
        If Secs > 59 Then
            Mins = Mins + 1
            Secs = 0
        End If
    
        If Secs < 10 Then
            lblTime.Caption = Mins & ":" & "0" & Secs & ":" & Milli
        Else: lblTime.Caption = Mins & ":" & Secs & ":" & Milli
        End If
    End If
End Sub

Private Sub tmrSpeedIncrease_Timer()
    If Score >= 1000 Then
        DX = 2
        DY = 2
        Score = Score + 1000
    ElseIf Score >= 5000 Then
        DX = 3
        DY = 3
        Score = Score + 1000
    ElseIf Score >= 10000 Then
        DX = 4
        DY = 4
        Score = Score + 1000
    ElseIf Score >= 15000 Then
        DX = 5
        DY = 5
        Score = Score + 1000
    End If
    
    Call GetScore
End Sub
