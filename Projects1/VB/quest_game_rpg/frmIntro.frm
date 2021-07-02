VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   5040
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   5040
      Top             =   120
   End
   Begin VB.TextBox txtStory 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmIntro.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2400
      Picture         =   "frmIntro.frx":0006
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   1920
      Picture         =   "frmIntro.frx":0C48
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   1440
      Picture         =   "frmIntro.frx":188A
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1920
      Picture         =   "frmIntro.frx":24CC
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   480
      Picture         =   "frmIntro.frx":310E
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   960
      Picture         =   "frmIntro.frx":3D50
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   1440
      Picture         =   "frmIntro.frx":4992
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   480
      Picture         =   "frmIntro.frx":55D4
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   960
      Picture         =   "frmIntro.frx":6216
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label etqLoadGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Load Current Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Width           =   2010
   End
   Begin VB.Label etqNewGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   1080
      Picture         =   "frmIntro.frx":6E58
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2670
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   320
      Y2              =   0
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub etqLoadGame_Click()
  Timer2.Enabled = False
    CloseMidi
    UnHook
  Call LoadGame
  Unload Me
End Sub

Private Sub etqLoadGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqNewGame.ForeColor = &HFFFFFF
  etqLoadGame.ForeColor = &HFF&
End Sub

Private Sub etqNewGame_Click()
  Timer1.Enabled = True
  Timer2.Enabled = False
  txtStory.Visible = True
  frmIntro.Cls
  frmIntro.BackColor = &H0&
  Midi = "Bad.mid"
  gHW = frmIntro.hWnd
  Call InitMusic
  Call LoadStory
End Sub

Private Sub etqNewGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  etqNewGame.ForeColor = &HFF&
  etqLoadGame.ForeColor = &HFFFFFF
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
    Case vbKeyEscape
    CloseMidi
    UnHook
  End Select

End Sub

Private Sub Form_Load()
  frmIntro.Height = 5175
  frmIntro.Width = 4890
  Midi = "intro.mid"
  gHW = frmIntro.hWnd
  Call InitMusic
  Call InitIntro
  'Call DrawIt2
End Sub

Public Sub InitIntro()
  'Starting the first Map.
  Call Map_Intro
  'Character Position.
  CharX = 4
  CharY = 5
  'Facing Position.
  CharFacing = 3
End Sub

Public Sub LoadStory()
  Dim Story(5) As String
  Dim DSpace As String
  DSpace = vbCrLf & vbCrLf
  Story(0) = "Far away, in the kigndom of Hyrule, there was a leyend of that the chosen one will save the kingdom from a great dark force."
  Story(1) = "The time has arrived for our hero sheik to do his job."
  Story(2) = "But the quest will not be easy, sheik must learn spells to get power and then meet with the Dark Force."
  Story(3) = "The future of the kingdom depends on you. Good Luck..."
  txtStory.Text = Story(0) & DSpace & Story(1) & DSpace & Story(2) & DSpace & Story(3)
End Sub

Private Sub Timer1_Timer()
  SAVE_MapLoaded = "A1"
  SAVE_Midi = "Town.mid"
  SAVE_OutsideMidi = "Town.mid"
  SAVE_SpeechLoaded = "A1"
  SAVE_CharX = 10
  SAVE_CharY = 7
  SAVE_CharFacing = 2
  SAVE_MapLoaded = "A1"
  SAVE_SpellCut = False
  SAVE_SpellDestroy = False
  Item = 0
  ItemTownFound1 = False
  
  frmGame.Show
  Timer1.Enabled = False
  Unload Me
End Sub

Public Sub IntroMov()
    
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    'PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    DrawIt2

End Sub


Private Sub Timer2_Timer()
  CharX = CharX + 1
  If CharX = 21 Then CharX = 10
  Call IntroMov
  DrawIt2
  'Timer2.Enabled = False
End Sub

Public Sub DrawIt2()
  For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler2
        If X + CharX + 0 < 1 Then PictureHandler2
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler2
        If Y + CharY + 0 > 51 Then PictureHandler2
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "A" 'Water
      '  PaintPicture imgBOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "S" 'Water
      '  PaintPicture imgBOR.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 32, (Y + 3) * 32
      'Case Is = "F" 'Border Bottom water
      '  PaintPicture ImgIB.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
                
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  End Select
  
End Sub

Public Sub PictureHandler2()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

