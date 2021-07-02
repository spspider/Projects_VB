VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLives 
      Caption         =   "Lives"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   3255
      Begin VB.ComboBox cmbLives 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "How many lives would you like?"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraGameSpeed 
      Caption         =   "Game Speed"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton optUhOh 
         Caption         =   "Uh Oh"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optHell 
         Caption         =   "Hell"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optHard 
         Caption         =   "Hard"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optEasy 
         Caption         =   "Easy"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOkay_Click()
    If optEasy.Value = True Then
        DX = 1
        DY = 1
        Difficulty = "Easy"
    ElseIf optHard.Value = True Then
        DX = 2
        DY = 2
        Difficulty = "Hard"
    ElseIf optHell.Value = True Then
        DX = 3
        DY = 3
        Difficulty = "Hell"
    ElseIf optUhOh.Value = True Then
        DX = 4
        DY = 4
        Difficulty = "Uh Oh"
    End If
    Life = cmbLives.Text
    frmPong.lblDifficulty.Caption = Difficulty
    frmPong.lblLife.Caption = Life
    Unload Me
End Sub

Private Sub Form_Load()
    cmbLives.AddItem "5"
    cmbLives.AddItem "10"
    cmbLives.AddItem "20"
    cmbLives.AddItem "30"
End Sub
