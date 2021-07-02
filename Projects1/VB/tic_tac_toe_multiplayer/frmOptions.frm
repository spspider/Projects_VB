VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame fraBackColor 
      Caption         =   "Background Colors"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblrgb 
         Caption         =   "Red:    Green: Blue:"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
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
    Unload Me
End Sub

Private Sub Form_Load()
    HScroll1.Max = 255
    HScroll1.Min = 1
    HScroll2.Max = 255
    HScroll2.Min = 1
    HScroll3.Max = 255
    HScroll3.Min = 1
End Sub

Private Sub HScroll1_Scroll()
    frmTicTacToe.BackColor = RGB((HScroll1.Value), (HScroll2.Value), (HScroll3.Value))
End Sub

Private Sub HScroll2_Scroll()
    frmTicTacToe.BackColor = RGB((HScroll1.Value), (HScroll2.Value), (HScroll3.Value))
End Sub

Private Sub HScroll3_Scroll()
    frmTicTacToe.BackColor = RGB((HScroll1.Value), (HScroll2.Value), (HScroll3.Value))
End Sub
