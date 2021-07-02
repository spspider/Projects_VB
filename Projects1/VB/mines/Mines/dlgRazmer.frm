VERSION 5.00
Begin VB.Form dlgRazmer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Размер игрового поля"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "dlgRazmer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "&Число мин:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "&Ширина:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "&Высота:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "dlgRazmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  SizeY = Int(Val(Text1))
  SizeX = Int(Val(Text2))
  KolMines = Int(Val(Text3))
  If SizeX < 8 Then SizeX = 8
  If SizeX > 30 Then SizeX = 30
  If SizeY < 8 Then SizeY = 8
  If SizeY > 20 Then SizeY = 20
  If KolMines < 10 Then KolMines = 10
  If KolMines > (SizeX - 1) * (SizeY - 1) Then KolMines = (SizeX - 1) * (SizeY - 1)
  Unload Me
End Sub

Private Sub Form_Load()
  Text1 = SizeY
  Text2 = SizeX
  Text3 = KolMines
End Sub
