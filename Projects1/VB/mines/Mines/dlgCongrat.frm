VERSION 5.00
Begin VB.Form dlgCongrat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Примите поздравления!"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   Icon            =   "dlgCongrat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "dlgCongrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim s As String
  
  s = "Поздравляю Вы стали чемпионом среди "
  Select Case Rang
    Case 0
      s = s + "новичков."
    Case 1
      s = s + "любителей."
    Case 2
      s = s + "профессионалов."
  End Select
  s = s + vbCrLf + "Введите свое имя:"
  Label1 = s
  Text1 = conDefName
  Text1.SelLength = Len(conDefName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Rec(Rang).Name = Text1
End Sub
