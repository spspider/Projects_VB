VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = X
Text2.Text = Y
End Sub
