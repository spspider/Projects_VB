VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "тариф -> мин"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "мин -> тариф"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = (Text1.Text / 60) * (323 / 100)
End Sub
Private Sub Command2_Click()
Text2.Text = (Text1.Text * 60) / (323 / 100)
End Sub
