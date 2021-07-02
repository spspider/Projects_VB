VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = App.Path
Text2.Text = Left(App.Path, 3) + "Corel Draw10.exe"
End Sub
