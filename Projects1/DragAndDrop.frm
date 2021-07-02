VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   240
      Left            =   1800
      Picture         =   "DragAndDrop.frx":0000
      Top             =   1200
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Image1.Left = X
Image1.Top = Y
End Sub

