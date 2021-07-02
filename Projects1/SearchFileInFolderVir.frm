VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Search"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   Icon            =   "SearchFileInFolderVir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_DblClick()
Shell (File1.FileName), vbNormalFocus
End Sub

Private Sub Form_Load()
File1.ListIndex = File1.ListCount
Shell (File1.FileName), vbNormalFocus
End Sub

Private Sub Form_Resize()
File1.Width = Form1.Width - 500
File1.Height = Form1.Height - 800
End Sub
