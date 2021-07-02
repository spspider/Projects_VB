VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
On Error GoTo er:
Text1.Text = Drive1.Drive
File1.FileName = Drive1.Drive + "\"
er:
End Sub


