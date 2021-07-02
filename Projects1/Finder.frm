VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   3780
   ClientTop       =   3690
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7005
   Begin VB.CommandButton Command4 
      Caption         =   "STOP ALL"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PLAY"
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Поиск"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   6375
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "D:\mus\"
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub Command1_Click()
File1.FileName = Text1.Text
End Sub
Private Sub Command2_Click()
Call mciExecute(a)
End Sub

Private Sub Command3_Click()
b = "close " + Text3.Text
Call mciExecute(b)
End Sub

Private Sub Command4_Click()
Call mciExecute("close all")
End Sub

Private Sub File1_Click()
Text2.Text = File1.FileName
Text4.Text = File1.ListIndex
Text3.Text = Text1.Text & Text2.Text
a = "play " + Text3.Text
End Sub

