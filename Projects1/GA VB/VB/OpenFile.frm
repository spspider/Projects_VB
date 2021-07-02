VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Пуск"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "файл..."
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Путь к открываемому файлу:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error GoTo go:
FileCopy Text2.Text, Text1.Text
go:
End Sub
Private Sub Command1_Click()
On Error GoTo er:
CC1.ShowSave
Text1.Text = CC1.FileName
er:
End Sub

Private Sub Command3_Click()
On Error GoTo er1:
CC1.ShowOpen
Text2.Text = CC1.FileName
er1:
End Sub

Private Sub Command4_Click()
CC1.ShowOpen
Text3.Text = CC1.FileName
End Sub

Private Sub Command5_Click()
Shell (Text3.Text), vbNormalFocus
End Sub

