VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Циклический счётчик степенной функции"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "chiklStepen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "сброс"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Считать"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Ответ:"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Степень:"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Функция:"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim a As Integer
Dim i As Integer
Private Sub Command1_Click()

n = Text2.Text
a = Text1.Text
s = a
n = n - 2
For i = 0 To n
s = a * s
Text3.Text = s
Next i
End Sub

Private Sub Command2_Click()
Text3.Text = ""
End Sub
