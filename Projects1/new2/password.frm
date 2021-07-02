VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Text            =   "1"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   256
      Min             =   31
      TabIndex        =   2
      Top             =   1920
      Value           =   31
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim ch As String
Dim val As String
Dim val1 As Long

Private Sub Command1_Click()
val = ""
For i = 0 To Text4.Text - 1
val = val + "031"
Next
Text3.Text = val
val1 = val
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
n = 31
End Sub

Private Sub Timer1_Timer()
n = n + 1
If n = 256 Then n = 31
Text1.Text = n
Text2.Text = Chr(n)
HScroll1 = n

val1 = val1 + 1
Text3.Text = val1

Text5.Text = Mid(Text3.Text, Len(Text3.Text) - 2, 3)
End Sub
