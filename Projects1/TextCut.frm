VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   " "
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Привет всем!"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v As Integer

Private Sub Timer1_Timer()
v = Len(Text1.Text)
Text1.Text = Mid((Text1.Text), 2, v)
    If Text1.Text = "" Then
         For l = 0 To 30
         Text1.Text = Text1.Text + "  "
         Next
    Text1.Text = Text1.Text + Text2.Text
    End If
End Sub
