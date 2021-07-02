VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   1440
   End
   Begin VB.OptionButton Op2 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Op1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Long
Dim v2 As Long
Dim a As Integer
Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer
Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 - 1
v2 = v2 + 100
For a = 1 To 15
Op1.Top = -Form1.Width
Op1.Left = Sin(((-v1 / 10) + a) / 2) * 3000 + Form1.Height / 2
Line (Op1.Left, Op1.Top)-(Op2.Left, Op2.Top), RGB(0, 255, 100)
Op2.Top = Form1.Width / 2 - Op1.Top
Op2.Left = Form1.Height / 2 - Op1.Left
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
