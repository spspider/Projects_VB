VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   8775
   Icon            =   "MainMenuDouble.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   Begin VB.OptionButton Op4 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op3 
      Caption         =   "Option1"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
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
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
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

For a = 0 To 20
Op1.Top = -Form1.Width
Op1.Left = Cos(((v1 / 10) + a) / 2) * 3000 + 1500
Line (Op1.Left, Op1.Top)-(Op2.Left, Op2.Top), RGB(0, 255, 100)
Op2.Top = Form1.Height - Op1.Top
Op2.Left = (Form1.Width - 6000) - Op1.Left
Next

For a1 = 1 To 20
Op3.Top = -Form1.Width
Op3.Left = Cos(((v1 / -10) + a1) / 2) * 3000 + 8000
Line (Op3.Left, Op3.Top)-(Op4.Left, Op4.Top), RGB(0, 255, 100)
Op4.Top = Form1.Height - Op3.Top
Op4.Left = (Form1.Width + 6000) - Op3.Left
Next

End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
