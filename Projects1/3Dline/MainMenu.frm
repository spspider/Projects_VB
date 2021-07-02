VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   8775
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   1440
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

Dim op1L As Integer
Dim op1T As Integer
Dim op2L As Integer
Dim op2T As Integer

Private Sub Command1_Click()
v1 = 0
a = 0
v2 = 0
End Sub

Private Sub Timer1_Timer()

Form1.Cls

v1 = v1 - 1
v2 = v2 + 100

For a = 1 To 15
op1T = -Form1.Width
op1L = Cos(((v1 / -10) + a) / 2) * 3000 + Form1.Width / 4
Line (op1L, op1T)-(op2L, op2T), RGB(0, 255, 100)
op2T = Form1.Height - op1T
op2L = (Form1.Width - Form1.Width / (15 / 10)) - op1L
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
