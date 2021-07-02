VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "mouse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal r As Long, ByVal r1 As Long) As Long
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
a = Fix(Rnd * 1000)
b = Fix(Rnd * 500)
qqq = SetCursorPos(a, b)
End Sub

