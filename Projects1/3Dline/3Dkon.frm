VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   240
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


Dim Op1l As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long
Private Sub Timer1_Timer()
Form1.Cls

v1 = v1 + 1
v2 = v2 + 1
For a = 5 To 200
op1T = Sin(((v1 / 50) + a / 10) / 2) * 3000 + Height / 2
Op1l = Cos(((v1 / 50) + a / 4) / 2) * 3000 + Form1.Width / 2
Line (Op1l, op1T)-(op2L, op2T), RGB(0, a / 2, 200) ', B
op2T = Height / 2
op2L = Form1.Width / 2
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
