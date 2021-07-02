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

Dim op1L As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long
Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 - 1
v2 = v2 + 2
For a = 1 To 30
op1T = Cos(((v1 / 10) + a) / 10) * 3000 + Form1.Width / 4 + Sin(v1 / 20) * 1000
op1L = Sin(((v1 / 10) + a) / 10) * 3000 + Form1.Height / 2 + Cos(v1 / 20) * 1000
Line (op1L, op1T)-(op2L, op2T), RGB(0, 255, 100)
op2T = Sin(((v1 / 10) + a) / 10) * 3000 + Form1.Width / 4 + Cos(v1 / 20) * 1000
op2L = Cos(((v1 / 10) + a) / 10) * 3000 + Form1.Height / 2 + Sin(v1 / 20) * 1000
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
