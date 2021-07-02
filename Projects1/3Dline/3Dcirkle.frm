VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
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

Dim v2 As Long

Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer

Dim a3 As Integer
Dim v13 As Long
Dim op1L3 As Long
Dim op1T3 As Long
Dim op2L3 As Long
Dim op2T3 As Long

Private Sub Timer1_Timer()
Form1.Cls
v13 = v13 - 1
For a3 = 1 To 50
op1T3 = Sin(((v13 / 150) + a3 * 400)) * 1300 + Height - 2100
op1L3 = Cos(((v13 / 150) + a3 * 400)) * 1300 + 2100
Line (op1L3, op1T3)-(op2L3, op2T3), RGB(255, 255, 0), B
op2T3 = Height - 2100   '- op1T3
op2L3 = 2100  '- op1L3
Next

For a1 = 1 To 50
op1T = Sin(((v13 / 150) + a1 * 400)) * 1300 + 2000
op1L = Cos(((v13 / 150) + a1 * 400)) * 1300 + 2100
Line (op1L, op1T)-(op2L, op2T), RGB(255, 0, 0), B
op2T = 2000   '- op1T3
op2L = 2100  '- op1L3
Circle (op1L, op1T), 100, RGB(a * 2, a * 2, 255 - a)

Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
