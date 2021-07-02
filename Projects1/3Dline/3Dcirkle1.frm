VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1L As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long

Dim v1 As Long

Private Sub Timer1_Timer()
Form1.Cls
v1 = v1 + 1
For a1 = 1 To 50
op1T = Sin(((v1 / 40) + a1 / 15)) * 300 + Height - 2100
op1L = Cos(((v1 / 40) + a1 / 15)) * 1300 + 2100
Line (op1L, op1T)-(op2L, op2T), RGB(a1 * 5, a1 * 5, 0) ', B
op2T = Sin(((v1 / 40) + a1 / 15)) * 290 + Height - 2100   '- op1T3
op2L = Cos(((v1 / 40) + a1 / 15)) * 1290 + 2100 '- op1L3
Next
End Sub
