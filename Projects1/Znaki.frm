VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serega"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   Icon            =   "Znaki.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   156
   ScaleMode       =   2  'Point
   ScaleWidth      =   234
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Скрыть значки"
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Показать значки"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow& Lib "user32" (ByVal q&, ByVal q1&)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal w As String, ByVal w1 As String) As Long
Private Sub Command1_Click()
r = FindWindow("progman", vbNullString)
Call ShowWindow(r, 1)
End Sub
Private Sub Command2_Click()
r = FindWindow("progman", vbNullString)
Call ShowWindow(r, 0)
End Sub
