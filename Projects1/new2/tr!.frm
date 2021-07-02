VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "tr!.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "tr!.frx":170A2
   ScaleHeight     =   6285
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3360
      TabIndex        =   0
      Top             =   3600
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub Timer1_Timer()
x = x + 1
If x = 1 Then Label1.Caption = "    8=--"
If x = 2 Then Label1.Caption = "   8-=-"
If x = 3 Then Label1.Caption = "  8--="
If x = 4 Then Label1.Caption = " 8---="
If x = 5 Then Label1.Caption = "8--- ="
If x = 6 Then Label1.Caption = " 8---="
If x = 7 Then Label1.Caption = "  8--="
If x = 8 Then Label1.Caption = "   8-=-"
If x = 9 Then Label1.Caption = "    8=--"
If x = 10 Then x = 1
End Sub
