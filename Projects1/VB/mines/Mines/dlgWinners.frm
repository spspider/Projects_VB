VERSION 5.00
Begin VB.Form dlgWinners 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Лучшее время"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "dlgWinners.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Сброс результатов"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Победители по разрядам"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Профессионалы:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Любители:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Новички:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "dlgWinners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdReset_Click()
  Dim i As Integer
  
  For i = 0 To 2
    Rec(i).Name = conDefName
    Rec(i).Time = conDefTime
  Next
  Form_Load
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  For i = 0 To 2
    Label2(i) = Trim$(Str$(Rec(i).Time)) + " сек."
    Label3(i) = Rec(i).Name
  Next
End Sub
