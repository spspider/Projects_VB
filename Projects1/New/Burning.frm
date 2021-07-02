VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calc"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label4"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Speed Kb/s:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Size Mb:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
Label3.Caption = Round(((Text3.Text * 1024) / Text1.Text) / 60)
Label4.Caption = ((((Text3.Text * 1024) / Text1.Text) / 60) - Round(((Text3.Text * 1024) / Text1.Text) / 60)) * 60
Label5.Caption = Round(((Text3.Text * 1024) / Text1.Text) / 3600)
er:
End Sub


