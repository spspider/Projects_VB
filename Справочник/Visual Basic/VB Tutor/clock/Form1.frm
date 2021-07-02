VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12:12:12"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12:12:12"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   50
      TabIndex        =   1
      Top             =   50
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label1.Caption = Time
Label2.Caption = Time
End Sub
