VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "play"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub
