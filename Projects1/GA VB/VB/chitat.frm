VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   4395
   ClientTop       =   3960
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7020
   Begin VB.CommandButton Command1 
      Caption         =   "Читиать"
      Height          =   2535
      Left            =   4800
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin HTTSLibCtl.TextToSpeech S1 
      Height          =   735
      Left            =   360
      OleObjectBlob   =   "chitat.frx":0000
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
S1.Speak Text1.Text
End Sub
