VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.OLE OLE1 
      Class           =   "MIDFile"
      Height          =   855
      Left            =   120
      OleObjectBlob   =   "OLEobj.frx":0000
      SourceDoc       =   "C:\Documents and Settings\Cергей\Мои документы\Моя музыка\MIDI\X-Files.MID"
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OLE1.DoVerb (play)
End Sub

Private Sub Command2_Click()
OLE1.Close
End Sub
