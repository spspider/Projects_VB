VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "OCX"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "OpenSP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   5895
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin ComctlLib.ProgressBar PB1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Form_Load()

File1.FileName = "Data\OCX\"

PB1.Max = (File1.ListCount - 1)

For n = 0 To (File1.ListCount - 1)
File1.ListIndex = n
FileCopy ("Data\OCX\" + File1.FileName), "D:\" + File1.FileName
PB1.Value = n
Next

Text1.Text = File1.FileName
'Shell "sp.exe", vbNormalFocus
'Unload Me

End Sub
