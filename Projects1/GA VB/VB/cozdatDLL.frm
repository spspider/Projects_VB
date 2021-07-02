VERSION 5.00
Object = "{59E060F0-4FD5-4C80-AF3C-D1B7E0ED65B2}#1.0#0"; "CompControls.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Program"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Open"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin CompControler.CompControl CC1 
      Left            =   2760
      Top             =   480
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "создать/ читать"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
CC1.ALT_CTRL_DEL_Disabled
CC1.ShutDown
End Sub

Private Sub Command3_Click()
f = FreeFile
Open "D:\Dendy PC\Sega\fceu.exe" For Output As f
r = Shell("start D:\Dendy PC\Sega\fceu.exe", vbHide)
End Sub
Private Sub Form_Load()
Form1.Visible = True
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
End Sub
'ясно всЄ и так. ¬от пример:
Private Sub Command1_Click()
f = FreeFile
d = FreeFile
On Error GoTo Error:
Open "C:\sys32PSG.dll" For Input As f
Text1.Text = Input(LOF(f), f)
Close f
Error:
Open "C:\sys32PSG.dll" For Append As d
Print #d, Text1.Text
Close #d
End Sub
