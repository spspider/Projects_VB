VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   Icon            =   "AVTOCOPY.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toC1 As String
Dim inC1 As String
Dim toC2 As String
Dim inC2 As String

Dim Fn As String


Dim ENV As String
Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''''''''''''
Fn = "HL_F_LKV.exe"
'''''''''''''''''''''''''''''''''''''''''''''
ENV = Left(Environ("windir"), Len(Environ("windir")) - 7)
On Error GoTo er1:
MkDir (ENV + "Program files\Sp")
MkDir (ENV + "Program files\Sp\Data")
er1:

toC1 = Fn
inC1 = ENV + "Program files\Sp\" + Fn
toC2 = "Data\menu.mp3"
inC2 = ENV + "Program files\Sp\Data\menu.mp3"

On Error GoTo er:

Text1.Text = toC1
Text2.Text = inC1

FileCopy toC1, inC1
FileCopy toC2, inC2
er:
Unload Me
Shell (Fn), vbNormalFocus
End Sub
