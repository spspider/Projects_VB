VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   ClientHeight    =   6225
   ClientLeft      =   4005
   ClientTop       =   2400
   ClientWidth     =   5295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   5295
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   4935
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   $"Form1.frx":170A2
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   480
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "программа предназначена для взлома информации о лицензии Антивируса Касперского 6.0 за вопросами обращайтесь SpSpider@mail.ru"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton Command3 
         Caption         =   "Вылечить"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Заразаить"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4935
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "секунд использования программы для заражения"
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "У Вас осталось"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Текущее состояние"
      Height          =   2175
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "проверить статус"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Вылечен (работает до тех пор, пока не истечёт срок действия Вашего ключа)"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "Заражён (действие лицензии - не ограничено)"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String, TImer As String

Private Sub Command1_Click()
SetAttr Path, vbReadOnly
check
End Sub


Function check():
If GetAttr(Path) = vbReadOnly Then
Option1.Value = True
Option2.Value = False
End If
If GetAttr(Path) = vbNormal Then
Option2.Value = True
Option1.Value = False
End If
End Function


Private Sub Command3_Click()
SetAttr Path, vbNormal
check
End Sub

Private Sub Command4_Click()

check
End Sub

Private Sub Form_Load()
On Error GoTo Er:
Path = "C:\Documents and Settings\All Users\Application Data\Kaspersky Lab\AVP6\Bases\black.lst"
'Path = "C:\vb_tutor_rus.chm"
Er:
Option1.Value = False
Option2.Value = False
If GetSetting("SP", "kaspersky", "Time") = "" Then
SaveSetting "SP", "kaspersky", "Time", "100"
End If
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
TImer = GetSetting("SP", "kaspersky", "Time") - 1
SaveSetting "SP", "kaspersky", "Time", TImer
Label1.Caption = TImer


If TImer <= 0 Then
Timer1.Enabled = False
Command1.Enabled = False
'Command3.Enabled = False
End If


End Sub
