VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   3735
   ClientTop       =   2400
   ClientWidth     =   5010
   Icon            =   "Shutdown.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      Caption         =   "запуск команды"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Скрыть"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "исполнить в:"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Shutdown.frx":1272
         Left            =   120
         List            =   "Shutdown.frx":1279
         TabIndex        =   15
         Text            =   "Команда"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "исполнимый код"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   4335
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Shutdown.frx":1287
         Left            =   2280
         List            =   "Shutdown.frx":1294
         TabIndex        =   12
         Text            =   "команда - ключ"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RUN NOW"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "до след. числа"
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "последний вызов"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label2 
            Caption         =   "ещё не запукалась..."
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "текущая дата"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2415
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Text            =   "00:00:00"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "Time"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   " время запуска"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   " текущее время"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   $"Shutdown.frx":12C5
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim InDate As String
Dim GCom As String
Dim ICom As String
Dim Path As String

Private Sub Check2_Click()
If Check2.Value = 1 Then
''''''''''''''''''''''''
Text1.Enabled = True
Text2.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Timer1.Enabled = True
End If
If Check2.Value = 0 Then
Text1.Enabled = False
Text2.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Timer1.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
''''''''''''''''''''''''''''''''''''''''''''SHUTDOWN_
If Combo2.ListIndex = 0 Then
If Combo1.ListIndex = 0 Then ICom = ""
If Combo1.ListIndex = 1 Then ICom = " -s"
If Combo1.ListIndex = 2 Then ICom = " -r"
If Combo1.ListIndex = 3 Then ICom = " -f"
End If

Path = ("C:\windows\system32\" + GCom + ICom)
Text3.Text = Path

End Sub

Private Sub Combo2_Click()
'''''''''''''''''''''''''''''''''''''''''''''SHUTDOWN
If Combo2.ListIndex = 0 Then
GCom = "Shutdown"
Combo1.Clear
Combo1.AddItem ("---")
Combo1.AddItem ("выключение")
Combo1.AddItem ("перезагрузка")
Combo1.AddItem ("завершение сеанса")
Combo1.Text = (Combo1.ListCount - 1) & " - подкоманд"
End If
'''''''''''''''''''''''''''''''''''''''''
Path = ("C:\windows\system32\" + GCom + ICom)
Text3.Text = Path
End Sub

Private Sub Command1_Click()
Shell (Path), vbHide 's
End Sub

Private Sub Command2_Click()
Form1.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo er:
InDate = GetSetting("Sp\Shutdown", "Date", "Newdate")
Label1.Caption = Date
Label2.Caption = InDate
er:
Combo2.ListIndex = 0
End Sub


Private Sub Timer1_Timer()
Text1.Text = Time
If Check1.Value = 1 Then
 If InDate = Date Then
 Shell (Path), vbHide 's
 End If
End If
If Text1.Text = Text2.Text Then
SaveSetting "Sp\Shutdown", "Date", "Newdate", Date
Shell (Path), vbHide 's
End If
End Sub
