VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{66D3CBC4-D446-4BAA-B8B2-AF97BC09A7D2}#1.0#0"; "AudioControl.ocx"
Object = "{59E060F0-4FD5-4C80-AF3C-D1B7E0ED65B2}#1.0#0"; "CompControls.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super"
   ClientHeight    =   7395
   ClientLeft      =   3990
   ClientTop       =   4050
   ClientWidth     =   9195
   Icon            =   "My corporation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MM2 
      Height          =   975
      Left            =   5040
      TabIndex        =   16
      Top             =   3360
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1720
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      DeviceType      =   ""
      FileName        =   "E:\Видео\Разное\приколы\___________.AVI"
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   4320
      ScaleHeight     =   2835
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   4320
      Width           =   3735
   End
   Begin AUDIOCONTROLLib.Knob Knob1 
      Height          =   975
      Left            =   720
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1720
      _StockProps     =   0
      Min             =   0
      Max             =   3000
   End
   Begin MCI.MMControl MM 
      Height          =   2970
      Left            =   3600
      TabIndex        =   13
      Top             =   4320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   5239
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      OLEDropMode     =   1
      DeviceType      =   ""
      FileName        =   "E:\Видео\Разное\приколы\_______________AE___.MPEG"
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Caption         =   "информация"
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4935
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   615
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   2040
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer3 
         Interval        =   10
         Left            =   2400
         Top             =   480
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ещё"
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "мой компьютер"
         Height          =   495
         Left            =   3600
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Настройки экрана"
         Height          =   495
         Left            =   3600
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Выключить компьютер"
         Height          =   495
         Left            =   3600
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Image Im1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   300
         Picture         =   "My corporation.frx":324A
         Top             =   405
         Width           =   330
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1695
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2990
      _Version        =   393216
      Enabled         =   0   'False
      Orientation     =   1
      Max             =   20
      TickStyle       =   2
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin CompControler.CompControl CC1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Frame Frame1 
      Caption         =   "Область ввода"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command6 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   2000
         Width           =   255
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Top             =   1200
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   720
         Top             =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ОК"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Menu File 
      Caption         =   "файл"
      NegotiatePosition=   1  'Left
      Begin VB.Menu Exit 
         Caption         =   "выход"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Dim sl As Integer
Dim k As Integer
Private Sub Command1_Click()
Timer2.Enabled = True
End Sub
Private Sub Command2_Click()
CC1.Move_File
End Sub
Private Sub Command3_Click()
CC1.Display_Settings
End Sub
Private Sub Command4_Click()
CC1.Add_HardWare
End Sub
Private Sub Command5_Click()
CC1.Sounds_Settings
End Sub
Private Sub Command7_Click()
CD1.ShowOpen
End Sub
Private Sub Exit_Click()
End
End Sub
Private Sub Form_Load()
Label1.Caption = Text1.Text
Dim a As Integer
Dim b As Integer
Timer1.Enabled = True
MM.Command = "open"
MM.hWndDisplay = Picture1.hWnd
MM2.hWndDisplay = Picture2.hWnd
MM2.Command = "open"
k = 1
t = 10
sl = 1
End Sub
Private Sub Text1_Change()
b = "   дней"
a = Text1.Text
Label1.Caption = a + b
If a = "" Then
Label1.Caption = "введите числа"
End If
End Sub
Private Sub Timer2_Timer()
Label2.Caption = Str(t)
t = t - 1
sl = sl + 1
Slider1.Value = sl
If Label2.Caption < 0 Then
Text1.BackColor = &H8000000D
Text1.Enabled = False
Timer1.Enabled = True
Timer2.Interval = 0
Label2.Caption = "время вышло"
Command1.Caption = "Ещё"
Timer2.Enabled = False
End If
End Sub
Private Sub Timer3_Timer()
Im1.Left = Knob1.Position
End Sub
