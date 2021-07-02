VERSION 5.00
Object = "{CECDEAC1-E92C-11D2-B1AA-300962C10000}#1.0#0"; "mp3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider2 
      Height          =   2895
      Left            =   5880
      TabIndex        =   26
      Top             =   1440
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   5106
      _Version        =   393216
      Orientation     =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   4800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VFMP3PLAYERLib.VFmp3player VFmp3player1 
      Height          =   480
      Left            =   240
      TabIndex        =   24
      Top             =   4200
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VFMP3PLAYERLib.VFmp3scope VFmp3scope1 
      Height          =   495
      Left            =   3480
      TabIndex        =   23
      Top             =   600
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VFMP3PLAYERLib.VFmp3scope VFmp3scope2 
      Height          =   495
      Left            =   2040
      TabIndex        =   22
      Top             =   600
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VFMP3PLAYERLib.VFmp3level VFmp3level1 
      Height          =   975
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   1720
      _StockProps     =   0
   End
   Begin VFMP3PLAYERLib.VFmp3level VFmp3level2 
      Height          =   975
      Left            =   5520
      TabIndex        =   20
      Top             =   120
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1720
      _StockProps     =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   300
      Left            =   4218
      TabIndex        =   19
      Top             =   4350
      Width           =   765
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   300
      Left            =   3305
      TabIndex        =   18
      Top             =   4350
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   300
      Left            =   2393
      TabIndex        =   17
      Top             =   4350
      Width           =   765
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   300
      Left            =   1481
      TabIndex        =   16
      Top             =   4350
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Songinfo"
      Height          =   2715
      Left            =   165
      TabIndex        =   7
      Top             =   1305
      Width           =   2235
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   345
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   465
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   345
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1065
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   345
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1635
         Width           =   1635
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   345
         TabIndex        =   8
         Text            =   "Text10"
         Top             =   2250
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Totaltime"
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   2010
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "TotalFrame"
         Height          =   225
         Left            =   360
         TabIndex        =   14
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Samplerate"
         Height          =   225
         Left            =   360
         TabIndex        =   13
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Bitrate"
         Height          =   225
         Left            =   360
         TabIndex        =   12
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3TagInfo"
      Height          =   2715
      Left            =   2520
      TabIndex        =   0
      Top             =   1305
      Width           =   3195
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   315
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   330
         Width           =   2580
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   315
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   708
         Width           =   2580
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   315
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1086
         Width           =   2580
      End
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   315
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1464
         Width           =   2580
      End
      Begin VB.TextBox Text8 
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1842
         Width           =   2580
      End
      Begin VB.TextBox Text9 
         Height          =   330
         Left            =   315
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2220
         Width           =   2580
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VFmp3player1.Play
End Sub
Private Sub Command2_Click()
VFmp3scope2.ClearScope
VFmp3scope1.ClearScope
VFmp3level1.ClearLevel
VFmp3level2.ClearLevel
VFmp3player1.Stop
'VFmp3spectrum1.ClearSpectrum
End Sub
Private Sub Command3_Click()
VFmp3scope2.ClearScope
VFmp3scope1.ClearScope
VFmp3level1.ClearLevel
VFmp3level2.ClearLevel
VFmp3player1.Pause
End Sub
Private Sub Command4_Click()
VFmp3scope2.ClearScope
VFmp3scope1.ClearScope
VFmp3level1.ClearLevel
VFmp3level2.ClearLevel
a = "D:\Моя музыка\Сборник 1\BLADE.mp3"
If a <> "" Then
VFmp3player1.SongName = a
End If
End Sub
Private Sub Form_Load()
Slider1.Left = 0
Slider1.Width = Form1.Width - 100
mymousedown = False
Slider2.Value = VFmp3player1.GetVolume
End Sub
Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
mymousedown = True
End Sub
Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    VFmp3player1.Seek (Slider1.Value)
    mymousedown = False
End Sub
Private Sub Text11_Change()
VFmp3player1.SetVolume (Text11.Text)
End Sub
Private Sub Slider2_Change()
VFmp3player1.SetVolume (Slider2.Value)
End Sub
Private Sub VFmp3player1_ID3TagInfo(ByVal SongName As String, ByVal Artist As String, ByVal Album As String, ByVal Year As String, ByVal Comment As String, ByVal Genre As String)
Text4.Text = SongName
Text5.Text = Artist
Text6.Text = Album
Text7.Text = Year
Text8.Text = Comment
Text9.Text = Genre
End Sub

Private Sub VFmp3player1_MediaInfoTotalTime(ByVal Seconds As Long)
Text10.Text = Seconds
If Seconds > 0 Then
    Slider1.Max = Seconds
End If
End Sub

Private Sub VFmp3player1_MediaTimeInfo(ByVal TimerValue As Long)
If mymousedown = False Then
    Slider1.Value = TimerValue
End If
End Sub

Private Sub VFmp3player1_MPEGInfo(ByVal bitrate As Integer, ByVal samplerate As Long, ByVal totalframe As Integer)
Text1.Text = bitrate
Text2.Text = samplerate
Text3.Text = totalframe
End Sub

Private Sub VFmp3player1_PlayedSeconds(ByVal pseconds As Long)
If mymousedown = False Then
    Slider1.Value = pseconds
End If
End Sub

Private Sub VFmp3player1_PlayerStatus(ByVal Status As Integer)
If Status = 1 Then
Form1.Caption = "Stopped."
End If
If Status = 2 Then
Form1.Caption = "Playing ..."
End If

End Sub

Private Sub VFmp3player1_SongPlayed()
MsgBox "done ..."
End Sub

Private Sub VFmp3player1_VisDataChanged(ByVal VisData As Long)

VFmp3scope1.SetVisualData (VisData)
VFmp3scope2.SetVisualData (VisData)
VFmp3level1.SetVisualData (VisData)
VFmp3level2.SetVisualData (VisData)

'VFmp3spectrum1.SetVisualData (VisData)
End Sub

