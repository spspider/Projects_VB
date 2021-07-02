VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAudioMon 
   BorderStyle     =   1  'Единственный Фиксированный
   ClientHeight    =   2775
   ClientLeft      =   4515
   ClientTop       =   2310
   ClientWidth     =   5295
   Icon            =   "FrmAudioMon.frx":0000
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   353
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2505
      Left            =   4560
      TabIndex        =   7
      Top             =   150
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   4419
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   120
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Горизонтальный
      TabIndex        =   6
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Timer t1 
      Interval        =   50
      Left            =   210
      Top             =   1320
   End
   Begin VB.PictureBox PicDisplay 
      Appearance      =   0  'Плоский
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Ничего
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   300
      LinkTimeout     =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Пиксель
      ScaleWidth      =   256
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3840
   End
   Begin VB.PictureBox PicFrame 
      Appearance      =   0  'Плоский
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Ничего
      ForeColor       =   &H00000000&
      Height          =   1080
      Left            =   120
      LinkTimeout     =   0
      Picture         =   "FrmAudioMon.frx":0442
      ScaleHeight     =   72
      ScaleMode       =   3  'Пиксель
      ScaleWidth      =   281
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox PicOscBB 
      Appearance      =   0  'Плоский
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Ничего
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   300
      LinkTimeout     =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Пиксель
      ScaleWidth      =   256
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1335
      Width           =   1035
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "&Monitor"
      Height          =   375
      Left            =   2220
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Label ll1 
      BorderStyle     =   1  'Единственный Фиксированный
      Height          =   435
      Left            =   870
      TabIndex        =   5
      Top             =   1320
      Width           =   1245
   End
End
Attribute VB_Name = "FrmAudioMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

    Caption = TTL

   ' Width = 4545
   ' Height = 2205
    'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 3

End Sub
Private Sub CmdMonitor_Click()

    Dim Rv&
    Dim Msg$
    Dim WF As WAVEFORMATEX

    With WF
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = 1
        .nBlockAlign = 1
        .nSamplesPerSec = 11025
        .wBitsPerSample = 8
        .nAvgBytesPerSec = (.nSamplesPerSec * .nBlockAlign) \ 8
        .cbSize = 0
    End With

    Rv = waveInOpen(hWaveIn, WAVE_MAPPER, WF, AddressOf waveInProc, 0, CALLBACK_FUNCTION)
    If Rv <> 0 Then
       Msg = "Unable to open wave input device."
       MsgBox Msg, vbCritical, TTL & " - Error"
       Exit Sub
    End If

    CmdMonitor.Enabled = False
    CmdStop.Enabled = True
    DoEvents

    MonitorAudio

End Sub
Private Sub CmdStop_Click()

    CmdStop.Enabled = False
    CmdMonitor.Enabled = True

    waveInReset hWaveIn
    waveInStop hWaveIn
    waveInClose hWaveIn
    hWaveIn = 0

    DoEvents

End Sub
Private Sub Form_QueryUnload(Cancel%, UnloadMode%)

    If CmdStop.Enabled Then CmdStop_Click
    End

End Sub

Private Sub ll1_Change()

'If ll1.Caption >= 68 Or ll1.Caption <= 63 Then
'MsgBox "qqq"
'End If
txt1.Text = txt1.Text & " * " & ll1.Caption

End Sub

