VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Winamp Control"
   ClientHeight    =   4620
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTitleScroll 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4335
      Top             =   2895
   End
   Begin ComctlLib.Slider sldTime 
      Height          =   390
      Left            =   165
      TabIndex        =   20
      Top             =   1590
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   688
      _Version        =   327682
      TickStyle       =   3
      TickFrequency   =   0
   End
   Begin ComctlLib.Slider sldVol 
      Height          =   255
      Left            =   6765
      TabIndex        =   16
      Top             =   1230
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   6135
      Top             =   1065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Next Track"
      Height          =   375
      Index           =   3
      Left            =   4485
      TabIndex        =   4
      Top             =   1110
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Pause"
      Height          =   375
      Index           =   4
      Left            =   3525
      TabIndex        =   3
      Top             =   1110
      Width           =   975
   End
   Begin VB.CommandButton cmdRecreatePlaylist 
      Caption         =   "Recreate Playlist"
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   3990
      Width           =   2055
   End
   Begin VB.CommandButton cmdMDLoad 
      Caption         =   "MD Load"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3990
      Width           =   1575
   End
   Begin VB.Timer tmrMDLoadInterval 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4335
      Top             =   3255
   End
   Begin VB.CommandButton cmdMDLoadStop 
      Caption         =   "Stop MD Load"
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   3990
      Width           =   1575
   End
   Begin VB.Frame fraRemChat 
      Caption         =   "Chat"
      Height          =   1650
      Left            =   4770
      TabIndex        =   10
      Top             =   2250
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtRemChatSend 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1290
         Width           =   2775
      End
      Begin VB.TextBox txtRemChat 
         Height          =   1050
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   195
         Width           =   2745
      End
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Remote Host Status"
      Height          =   1320
      Left            =   105
      TabIndex        =   5
      Top             =   2250
      Width           =   4215
      Begin VB.CommandButton cmdRemChat 
         Caption         =   "Chat with User"
         Height          =   375
         Left            =   1380
         TabIndex        =   7
         Top             =   810
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemKick 
         Caption         =   "Kick User"
         Height          =   375
         Left            =   2940
         TabIndex        =   8
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lblRemStatus 
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblRemInfo 
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   495
         Width           =   3735
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   500
      Left            =   4335
      Top             =   2535
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   4335
      Top             =   2175
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9998
      LocalPort       =   2346
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Play"
      Height          =   375
      Index           =   0
      Left            =   2685
      TabIndex        =   2
      Top             =   1110
      Width           =   855
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Stop"
      Height          =   375
      Index           =   1
      Left            =   1965
      TabIndex        =   1
      Top             =   1110
      Width           =   735
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Previous Track"
      Height          =   375
      Index           =   2
      Left            =   525
      TabIndex        =   0
      Top             =   1110
      Width           =   1455
   End
   Begin ComctlLib.Slider sldBal 
      Height          =   255
      Left            =   6765
      TabIndex        =   18
      Top             =   1830
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   7875
      TabIndex        =   21
      Top             =   0
      Width           =   7875
      Begin VB.Label lblShuffle 
         BackColor       =   &H00000000&
         Caption         =   "SHUFFLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblRepeat 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "REPEAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTrackName 
         BackColor       =   &H00000000&
         Caption         =   "ABCDEFGHIJKLMNOPQRSTUVWXYZABC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00000000&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Balance"
      Height          =   255
      Left            =   6885
      TabIndex        =   19
      Top             =   1575
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Volume"
      Height          =   375
      Left            =   6885
      TabIndex        =   17
      Top             =   990
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Preferences 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuFile_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuHidden_do_aa 
         Caption         =   "--Hidden--activate app"
         Shortcut        =   +^{F3}
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All variables must be declared!
'************************************************
'* This form is used for the Winamp Control     *
'* Module example, but is not necessary for use *
'* of the Winamp control module. The winamp     *
'* control module (WinampModule) can operate    *
'* by itself in any other application.          *
'************************************************


Dim mdLoadstop As Boolean
Dim chatmode As Boolean
Dim waitDone As Boolean
Dim curtime As Long
Dim tlen As Integer
Public TrackNameMod As String
Public allowkill As Boolean
Public HideScroll As Boolean
Dim lastTimeDisp As Long
Dim repeatStat As Integer
Dim shuffleStat As Integer
Dim remLastTitle As String
Dim remLastnum As Long
Dim remLastTime As String
Dim remMode As Byte







Public Sub cmdAction_Click(index As Integer)
On Error GoTo err
Select Case index
    Case 0
        WM_DO (PLAY_TRACK)
    Case 1
        WM_DO (STOP_TRACK)
    Case 2
        WM_DO (PREV_TRACK)
    Case 3
        WM_DO (NEXT_TRACK)
    Case 4
        WM_DO (PAUSE_TRACK)
End Select
TrackNameMod = WM_GET(SONG_TITLE)
tlen = WM_GET(SONG_LENGTH)
err:
End Sub





Private Sub cmdMDLoadStop_Click()
mdLoadstop = True
End Sub

Private Sub cmdRecreatePlaylist_Click()
Dim thisentry As String, curnum As Integer
curnum = 1
WM_GET (CLEAR_PLAYLIST)
Do
    thisentry = Parse(recreate(1), curnum)
    If thisentry <> "" Then
       ' WinampPlaylistAdd thisentry
    End If
    curnum = curnum + 1
Loop Until thisentry = ""
End Sub

Private Sub cmdRemChat_Click()
On Error Resume Next
chatmode = Not chatmode
fraRemChat.Visible = Not fraRemChat.Visible
If chatmode = False Then
    cmdRemChat.Caption = "enable chat mode"
    dS vbNewLine + String(30, "*") + vbNewLine + "---------CHAT MODE OFF--------" + vbNewLine + String(30, "*") + vbNewLine
    TelnetSendIface
    UpdateRemCaption
Else
    cmdRemChat.Caption = "disable chat mode"
    dS Chr(27) + "[2J" 'clear screen
    dS vbNewLine + String(30, "*") + vbNewLine + "----------CHAT  MODE----------" + vbNewLine + String(30, "*") + vbNewLine
    dS "You are now in chat mode. All normal mode commands are disabled. " + vbNewLine + "You can now chat with the host" + vbNewLine
End If
End Sub

Private Sub cmdRemKick_Click()
On Error Resume Next
dS vbNewLine + vbNewLine + vbNewLine + "The server has disconnected you"
dS "bye!"
DoEvents: DoEvents
lblRemStatus.Caption = "User kicked"
TelnetReset
End Sub

Private Sub cmdMDLoad_Click()

Dim totalnum As Integer, curnum As Integer
Dim dafile As String, mm As Integer
'cdlopen.ShowOpen
'dafile = cdlopen.filename
mdLoadstop = False
WM_DO (STOP_TRACK)
WM_DO (CLEAR_PLAYLIST)
'MsgBox "T1"
'W.AddFile dafile
totalnum = WM_GET(PLAYLIST_LENGTH)
curnum = 1
mm = MsgBox("Click OK when ready to begin load of " + dafile, vbOKCancel)
If mm = vbCancel Then GoTo noDo
Do Until mdLoadstop = True Or curnum = totalnum
    WM_DO (PLAY_TRACK)
    Do Until Int(WM_GET(SONG_POSITION) / 1000) = WM_GET(SONG_LENGTH)
        DoEvents
    Loop
    WM_DO (STOP_TRACK)
    WM_DO (NEXT_TRACK)
    WM_DO (STOP_TRACK)
    'MsgBox "HI@"
    waitDone = False
    tmrMDLoadInterval.Enabled = True
    Do Until waitDone = True
        DoEvents
    Loop
    tmrMDLoadInterval.Enabled = False
    curnum = curnum + 1
Loop
noDo:
End Sub

Sub TelnetSendIface()
dS Chr(27) + "[2J"  'clear screen
Dim bleh As String
Open App.Path + "\txtSend.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, bleh
    dS bleh + vbNewLine
Loop
Close #1
End Sub
Sub TelnetSendTrackNum(Optional TheTrackNum As Long = -2)
    If TheTrackNum = -2 Then
        TheTrackNum = CLng(WM_GET(PLAYLIST_POSITION) + 1)
    End If
    dS Chr(27) + "[6;17H"
    dS CStr(TheTrackNum)
End Sub




Private Sub Form_Load()
WinampPath = d.winamp_path
FindWinamp 'Calls the findwinamp function
If d.telnet_enable = 1 Then
    TelnetReset
End If
presets(1) = "G:\sound\Electronic.m3u"
presets(2) = "G:\sound\rock.m3u"
presets(3) = "G:\sound\Main.m3u"
recreate(1) = "G:\sound\Drumnbass G:\sound\Electronic G:\sound\Techno G:\sound\Trance"
chatmode = False
tmrTime_Timer
tmrTitleScroll.Interval = d.title_scroll_speed
sldTime.Visible = True
SaveSetting "WinAmpControl", "main", "PID", Str(Me.hwnd)
lblRemStatus.Caption = "Waiting for connection"
shuffleStat = WM_GetShufflestatus
repeatStat = WM_GetRepeatStatus
'Show
updatecolors
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdLoadstop = True
waitDone = True
If allowkill = False Then
    End
Else
    'Unload Me
End If
End Sub


Private Sub lblRepeat_Click()
If WM_GetRepeatStatus <> repeatStat Then
    repeatStat = 1 - repeatStat
Else
    WM_DO TOGGLE_REPEAT
    repeatStat = 1 - repeatStat
End If
updatecolors
End Sub

Private Sub lblShuffle_Click()
If WM_GetShufflestatus <> shuffleStat Then
    shuffleStat = 1 - shuffleStat
Else
    WM_DO TOGGLE_SHUFFLE
    shuffleStat = 1 - shuffleStat
End If
updatecolors
End Sub


Private Sub mnuFile_Exit_Click()
mdLoadstop = True
End
End Sub




Private Sub mnuFile_Preferences_Click()
frmOptions.Show
End Sub



Private Sub mnuHidden_do_aa_Click()
Me.Show
Me.SetFocus
Me.WindowState = 0
End Sub

Private Sub sldBal_Change()
Call sldBal_Scroll
End Sub


Private Sub sldBal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
HideScroll = True
End Sub

Private Sub sldBal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
HideScroll = False
currentTrackName = ""
End Sub

Private Sub sldBal_Scroll()
WM_SetPanning sldBal.Value
Dim prcnt As Integer, lrstatus As String
prcnt = Int((sldBal.Value - 127) / 1.27)
If prcnt = 0 Then
    lrstatus = " CENTER"
ElseIf prcnt < 0 Then
    lrstatus = Str(Abs(prcnt)) + "% LEFT"
Else
    lrstatus = Str(prcnt) + "% RIGHT"
End If
lblTrackName.Caption = "BALANCE:" + lrstatus
End Sub

Private Sub sldVol_Change()
Call sldVol_Scroll
End Sub

Private Sub sldVol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
HideScroll = True
End Sub

Private Sub sldVol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
restorecaption
HideScroll = False
currentTrackName = ""
End Sub

Private Sub sldVol_Scroll()
lblTrackName.Caption = "VOLUME:" + Str(Int(sldVol.Value / 2.55)) + "%"
WM_SetVolume sldVol.Value
End Sub








Private Sub sldTime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'tmrTime.Enabled = False
AllowTimeChange = True
HideScroll = True
End Sub

Private Sub sldTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'tmrTime.Enabled = true
AllowTimeChange = False
Call WM_SetSongPosition(sldTime.Value)
HideScroll = False
currentTrackName = ""
End Sub

Private Sub sldTime_Scroll()
'If Check1.Value = 1 Then
    'WM_SetSongPosition (sldTime.Value)
    lblTrackName.Caption = "SEEK TO: " + FormatTime(sldTime.Value) + "/" + FormatTime(tlen) + " (" + Mid(Str(Int(sldTime.Value / tlen * 100)), 2) + "%)"
'End If
End Sub

Private Sub tmrTitleScroll_Timer()
If HideScroll = False And Len(TrackNameMod) > TrackBoxLen Then
    Dim a As String
    a = TrackNameMod
    TrackNameMod = Mid(a, 2) + Mid(a, 1, 1)
    'Label3.Caption = tracknamemod
    lblTrackName.Caption = Mid(TrackNameMod, 1, TrackBoxLen)
    If d.title_caps = 1 Then
        lblTrackName.Caption = UCase(lblTrackName.Caption)
    End If
End If
End Sub

Private Sub tmrMDLoadInterval_Timer()
waitDone = True
End Sub



Private Sub tmrTime_Timer()
On Error Resume Next
If WM_GET(SONG_POSITION) <> -1 Then
    Dim t As Integer, c As String, c2 As String
    t = Int(WM_GET(SONG_POSITION) / 1000)
    c = FormatTime(t)
    If wsMain.State <> sckClosed Then
        If c <> remLastTime And chatmode = False And d.telnet_enable = 1 Then
            dS Chr(27) + "[8;9H"
            dS c
            remLastTime = c
            dS String(3, " ")
            dS Chr(27) + "[17;1H"
        End If
    End If
    
    c2 = WM_GET(SONG_TITLE)
    
    If curtime > t Or c2 <> currentTrackName Then
        sldTime.Max = WM_GET(SONG_LENGTH)
        currentTrackName = c2
        TrackNameMod = currentTrackName
        tlen = WM_GET(SONG_LENGTH)
        If Len(TrackNameMod) > TrackBoxLen Then
            TrackNameMod = " " + TrackNameMod + "  <<<>>> "
            tmrTitleScroll.Enabled = True
        Else
            tmrTitleScroll.Enabled = False
            restorecaption
        End If
        If d.telnet_enable = 1 Then
            UpdateRemCaption
        End If
    End If
    curtime = t
    If AllowTimeChange = False Then
        lblTime.Caption = c + "/" + FormatTime(tlen)
        sldTime.Value = t
    End If
End If
End Sub






Private Sub txtRemChatSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    dS vbNewLine + "Host sent message: " + txtRemChatSend.Text + vbNewLine
    txtRemChatSend.Text = ""
End If
txtRemChat.Text = txtRemChat.Text + vbNewLine
End Sub

Private Sub wsMain_Close()
lblRemStatus.Caption = "user disconnnected"
TelnetReset
End Sub

Private Sub wsMain_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
If wsMain.State <> sckClosed Then wsMain.Close
    remMode = 1
    If chatmode = True Then
        Call cmdRemChat_Click
    End If
    wsMain.Accept requestID
    dS Chr(27) + "[46;m"
    TelnetSendIface
    dS Chr(27) + "[31;m"
    lblRemInfo.Caption = "Remote IP: " & wsMain.RemoteHostIP
    dS "You are now connected to the system.  "
    
    'Sending version string
    dS Chr(27) + "[15;12H"
    dS CStr(App.Major) + "." + Format(App.Minor, "00") & "." + Format(App.Revision, "00")
    
    currentTrackName = ""
    'Command3_
    TelnetSendTrackNum
    
'    ds "Commands:" + vbNewLine + "p: Plays" + vbNewLine + _
'     "s: Stop" + vbNewLine + "n : next track" + vbNewLine + _
'     "r: Previous track" + vbNewLine + "x: exit" + vbNewLine + _
'     "t: Enable/Disable time display" + vbNewLine + "1-5: Preset Playlists" + vbNewLine + vbNewLine
'    ds "" + vbNewLine + "current track: " + Str(WM_GET(PLAYLIST_POSITION) + 1) + ":  " + WM_GET(SONG_TITLE) + vbNewLine
    lblRemStatus.Caption = "User Connected"

    updatecolors
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim n As Boolean
Dim strdata As String
wsMain.GetData strdata
If chatmode = False Then
    Select Case LCase(strdata)
        Case "t"
            timedisp = Not timedisp
            dS vbNewLine + vbNewLine + String$(40, "*") + vbNewLine + "  Time Display " + _
            returnED(timedisp) + vbNewLine + String$(40, "*") + vbNewLine + vbNewLine
            n = True
        Case "v"
            lblRemStatus.Caption = "User stopped player"
            WM_DO STOP_TRACK
        Case "x"
            lblRemStatus.Caption = "User began play"
            WM_DO PLAY_TRACK
        Case "b", Chr(27) + "OC"
            lblRemStatus.Caption = "User changed to next track"
            WM_DO NEXT_TRACK
        Case "z", Chr(27) + "OD"
            lblRemStatus.Caption = "User changed to previous track"
            WM_DO PREV_TRACK
        Case "q"
            dS Chr(27) + "[2J"
            dS vbNewLine + "Bye" + vbNewLine
            lblRemStatus.Caption = "User disconnected"
            DoEvents: DoEvents: DoEvents
            TelnetReset
        Case "c"
            lblRemStatus.Caption = "User paused playback"
            WM_DO PAUSE_TRACK
        Case "s"
            lblShuffle_Click
        Case "r"
            lblRepeat_Click
        Case Else
            Dim wz As Integer
            wz = Val(strdata)
            If wz < 6 And wz > 0 Then
                WM_GET (CLEAR_PLAYLIST)
                'W.AddFile presets(wz)
                Dim b
                dS vbNewLine + vbNewLine + String$(40, "*") + vbNewLine + "Loaded playlist: " + presets(wz) + vbNewLine + String$(40, "*") + vbNewLine + vbNewLine
            Else
                lblRemStatus.Caption = "User queried status"
                'Debug.Print strdata
                'ds Str(WM_GET(PLAYLIST_POSITION) + 1) + ": " + WM_GET(SONG_TITLE) + "--" + lblTime.Caption + vbNewLine
                n = True
            End If
    End Select
    If n <> True Then
    
    End If
Else
    dS strdata
    If strdata = Bs Then
        dS " " + Bs
        txtRemChat.Text = Mid(txtRemChat.Text, 1, Len(txtRemChat.Text) - 1)
    Else
        txtRemChat.Text = txtRemChat.Text & strdata
    End If
End If
End Sub

Sub restorecaption()

    If Len(currentTrackName) < TrackBoxLen Then
        lblTrackName.Caption = currentTrackName
        If d.title_caps = 1 Then
            lblTrackName.Caption = UCase(lblTrackName.Caption)
        End If
    End If
End Sub

Sub updatecolors()
On Error Resume Next
dS Chr(27) + "[6;45H"
If repeatStat = 1 Then
    lblRepeat.ForeColor = &HC000&
    dS "X"
Else
    lblRepeat.ForeColor = &H4000&
    dS " "
End If
dS Chr(27) + "[6;33H"
If shuffleStat = 1 Then
    lblShuffle.ForeColor = &HC000&
    dS "X"
Else
    lblShuffle.ForeColor = &H4000&
    dS " "
End If
End Sub

Sub UpdateRemCaption()

        'ds Trim(Str(WM_GET(PLAYLIST_POSITION) + 1)) & ":  " & WM_GET(SONG_TITLE) & vbNewLine
            If Len(remLastTitle) > Len(currentTrackName) Then
                dS Chr(27) + "[4;" + Trim(Str(15 + Len(currentTrackName))) + "H"
                dS String(Len(remLastTitle) - Len(currentTrackName), " ")
            End If
            remLastTitle = currentTrackName
            dS Chr(27) + "[4;15H" 'positions cursor
            If Len(remLastTitle) > 61 Then remLastTitle = Mid(remLastTitle, 1, 61) 'makes sure title does not overflow
            dS remLastTitle
Dim L As Long
L = CLng(WM_GET(PLAYLIST_POSITION) + 1)
    If remLastnum <> L Then
        remLastnum = L
        TelnetSendTrackNum (L)
    End If
dS Chr(27) + "[17;1H"
End Sub


Sub dS(theData As String)
    On Error Resume Next
    wsMain.SendData theData
End Sub

Sub TelnetReset()
    On Error Resume Next
    wsMain.Close
    If d.telnet_enable = 1 Then
        wsMain.LocalPort = d.telnet_port
        wsMain.Listen
    End If
End Sub
