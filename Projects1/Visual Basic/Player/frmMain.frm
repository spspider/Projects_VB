VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectSound Player"
   ClientHeight    =   5340
   ClientLeft      =   1800
   ClientTop       =   2220
   ClientWidth     =   7515
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   5340
   ScaleWidth      =   7515
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop"
      Height          =   255
      Left            =   6480
      TabIndex        =   34
      Top             =   1440
      Width           =   735
   End
   Begin MSComctlLib.Slider sldPosition 
      Height          =   375
      Left            =   1560
      TabIndex        =   32
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":0884
      LargeChange     =   0
      SmallChange     =   0
      Min             =   1
      Max             =   10000
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Frame frmAdvanced 
      Caption         =   "Advanced"
      Height          =   2535
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdSimple 
         Caption         =   "<< Simple"
         Height          =   375
         Left            =   5640
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame frmPControls 
         Caption         =   "Playback Controls"
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6615
         Begin MSComctlLib.Slider sldPan 
            Height          =   375
            Left            =   2640
            TabIndex        =   21
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            MousePointer    =   99
            MouseIcon       =   "frmMain.frx":0B9E
            LargeChange     =   1000
            SmallChange     =   100
            Min             =   -10000
            Max             =   10000
            TickStyle       =   3
         End
         Begin MSComctlLib.Slider sldFreq 
            Height          =   375
            Left            =   2640
            TabIndex        =   22
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            MousePointer    =   99
            MouseIcon       =   "frmMain.frx":0EB8
            LargeChange     =   1000
            SmallChange     =   100
            Min             =   100
            Max             =   100000
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "100 KHz"
            Height          =   195
            Index           =   5
            Left            =   5640
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
            Height          =   195
            Index           =   7
            Left            =   5700
            TabIndex        =   30
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "100 Hz"
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   29
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
            Height          =   195
            Index           =   6
            Left            =   2040
            TabIndex        =   28
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pan"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblFrequency 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblPan 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame frmSettings 
         Caption         =   "Buffer Settings"
         Height          =   555
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   6615
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   1260
            ScaleHeight     =   255
            ScaleWidth      =   5235
            TabIndex        =   15
            Top             =   180
            Width           =   5235
            Begin VB.OptionButton optGlobal 
               Caption         =   "Global"
               Height          =   195
               Left            =   2400
               TabIndex        =   18
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton optSticky 
               Caption         =   "Sticky"
               Height          =   195
               Left            =   1200
               TabIndex        =   17
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton optNormal 
               Caption         =   "Normal"
               Height          =   195
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Focus"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced >>"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   7440
      Top             =   120
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":11D2
      LargeChange     =   1000
      SmallChange     =   100
      Min             =   -2500
      Max             =   0
      TickStyle       =   3
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Min"
      Height          =   255
      Left            =   3000
      TabIndex        =   36
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Db * 100"
      Height          =   255
      Left            =   2280
      TabIndex        =   35
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   " Play position (bytes)"
      Height          =   495
      Left            =   600
      TabIndex        =   33
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblVolume 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 2002 Tchouprina Andrew. All Rights Reserved.
'
'  File:       frmMain.frm
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private dx As New DirectX8
Private ds As DirectSound8
Private dsb As DirectSoundSecondaryBuffer8
Private msFile As String
Private lPausedPos As Long
Private lBufferSize As Long
Private bCanUpdPosition As Long

Private Sub cmdAdvanced_Click()
    frmAdvanced.Visible = True
    cmdAdvanced.Visible = False
End Sub
Private Sub cmdOpen_Click()

    Static sCurDir As String
    Static lFilter As Long
    Dim dsBuf As DSBUFFERDESC
    Dim dsBCap As DSBCAPS
    
    UpdateStatus "Loading file..."
    With cdlFile
        .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .FilterIndex = lFilter
        .Filter = "Wave Files (*.wav)|*.wav"
        .FileName = vbNullString
        If sCurDir = vbNullString Then
            Dim sWindir As String
            sWindir = Space$(255)
            If GetWindowsDirectory(sWindir, 255) = 0 Then
                .InitDir = "C:\"
            Else
                Dim sMedia As String
                sWindir = Left$(sWindir, InStr(sWindir, Chr$(0)) - 1)
                If Right$(sWindir, 1) = "\" Then
                    sMedia = sWindir & "Media"
                Else
                    sMedia = sWindir & "\Media"
                End If
                If Dir$(sMedia, vbDirectory) <> vbNullString Then
                    .InitDir = sMedia
                Else
                    .InitDir = sWindir
                End If
            End If
        Else
            .InitDir = sCurDir
        End If
        .ShowOpen
        If .FileName = vbNullString Then
            UpdateStatus "No file loaded."
            Exit Sub
        End If
        sldPosition.Enabled = True

        sCurDir = GetFolder(.FileName)
        lFilter = .FilterIndex
        UpdateStatus "File loaded."
        
        msFile = .FileName
        If Not (dsb Is Nothing) Then dsb.Stop
        Set dsb = Nothing
        dsBuf.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
        On Error Resume Next
        Set dsb = ds.CreateSoundBufferFromFile(msFile, dsBuf)
        If Err Then
            UpdateStatus "Could not create buffer."
            Exit Sub
        End If
        
        lPausedPos = 0
        dsb.GetCaps dsBCap
        lBufferSize = dsBCap.lBufferBytes
        
        sldPosition.Min = 0
        sldPosition.Max = lBufferSize - 1
        
        sldFreq.Value = dsBuf.fxFormat.lSamplesPerSec
        lblFile.Caption = .FileName
        EnablePlayUI True
    End With
    
End Sub
Private Sub UpdateStatus(ByVal sStatus As String)
    lblStatus.Caption = sStatus
End Sub
Private Function GetFolder(ByVal sFile As String) As String
    Dim lCount As Long
    
    For lCount = Len(sFile) To 1 Step -1
        If Mid$(sFile, lCount, 1) = "\" Then
            GetFolder = Left$(sFile, lCount)
            Exit Function
        End If
    Next
    GetFolder = vbNullString
End Function
Private Sub EnablePlayUI(ByVal fEnable As Boolean)
    On Error Resume Next
    If fEnable Then
        cmdPlay.Visible = True
        cmdPause.Visible = False
        optNormal.Enabled = True
        optSticky.Enabled = True
        optGlobal.Enabled = True
        cmdOpen.Enabled = True
        cmdPlay.SetFocus
    Else
        cmdPlay.Visible = False
        cmdPause.Visible = True
        optNormal.Enabled = False
        optSticky.Enabled = False
        optGlobal.Enabled = False
        cmdOpen.Enabled = False
        cmdPause.SetFocus
    End If
End Sub
Private Function InitDSound() As Boolean
    On Error GoTo FailedInit
    InitDSound = True
    Set ds = dx.DirectSoundCreate(vbNullString)
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    Exit Function

FailedInit:
    InitDSound = False
End Function
Private Sub cmdPause_Click()
    If Not (dsb Is Nothing) Then
        Dim dsCur As DSCURSORS
        dsb.GetCurrentPosition dsCur
        lPausedPos = dsCur.lPlay
        dsb.Stop
        lblStatus.Caption = "File paused."
        EnablePlayUI True
    End If
End Sub
Private Sub cmdSimple_Click()
    cmdAdvanced.Visible = True
    frmAdvanced.Visible = False
End Sub
Private Sub cmdStop_Click()
    If Not (dsb Is Nothing) Then
        dsb.Stop
        lPausedPos = 0
        dsb.SetCurrentPosition lPausedPos
        sldPosition.Value = sldPosition.Min
        lblStatus.Caption = "File stopped."
        EnablePlayUI True
    End If
End Sub
Private Sub cmdExit_Click()
    Cleanup
    Unload Me
End Sub
Private Sub Form_Load()
    bCanUpdPosition = True
    sldPosition.Enabled = False
    
    If Not InitDSound Then
        MsgBox "Could not initialize DirectSound.  This sample is exiting.", vbOKOnly Or vbInformation, "Failed."
        Cleanup
        Unload Me
        End
    End If
    OnSliderChange
    UpdateStatus "No file loaded."
End Sub
Private Sub Cleanup()
    If Not (dsb Is Nothing) Then dsb.Stop
    Set dsb = Nothing
    Set ds = Nothing
    Set dx = Nothing
End Sub
Private Sub OnSliderChange()
    Dim lFrequency As Long, lPan As Long, lVolume As Long
    
    lFrequency = sldFreq.Value
    lPan = sldPan.Value
    lVolume = sldVolume.Value

    lblFrequency.Caption = CStr(lFrequency)
    lblPan.Caption = CStr(lPan)
    lblVolume.Caption = CStr(lVolume)

    If Not (dsb Is Nothing) Then
        dsb.SetFrequency lFrequency
        dsb.SetPan lPan
        dsb.SetVolume lVolume
    End If
End Sub
Private Sub cmdPlay_Click()
    Dim dsBuf As DSBUFFERDESC
    Dim bFocusSticky As Boolean, bFocusGlobal As Boolean
    Dim bMixHardware As Boolean, bMixSoftware As Boolean

    On Error GoTo ErrOut
    bFocusSticky = (optSticky.Value)
    bFocusGlobal = (optGlobal.Value)
        
    dsBuf.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    If bFocusGlobal Then
        dsBuf.lFlags = dsBuf.lFlags Or DSBCAPS_GLOBALFOCUS
    End If
    
    If bFocusSticky Then
        dsBuf.lFlags = dsBuf.lFlags Or DSBCAPS_STICKYFOCUS
    End If

    Set dsb = ds.CreateSoundBufferFromFile(msFile, dsBuf)
    OnSliderChange
    If chkLoop.Value = vbChecked Then
        dsb.Play DSBPLAY_LOOPING
    Else
        dsb.Play 0
    End If
    dsb.SetCurrentPosition (lPausedPos)
        
    lblStatus.Caption = "File playing."
    EnablePlayUI False
    Exit Sub
    
ErrOut:
    lblStatus.Caption = "An error occured trying to play this file with these settings."

End Sub

Private Sub sldFreq_Click()
    OnSliderChange
End Sub

Private Sub sldPan_Click()
    OnSliderChange
End Sub

Private Sub sldPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bCanUpdPosition = False
End Sub

Private Sub sldPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim a As Single
        
        bCanUpdPosition = True
        
        If x < 0 Then x = 1
        If x > sldPosition.Width Then x = sldPosition.Width
        
        a = (sldPosition.Max - sldPosition.Min) / sldPosition.Width
        lPausedPos = a * x
        sldPosition.Value = lPausedPos
        dsb.SetCurrentPosition lPausedPos
     End If
End Sub

Private Sub sldVolume_Click()
    OnSliderChange
End Sub

Private Sub tmrUpdate_Timer()
    If Not (dsb Is Nothing) Then
        If (dsb.GetStatus And DSBSTATUS_PLAYING) <> DSBSTATUS_PLAYING Then
            If cmdPlay.Visible = False Then
                lPausedPos = 0
                EnablePlayUI True
                lblStatus.Caption = "File stopped."
                
                sldPosition.Value = sldPosition.Min
            End If
        Else
            If bCanUpdPosition Then
                Dim dsCurPos As DSCURSORS
                Dim lPlayPos As Long
            
                dsb.GetCurrentPosition dsCurPos
                lPlayPos = dsCurPos.lPlay
                
                sldPosition.Value = (sldPosition.Max - sldPosition.Min) _
                    / lBufferSize * lPlayPos
            End If
        End If
    End If
End Sub

