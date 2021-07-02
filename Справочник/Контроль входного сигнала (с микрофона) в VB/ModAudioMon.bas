Attribute VB_Name = "ModAudioMon"
Option Explicit

Public hWaveIn&

Private bBufferFull As Boolean
Private WaveData(255) As Byte

Private ColorIdx%
Private CurrentColor&
Private ColorCycle&(1275)

Type WAVEFORMATEX
     wFormatTag As Integer
     nChannels As Integer
     nSamplesPerSec As Long
     nAvgBytesPerSec As Long
     nBlockAlign As Integer
     wBitsPerSample As Integer
     cbSize As Integer
End Type

Type WAVEHDR
     lpData As Long
     dwBufferLength As Long
     dwBytesRecorded As Long
     dwUser As Long
     dwFlags As Long
     dwLoops As Long
     lpNext As Long
     Reserved As Long
End Type

Type WAVEINCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * 32
     dwFormats As Long
     wChannels As Integer
End Type

Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)

Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

Public Const CALLBACK_FUNCTION = &H30000
Public Const MM_WIM_DATA = &H3C0
Public Const WAVE_FORMAT_PCM = 1
Public Const WAVE_MAPPER = -1&

Public Const WAVE_FORMAT_1M08 = &H1     ' 11025 Hz, 8 Bit, Mono.
Public Const WAVE_FORMAT_2M08 = &H10    ' 22050 Hz, 8 Bit, Mono.
Public Const WAVE_FORMAT_4M08 = &H100   ' 44100 Hz, 8 Bit, Mono.

Public Const TTL = "Audio Monitor"
Private Sub Main()

    #If Win32 Then
        If PrevInst Then End
        If Not WaveInPresent Then End
        GetColors
        FrmAudioMon.Show
    #Else
        End
    #End If

End Sub
Private Function PrevInst() As Boolean

    Dim Msg$

    If App.PrevInstance Then
       Msg = TTL & " is already running."
       MsgBox Msg, vbInformation, TTL
       PrevInst = True
    Else
       PrevInst = False
    End If

End Function
Private Function WaveInPresent() As Boolean

    Dim Msg$

    If waveInGetNumDevs() > 0 Then
       WaveInPresent = True
    Else
       Msg = "Unable to detect any wave input devices."
       MsgBox Msg, vbCritical, TTL & " - Error"
       WaveInPresent = False
    End If

End Function
Private Sub GetColors()

    Dim K%

    For K = 0 To 1275
        ColorCycle(K) = Val(LoadResString((K + 1) * 10))
    Next

    CurrentColor = ColorCycle(0)

End Sub
Public Sub waveInProc(ByVal hwi&, ByVal uMsg&, ByVal dwInstance&, ByVal dwParam1&, ByVal dwParam2&)

    ' ===========================================================================
    ' <<< WARNING >>>
    ' Do not call any system-defined functions from inside This callback Routine.
    ' Calling other wave functions will cause deadlock.
    ' ===========================================================================

    If uMsg = MM_WIM_DATA Then bBufferFull = True

End Sub
Public Sub MonitorAudio()

    Dim WH As WAVEHDR

    waveInStart hWaveIn

    Do
      With WH
          .lpData = VarPtr(WaveData(0))
          .dwBufferLength = 256
          .dwFlags = 0
      End With

      waveInPrepareHeader hWaveIn, WH, Len(WH)
      bBufferFull = False
      waveInAddBuffer hWaveIn, WH, Len(WH)

      Do
        DoEvents
      Loop Until bBufferFull Or hWaveIn = 0

      waveInUnprepareHeader hWaveIn, WH, Len(WH)
      DrawData

      DoEvents
    Loop Until hWaveIn = 0

    FrmAudioMon.PicDisplay.Cls

End Sub
Private Sub DrawData()
Dim y
    Dim x%

    FrmAudioMon.PicOscBB.Cls

    For x = 0 To 254
        FrmAudioMon.PicOscBB.Line (x, WaveData(x) \ 4)-(x + 1, WaveData(x + 1) \ 4), CurrentColor
   
    Next

    BitBlt FrmAudioMon.PicDisplay.hDC, 0, 0, 256, 64, FrmAudioMon.PicOscBB.hDC, 0, 0, vbSrcCopy
'********************
 FrmAudioMon.ll1.Caption = (WaveData(x) \ 4) + (WaveData(x - 1) \ 4)
 FrmAudioMon.ProgressBar1.Value = (WaveData(x) \ 4) + (WaveData(x - 1) \ 4)

If FrmAudioMon.ProgressBar1.Value >= 80 Or FrmAudioMon.ProgressBar1.Value <= 40 Then MsgBox "aaa"

'*********************************
    ColorIdx = ColorIdx + 1
    If ColorIdx > 1275 Then ColorIdx = 0
    CurrentColor = ColorCycle(ColorIdx)

End Sub
