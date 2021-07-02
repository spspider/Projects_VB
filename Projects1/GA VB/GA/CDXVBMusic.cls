VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDXVBMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WORKING!

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public m_HWND As Long

Public Sub Init(hWnd As Long)
      m_HWND = hWnd
End Sub

Public Function Play(Filename As String) As Boolean
      If Filename = "" Then Play = False
      Dim Buffer As String
      
      Buffer = "open " & Filename & " type sequencer alias MUSIC"
      
      If mciSendString("close all", "", 0, 0) <> 0 Then
            Play = False
            Exit Function
      End If
      
      If mciSendString(Buffer, "", 0, 0) <> 0 Then
            Play = False
            Exit Function
      End If
      
      If mciSendString("play MUSIC from 0 notify", "", 0, m_HWND) <> 0 Then
            Play = False
            Exit Function
      End If
      
      Play = True
End Function

Public Function StopPlaying() As Boolean
      If mciSendString("close all", "", 0, 0) <> 0 Then
            StopPlaying = False
            Exit Function
      End If
      
      StopPlaying = True
End Function

Public Function Pause() As Boolean
      If mciSendString("stop MUSIC", "", 0, 0) <> 0 Then
            Pause = False
            Exit Function
      End If
      
      Pause = True
End Function

Public Function Restart() As Boolean
      If mciSendString("play MUSIC from 0 notify", "", 0, m_HWND) <> 0 Then
            Restart = False
            Exit Function
      End If
      
      Restart = True
End Function
