VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDXVBFPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FPS As Long
Private freqPerformance, firstTick, freqAdjust As Long

Public Sub GetFPS()
      Dim diffTick As Long
      Static FpsTime, lastTick As Long
      Static FramesRendered As Long
      
      Call GetTick
      
      diffTick = firstTick - lastTick
      lastTick = firstTick
      
      FpsTime = FpsTime + diffTick
      FramesRendered = FramesRendered + 1
      
      If FpsTime > ((1000 * freqPerformance) / freqAdjust) Then
            FPS = FramesRendered
            FramesRendered = 0
            FpsTime = 0
      End If
End Sub

Private Sub GetTick()
      firstTick = timeGetTime()
      freqAdjust = 1
      freqPerformance = 1
End Sub

Public Function ReturnFPS() As Long
      ReturnFPS = FPS
End Function
