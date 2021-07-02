Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, _
       ByVal x As Long, ByVal y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, _
     ByVal hSrcDC As Long, _
       ByVal xSrc As Long, ByVal ySrc As Long, _
     ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020

Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Global SoundBuffer As String

Sub BeginPlaySound(ByVal ResourceId As Integer)
    
    Dim Ret As Variant
    SoundBuffer = StrConv(LoadResData(ResourceId, "SOUND"), vbUnicode)
    Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    DoEvents
    
'60 the loud tick                                       (Counter of winning points)
'61 the sliding tick                                    (starting the wheels)
'62 dry pop (high pitch)                                (counter of the 4 starting points)
'63 slow pop (a book falling, low pitch or tone)        (stopping of wheels)
End Sub
