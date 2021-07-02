Attribute VB_Name = "modMusic"
Option Explicit

Public Declare Function MSS Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Declare Function SetPixel& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long)
Declare Function GetPixel& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long)

Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As _
Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)

Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)

Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)


Function RZN(N1, N2)
Dim S As Single
S = N2 - N1
If S < 0 Then S = -S
RZN = S
End Function

Sub OpenFirst(FileName As String)
MSS "OPEN " & GetShortPath(FileName) & " ALIAS F", 0, 0, 0
End Sub

Sub PlayFirst()
    MSS "PLAY F", 0, 0, 0
End Sub

Sub PlayFrom(N)
MSS "PLAY F FROM " & N, 0, 0, 0
End Sub

Sub StopFirst()
    MSS "STOP F", 0, 0, 0
    End Sub

Sub MClose()
MSS "CLOSE F", 0, 0, 0
End Sub

Function GetFirstLen() As Long
    Dim GFL As String * 256
    MSS "STATUS F LENGTH", GFL, Len(GFL), 0
    GetFirstLen = Val(GFL)
End Function

Function GetFirstPos() As Long
    Dim GFP As String * 256
    MSS "STATUS F POSITION", GFP, Len(GFP), 0
    GetFirstPos = Val(GFP)
End Function

Public Function GetShortPath(strFileName As String) As String
    Dim strPath As String
    strPath = Space(265)
    GetShortPath = Left(strPath, GetShortPathName(strFileName, strPath, 264))
End Function





