Attribute VB_Name = "Module1"
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'''''''''''''
Public Const SRCCOPY = &HCC0020
''''''''''''''''''
Public Type POINTAPI

    x As Long

    y As Long

End Type
''''''''''
Public auto As POINTAPI
''''''''''






