Attribute VB_Name = "Module1"
Public Declare Function SetPixel& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long)
Public Declare Function GetPixel& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long)
Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Public Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)

