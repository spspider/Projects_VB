Attribute VB_Name = "Модуль1"
'Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Type POINTAPI
         x As Long
         y As Long
End Type
'Public Type LOGBRUSH
'         lbStyle As Long
'         lbColor As Long
'         lbHatch As Long
'End Type
Public Type Pos
    x As Single
    y As Single
    z As Single
End Type






Public Sub RGB1024(color As Single, ByRef cr As Integer, ByRef cg As Integer, ByRef cb As Integer)
    Select Case color * 1024
        Case Is <= 256
            cr = 0
            cg = (color * 1024)
            cb = 255
        Case Is <= 512
            cr = 0
            cg = 255
            cb = 255 - Int((color * 1024) - 256)
        Case Is <= 768
            cr = (color * 1024) - 512
            cg = 255
            cb = 0
        Case Is <= 1024
            cr = 255
            cg = 256 - Int((color * 1024) - 768)
            cb = 0
    End Select
End Sub

