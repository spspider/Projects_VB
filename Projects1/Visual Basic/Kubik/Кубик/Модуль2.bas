Attribute VB_Name = "Модуль2"
'Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type Pol
     x As Single
     y As Single
     Z As Single
'     xp(1 To 3) As Single
'     yp(1 To 3) As Single
'     zp(1 To 3) As Single
'     color As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type
'Public Type LOGBRUSH
'    lbStyle As Long
'    lbColor As Long
'    lbHatch As Long
'End Type

Public Const Pi = 3.14159265358979
'Public Const ALTERNATE = 1
'Public MoveG As Boolean
Public hwdRgn&, hwdBrush& ', brush As LOGBRUSH
Public ax As Single, ay As Single, az As Single
Public z0 As Single
Public xmax As Single, ymax As Single
Public zs As Single
Public dx As Integer, dy As Integer
'Public ox(1 To 3, 1 To 2) As Single
'Public oy(1 To 3, 1 To 2) As Single
'Public oz(1 To 3, 1 To 2) As Single
'Public orx(1 To 2) As Single
'Public oyr(1 To 2) As Single
'Public pts As POINTAPI
Public pts(1 To 4) As POINTAPI

Public Poligon() As Pol
Public PoligonV() As Pol









'Public Type BITMAP '14 bytes
'        bmType As Long
'        bmWidth As Long
'        bmHeight As Long
'        bmWidthBytes As Long
'        bmPlanes As Integer
'        bmBitsPixel As Integer
'        bmBits As Long
'End Type


'Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
'Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'
'Public Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
'Public Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
'Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

