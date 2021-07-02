Attribute VB_Name = "Модуль4"
Public Sub PixelPoint()
    a = Array(0, 1, 2, 3)
    Polig a
    a = Array(4, 5, 6, 7)
    Polig a
    a = Array(0, 1, 5, 6)
    Polig a
    a = Array(3, 2, 6, 7)
    Polig a
    a = Array(1, 2, 6, 5)
    Polig a
    a = Array(0, 3, 7, 4)
    Polig a
    
    a = Array(0, 1)
    Lines a
    a = Array(1, 2)
    Lines a
    a = Array(2, 3)
    Lines a
    a = Array(3, 0)
    Lines a
    a = Array(0, 4)
    Lines a
    a = Array(1, 5)
    Lines a
    a = Array(2, 6)
    Lines a
    a = Array(3, 7)
    Lines a
    a = Array(4, 5)
    Lines a
    a = Array(5, 6)
    Lines a
    a = Array(6, 7)
    Lines a
    a = Array(7, 4)
    Lines a
End Sub

Private Sub Polig(a)
    pts(1).x = xmax / 2 + zs * PoligonV(a(0)).x / (PoligonV(a(0)).Z - z0)
    pts(1).y = ymax / 2 + zs * PoligonV(a(0)).y / (PoligonV(a(0)).Z - z0)
    pts(2).x = xmax / 2 + zs * PoligonV(a(1)).x / (PoligonV(a(1)).Z - z0)
    pts(2).y = ymax / 2 + zs * PoligonV(a(1)).y / (PoligonV(a(1)).Z - z0)
    pts(3).x = xmax / 2 + zs * PoligonV(a(2)).x / (PoligonV(a(2)).Z - z0)
    pts(3).y = ymax / 2 + zs * PoligonV(a(2)).y / (PoligonV(a(2)).Z - z0)
    pts(4).x = xmax / 2 + zs * PoligonV(a(3)).x / (PoligonV(a(3)).Z - z0)
    pts(4).y = ymax / 2 + zs * PoligonV(a(3)).y / (PoligonV(a(3)).Z - z0)
    
    hwdRgn& = CreatePolygonRgn(pts(1), 4, 1)
    hwdBrush& = CreateSolidBrush(RGB(150, 150, 0))
    FillRgn Form1.Picture1.hdc, hwdRgn&, hwdBrush&
    DeleteObject hwdRgn&
    DeleteObject hwdBrush&
End Sub

Private Sub Lines(a)
    Form1.Picture1.Line (xmax / 2 + zs * PoligonV(a(0)).x / (PoligonV(a(0)).Z - z0) _
        , ymax / 2 + zs * PoligonV(a(0)).y / (PoligonV(a(0)).Z - z0))- _
        (xmax / 2 + zs * PoligonV(a(1)).x / (PoligonV(a(1)).Z - z0) _
        , ymax / 2 + zs * PoligonV(a(1)).y / (PoligonV(a(1)).Z - z0)), _
        RGB(255, 255, 0)
End Sub

