Attribute VB_Name = "Модуль3"
Public Sub calc()
On Error GoTo ErrExit
    For i = 0 To UBound(Poligon)
        With Poligon(i)
            zz1 = .Z * Cos(ay * Pi / 180) + .x * Sin(ay * Pi / 180)
            xx1 = .x * Cos(ay * Pi / 180) - .Z * Sin(ay * Pi / 180)
            zz2 = zz1 * Cos(ax * Pi / 180) + .y * Sin(ax * Pi / 180)
            yy1 = .y * Cos(ax * Pi / 180) - zz1 * Sin(ax * Pi / 180)
            xx2 = xx1 * Cos(az * Pi / 180) + yy1 * Sin(az * Pi / 180)
            yy2 = yy1 * Cos(az * Pi / 180) - xx1 * Sin(az * Pi / 180)
            PoligonV(i).x = xx2
            PoligonV(i).y = yy2
            PoligonV(i).Z = zz2
        End With
    Next
'    t = UBound(PoligonV)
'    c = t
'nn2:
'    c = Int(c / 2)
'    While c > 0
'        d = t - c
'        e = 0
'        Do
'            f = e
'nn1:
'            g = f + c
'            s1 = PoligonV(f).Z
'            s2 = PoligonV(g).Z
'
'            If s1 < s2 Then
'                For o = 1 To 3
'                    temp = PoligonV(f).Z
'                    PoligonV(f).Z = PoligonV(g).Z
'                    PoligonV(g).Z = temp
'
'                    temp = PoligonV(f).x
'                    PoligonV(f).x = PoligonV(g).x
'                    PoligonV(g).x = temp
'
'                    temp = PoligonV(f).y
'                    PoligonV(f).y = PoligonV(g).y
'                    PoligonV(g).y = temp
'                Next
'                f = f - c
'                If f > 0 Then GoTo nn1
'            End If
'            e = e + 1
'        Loop Until e > d
'        GoTo nn2
'    Wend
ErrExit:
End Sub

Public Sub AddMass()
    a = Array(-1, 1, 1, -1, 1, -1, 1, 1, -1, 1, 1, 1, -1, -1, 1, -1, -1, -1, 1, -1, -1, 1, -1, 1)
    ReDim Poligon(7)
    ReDim PoligonV(7)
    n = 0
    For i = 0 To 7
        Poligon(i).x = a(n)
        n = n + 1
        Poligon(i).y = a(n)
        n = n + 1
        Poligon(i).Z = a(n)
        n = n + 1
    Next
End Sub
    
Public Sub Resize()
    Form1.Picture1.Cls
    Call calc
    Call PixelPoint
End Sub

