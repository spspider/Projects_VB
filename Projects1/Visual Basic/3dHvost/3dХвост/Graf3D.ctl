VERSION 5.00
Begin VB.UserControl Graf3D 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LockControls    =   -1  'True
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   2040
   End
End
Attribute VB_Name = "Graf3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Написано Vir <ou953@ipn.ru>

Const Pi = 3.14159265358979
Dim z0 As Single
Dim xMax As Single, yMax As Single
Dim zs As Single
'Dim dx As Integer, dy As Integer
Dim V() As Pos
Dim U(20) As Pos
Dim R() As Pos
Dim RZ() As Single
Dim RP() As Integer
Dim S As Pos
Dim T As Pos
Dim tMax As Pos
'Dim kk as Boolean

Sub calc()
    m = UBound(U)
'    If kk = True Then
        For i = m To 1 Step -1
            U(i).x = U(i - 1).x
            U(i).y = U(i - 1).y
            U(i).z = U(i - 1).z
        Next
'        kk = False
'    Else
'        kk = True
'    End If
    For o = 1 To m
        For i = 0 To 23
        
            n = o + i * m
            zz1 = V(n).z * Cos(U(o).y * Pi / 180) + V(n).x * Sin(U(o).y * Pi / 180)
            xx1 = V(n).x * Cos(U(o).y * Pi / 180) - V(n).z * Sin(U(o).y * Pi / 180)
            zz2 = zz1 * Cos(U(o).x * Pi / 180) + V(n).y * Sin(U(o).x * Pi / 180)
            yy1 = V(n).y * Cos(U(o).x * Pi / 180) - zz1 * Sin(U(o).x * Pi / 180)
            
            xx2 = xx1 * Cos(U(o).z * Pi / 180) + yy1 * Sin(U(o).z * Pi / 180)
            yy2 = yy1 * Cos(U(o).z * Pi / 180) - xx1 * Sin(U(o).z * Pi / 180)
            V(n).x = xx2
            V(n).y = yy2
            V(n).z = zz2
            
            R(n).x = xMax / 2 + zs * xx2 / (zz2 * 5 - z0)
            R(n).y = yMax / 2 + zs * yy2 / (zz2 * 5 - z0)
            R(n).z = zz2
            RZ(n) = zz2
            RP(n) = n
        Next
    Next
    
    b = UBound(RZ)
    c = b
nn2:
    c = Int(c / 2)
    If c <= 0 Then Exit Sub
    d = b - c
    e = 0
nn3:
    f = e
nn1:
    g = f + c
    If RZ(f) < RZ(g) Then
    temp = RZ(f)
    RZ(f) = RZ(g)
    RZ(g) = temp
    
    temp = RP(f)
    RP(f) = RP(g)
    RP(g) = temp
    temp = RP(f)
'    R(f).y = R(g).y
'    R(g).y = temp

    f = f - c
    If f > 0 Then GoTo nn1
    End If
    e = e + 1
    If e > d Then GoTo nn2 Else GoTo nn3
End Sub

Sub PixelPoint()
    m = UBound(U)
    For i = 1 To UBound(V)  ' Step m
        
        'For o = 0 To m - 2
        If i / m <> i \ m Then
            c = 255 - (R(i + o).z + m * 2) * 255 / (m * 4)
            Line (R(i + o).x, R(i + o).y)-(R(i + o + 1).x, R(i + o + 1).y), RGB(c, c, 0)
        End If


        'Next


    Next
    For i = 1 To UBound(V) Step m * 4
'        If i / (m) <> i \ (m) And i / (m - 1) <> i \ (m - 1) And i / (m - 2) <> i \ (m - 2) Then
'        If i / (m * 4) = i \ (m * 4) Or i = 1 Then
        For o = 0 To m - 1
            c = 255 - (R(i + o).z + m * 2) * 255 / (m * 4)
            Line (R(i + o + m * 0).x, R(i + o + m * 0).y)-(R(i + o + m * 1).x, R(i + o + m * 1).y), RGB(c, c, 0)
            Line (R(i + o + m * 1).x, R(i + o + m * 1).y)-(R(i + o + m * 2).x, R(i + o + m * 2).y), RGB(c, c, 0)
            Line (R(i + o + m * 2).x, R(i + o + m * 2).y)-(R(i + o + m * 3).x, R(i + o + m * 3).y), RGB(c, c, 0)
            Line (R(i + o + m * 3).x, R(i + o + m * 3).y)-(R(i + o + m * 0).x, R(i + o + m * 0).y), RGB(c, c, 0)
        Next
'        End If
'        End If
    Next
    
End Sub

Private Sub Timer1_Timer()
    If T.x < tMax.x Then
        T.x = T.x + 1
    Else
        S.x = Rnd * 10 - 5
        T.x = 0
        tMax.x = Rnd * 100
    End If
    If T.y < tMax.y Then
        T.y = T.y + 1
    Else
        S.y = Rnd * 10 - 5
        T.y = 0
        tMax.y = Rnd * 100
    End If
    If T.z < tMax.z Then
        T.z = T.z + 1
    Else
        S.z = Rnd * 10 - 5
        T.z = 0
        tMax.z = Rnd * 100
    End If
'    S.x = Sin(Timer) * 5
'    S.y = Sin(Timer) * 8
'    S.z = Sin(Timer) * 6
    
    U(0).x = S.x
    U(0).y = S.y
    U(0).z = S.z
    Resize
End Sub

Private Sub UserControl_Click()
    End
End Sub

Private Sub UserControl_Hide()
    Timer1.Enabled = False
End Sub

Public Sub AddMass()
    m = UBound(U)
    ReDim V(24 * m)
    ReDim R(24 * m)
    ReDim RZ(24 * m)
    ReDim RP(24 * m)
    
    h = 2 'длинна H>0
    d = 2 'Диаметр D>0
    For i = 1 To m
        V(i + m * 0).x = i * h
        V(i + m * 0).y = (i / m - 1) * d
        V(i + m * 0).z = (i / m - 1) * d
        
        V(i + m * 1).x = i * h
        V(i + m * 1).y = (-i / m + 1) * d
        V(i + m * 1).z = (i / m - 1) * d
        
        V(i + m * 3).x = i * h
        V(i + m * 3).y = (i / m - 1) * d
        V(i + m * 3).z = (1 + -i / m) * d
        
        V(i + m * 2).x = i * h
        V(i + m * 2).y = (1 + -i / m) * d
        V(i + m * 2).z = (1 + -i / m) * d
        '''''''''
        V(i + m * 4).x = (i / m - 1) * d
        V(i + m * 4).y = i * h
        V(i + m * 4).z = (i / m - 1) * d
        
        V(i + m * 5).x = (-i / m + 1) * d
        V(i + m * 5).y = i * h
        V(i + m * 5).z = (i / m - 1) * d
        
        V(i + m * 7).x = (i / m - 1) * d
        V(i + m * 7).y = i * h
        V(i + m * 7).z = (1 + -i / m) * d
        
        V(i + m * 6).x = (1 + -i / m) * d
        V(i + m * 6).y = i * h
        V(i + m * 6).z = (1 + -i / m) * d
        ''''''''
        V(i + m * 8).x = (i / m - 1) * d
        V(i + m * 8).y = (i / m - 1) * d
        V(i + m * 8).z = i * h
        
        V(i + m * 9).x = (i / m - 1) * d
        V(i + m * 9).y = (-i / m + 1) * d
        V(i + m * 9).z = i * h
        
        V(i + m * 11).x = (1 + -i / m) * d
        V(i + m * 11).y = (i / m - 1) * d
        V(i + m * 11).z = i * h
        
        V(i + m * 10).x = (1 + -i / m) * d
        V(i + m * 10).y = (1 + -i / m) * d
        V(i + m * 10).z = i * h
        ''''''''''''
        V(i + m * 12).x = -i * h
        V(i + m * 12).y = (i / m - 1) * d
        V(i + m * 12).z = (i / m - 1) * d
        
        V(i + m * 13).x = -i * h
        V(i + m * 13).y = (-i / m + 1) * d
        V(i + m * 13).z = (i / m - 1) * d
        
        V(i + m * 15).x = -i * h
        V(i + m * 15).y = (i / m - 1) * d
        V(i + m * 15).z = (1 + -i / m) * d
        
        V(i + m * 14).x = -i * h
        V(i + m * 14).y = (1 + -i / m) * d
        V(i + m * 14).z = (1 + -i / m) * d
        '''''''''
        V(i + m * 16).x = (i / m - 1) * d
        V(i + m * 16).y = -i * h
        V(i + m * 16).z = (i / m - 1) * d
        
        V(i + m * 17).x = (-i / m + 1) * d
        V(i + m * 17).y = -i * h
        V(i + m * 17).z = (i / m - 1) * d
        
        V(i + m * 19).x = (i / m - 1) * d
        V(i + m * 19).y = -i * h
        V(i + m * 19).z = (1 + -i / m) * d
        
        V(i + m * 18).x = (1 + -i / m) * d
        V(i + m * 18).y = -i * h
        V(i + m * 18).z = (1 + -i / m) * d
        ''''''''''''
        V(i + m * 20).x = (i / m - 1) * d
        V(i + m * 20).y = (i / m - 1) * d
        V(i + m * 20).z = -i * h
        
        V(i + m * 21).x = (i / m - 1) * d
        V(i + m * 21).y = (-i / m + 1) * d
        V(i + m * 21).z = -i * h
        
        V(i + m * 23).x = (1 + -i / m) * d
        V(i + m * 23).y = (i / m - 1) * d
        V(i + m * 23).z = -i * h
        
        V(i + m * 22).x = (1 + -i / m) * d
        V(i + m * 22).y = (1 + -i / m) * d
        V(i + m * 22).z = -i * h
    Next
    
    Resize
End Sub

Private Sub UserControl_Initialize()
    Randomize Time
    z0 = -500
End Sub

'Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 1 Then
'        U(0).x = U(0).x + (y - dy) / 5
'        U(0).y = U(0).y + (x - dx) / 5
'    End If
'    dx = x
'    dy = y
'End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "\n\n\n\n\n\n\n                                                                                                                Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Timer1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Timer1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Timer1.Enabled = PropBag.ReadProperty("Enabled", False)
End Sub

Private Sub UserControl_Show()
'    Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Timer1.Enabled, False)
End Sub
    
Public Sub Resize()
    If ScaleHeight < ScaleWidth Then zs = ScaleHeight * 6 Else zs = ScaleWidth * 6
    xMax = ScaleWidth
    yMax = ScaleHeight
    Call calc
    Cls
    Call PixelPoint
    U(0).x = 0
    U(0).y = 0
    U(0).z = 0
End Sub
