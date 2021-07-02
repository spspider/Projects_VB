VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Not defined right in win32.tlb
Private Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long
Private Const VTA_BASELINE = TA_BASELINE
Private Const VTA_BOTTOM = TA_RIGHT
Private Const VTA_CENTER = TA_CENTER
Private Const VTA_LEFT = TA_BOTTOM
Private Const VTA_RIGHT = TA_TOP
Private Const VTA_TOP = TA_LEFT

' Handle to current font
Public m_Font As Long

' Internal data type for RGB and Hex conversions
Private Type RGBLongs
    r As Long
    G As Long
    b As Long
End Type

Private m_TextSize As Size

' Used to fill a listbox with all the fonts on the
' computer, which can then be used with CreateNewFont
' for maybe a font selector
Public Sub FillListBoxWithFonts(lstbx As ListBox)
    For i = 0 To Screen.FontCount
        lstbx.AddItem Screen.Fonts(i), i
    Next i
End Sub

' Sets intercharacter spacing
Public Sub SetTextSpacing(obj As Object, spacing As Long)
    SetTextCharacterExtra obj.hDC, spacing
End Sub

' Get width of string of text in pixels
Public Function GetTextWidth(obj As Object, txt As String) As Long
    GetTextExtentPoint32 obj.hDC, txt, Len(txt), m_TextSize
    GetTextWidth = m_TextSize.cx
End Function

' Get height of string of text in pixels
Public Function GetTextHeight(obj As Object, txt As String) As Long
    GetTextExtentPoint32 obj.hDC, txt, Len(txt), m_TextSize
    GetTextHeight = m_TextSize.cy
End Function

' Creates a new font according to your specifications
Public Sub CreateNewFont(FontName As String, Optional Width As Integer = 0, Optional Height As Integer = 0, Optional FontWeight As Integer = FW_NORMAL, Optional Italic As Boolean = False, Optional Underline As Boolean = False, Optional Strikeout As Boolean = False)
    DeleteObject m_Font
    m_Font = CreateFont(Height, Width, 0, 0, FontWeight, Italic, Underline, Strikeout, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, FontName)
    If m_Font = 0 Then MsgBox "Unable to create font: " & FontName
End Sub

' Sets the objects font to the currently selected one
Public Sub SetFont(hDC As Long)
    SelectObject hDC, m_Font
End Sub

Public Sub GradTextOut(obj As Object, X As Integer, Y As Integer, RGBStart As Long, RGBEnd As Long, text As String)
    Dim rgbls As RGBLongs
    Dim R1 As Long, G1 As Long, B1 As Long, R2 As Long, G2 As Long, B2 As Long
    Dim RI As Long, GI As Long, BI As Long
    Dim RDiff As Integer, GDiff As Integer, BDiff As Integer
    Dim CurR As Integer, CurG As Integer, CurB As Integer
    Dim currenttl As Long
    
    rgbls = RGBConv(RGBStart)
    R1 = rgbls.r
    G1 = rgbls.G
    B1 = rgbls.b

    rgbls = RGBConv(RGBEnd)
    R2 = rgbls.r
    G2 = rgbls.G
    B2 = rgbls.b
    
    RDiff = R2 - R1
    RI = RDiff / (Len(text) - 1)

    GDiff = G2 - G1
    GI = GDiff / (Len(text) - 1)

    BDiff = B2 - B1
    BI = BDiff / (Len(text) - 1)

    For i = 0 To (Len(text) - 1)
        CurR = R1 + (RI * i)
        CurG = G1 + (GI * i)
        CurB = B1 + (BI * i)

        If CurR < 0 Then CurR = 0
        If CurG < 0 Then CurG = 0
        If CurB < 0 Then CurB = 0
 
        obj.ForeColor = RGB(CurR, CurG, CurB)
        
        currenttl = currenttl + GetTextWidth(obj, Mid(text, i + 1, 1))

        TextOut obj.hDC, X + (currenttl) - GetTextWidth(obj, Mid(text, 1, 1)), Y, Mid(text, i + 1, 1), 1
    Next i
End Sub

' Center aligns text on the screen when you use TextOut
Public Sub CenterAlignText(obj As Object)
    SetTextAlign obj.hDC, TA_CENTER Or VTA_CENTER
End Sub

' Make background behind text transparent
Public Sub SetTextBackTrans(obj As Object)
    SetBkMode obj.hDC, TRANSPARENT
End Sub

' Set background colour
Public Sub SetTextBackColour(obj As Object, RGBcolor As Long)
    SetBkMode obj.hDC, OPAQUE
    SetBkColor obj.hDC, RGBcolor
End Sub

' Converts RGB to its seperate components
Private Function RGBConv(RGBC As Long) As RGBLongs
    Dim hRGB As String
    Dim rgbl As RGBLongs
    
    hRGB = Hex(RGBC)
    
    Select Case Len(hRGB)
        Case 5:
            hRGB = "0" & hRGB
        Case 4:
            hRGB = "00" & hRGB
        Case 3:
            hRGB = "000" & hRGB
        Case 2:
            hRGB = "0000" & hRGB
        Case 1:
            hRGB = "00000" & hRGB
    End Select
    
    RGBConv = HexToRGB(hRGB)
End Function

Private Function HexToRGB(H As String) As RGBLongs
    Dim Tmp$
    Dim r As Integer, b As Integer
    Dim G As Long
    Dim tmprgbl As RGBLongs
    
    Const Hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    
    Tmp = H
    
    If UCase(left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
    
    Tmp = right$("0000000" & Tmp, 8)
    
    If IsNumeric(Hx & Tmp) Then
        r = CInt(Hx & right$(Tmp, Two))
        G = CLng(Hx & Mid$(Tmp, 5, Two))
        b = CInt(Hx & Mid$(Tmp, 3, Two))
    End If
    
    tmprgbl.r = r
    tmprgbl.G = G
    tmprgbl.b = b
    
    HexToRGB = tmprgbl
End Function

Private Sub Class_Initialize()
    m_Font = CreateFont(0, 0, 0, 0, FW_NORMAL, False, False, False, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, "COMIC SANS MS")
End Sub

Private Sub Class_Terminate()
    DeleteObject m_Font
End Sub
