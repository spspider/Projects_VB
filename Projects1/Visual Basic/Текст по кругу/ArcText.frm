VERSION 5.00
Begin VB.Form frmArcText 
   AutoRedraw      =   -1  'True
   Caption         =   "Text along arc or ellipse"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   405
      Left            =   2610
      TabIndex        =   3
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   """Buiging"""
      Height          =   405
      Left            =   1530
      TabIndex        =   2
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sin && Cos"
      Height          =   405
      Left            =   450
      TabIndex        =   1
      Top             =   90
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4275
      Left            =   150
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   371
      TabIndex        =   0
      Top             =   600
      Width           =   5625
   End
End
Attribute VB_Name = "frmArcText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ArcText.frm
'
' By Herman Liu
'
' Print text at angles to form an arc or ellipse shape.  Two methods are demonstrated to
' compute the position points to place the letterings, and in both cases, the API function
' CreateFontIndirect is called to create the text effect (a crude one for demo only).
'
' Method one is based on the conventional approach using sine, cosine and tangent...of an
' angle, e.g.
'                      / |
'           Hypotenuse/  |Opposite side
'                    /   |
'          Angle -> /    |
'              Adjacent side
'
'     Length of its opposite side = Sine of angle  X  Hypotenuse,
'     Length of its adjacent side = Cosine of angle  X  Hypotenuse.
'
' Method two is based on materials of a snippet posted by "LPChip InterActive".  The said
' snippet consists of two functions, GetXfromAngle and GetYfromAngle.  The following is
' an extract of its description):
'
'    ... for games or other applications where you need a corner with an angle. You specify
'     the starting XY(center), the Length (up) and the Angle, then you'll retrieve the XY
'     of where it is. (This routine is 10 times faster than Tan, Cos and Sin.)
'                    |   +Retrieved XY
'                    |  /
'                    | /Length
'                    |/<- Angle
'                   XY+
'     A variable called "buiging" is used. Its default value is set at 1.80; changing this
'     could create cool things. You can change the size of the circle without changing the
'     width on the horizontal and vertical points of the circle (to draw ellipses).
'
Option Explicit

Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
       "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, _
      ByVal iMode As Long) As Long
      
Private Const GM_ADVANCED = 2                  ' Required for NT
Private Const FW_BOLD = 700

Private Type LOGFONT
    lfHeight As Long                           ' In logical units
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 33                  ' L_FACESIZE
End Type

Private Const PI = 3.14159265
Dim Xcurr As Single
Dim Ycurr As Single
Dim mresult



Private Sub Command1_Click()
    Dim Xradius As Single
    Dim Yradius As Single
    Dim mWhole As Single
    Dim mSection As Single
    Dim mAngle As Integer
    Dim i As Integer
    Dim mString As String
    
    mString = "Using Sin & Cos?"
       ' A center
    Xcurr = Picture1.ScaleWidth / 2               ' At middle horizontally
    Ycurr = Picture1.ScaleHeight / 1.8            ' A suitable verticl point
    mWhole = PI               ' Start with whole ratio of circumference to diameter
    mSection = PI / Len(mString)
       ' Set radius values appropriately (according to shape and size )
    Xradius = Picture1.ScaleWidth * 0.5
    Yradius = Picture1.ScaleHeight * 0.5
       ' We start from 270, but will less 180 on calling doPrintChr, so
       ' effectively 90.
    mAngle = 270
    Picture1.Cls
    For i = 1 To Len(mString)
        Picture1.CurrentX = Xcurr + (Xradius * Cos(mWhole))
        Picture1.CurrentY = Ycurr - (Yradius * Sin(mWhole))
        doPrintChr mAngle - 180, Mid$(mString, i, 1)
           ' Note we have to find a suitable interval (according to font and size)
        mAngle = mAngle - Fix(180 / Len(mString))
        mWhole = mWhole - mSection
    Next i
End Sub



Private Sub Command2_Click()
    Dim mLength As Integer
    Dim mAngle As Integer
    Dim Buiging As Single
    Dim i As Integer
    Dim mString As String
    
    mString = "Using buiging?"
       ' A center
    Xcurr = Picture1.ScaleWidth / 2
    Ycurr = Picture1.ScaleHeight + 50           ' To lower the display a bit
    Buiging = 1.8
    mLength = 140
       ' We start from 270, but will less 180 on calling doPrintChr, so
       ' effectively 90.
    mAngle = 270
       ' Allowance for half of the interval size, for symmetrical display.
       ' (Due to proportional font being used here, result would unlikely
       ' be symmetrical though.)
    mAngle = mAngle - (Fix(180 / Len(mString)) / 2)
    For i = 1 To Len(mString)
        Picture1.CurrentX = Xcurr + GetXfromAngle(mAngle, mLength, Buiging)
        Picture1.CurrentY = Ycurr - GetYfromAngle(mAngle, mLength, Buiging)
           ' Note we have to find a suitable interval (according to font and size)
        doPrintChr mAngle - 180, Mid$(mString, i, 1)
        mAngle = mAngle - Fix(180 / Len(mString))
    Next i
End Sub


' Read comment at top: The following two functions, GetXfromAngle and GetYfromAngle, are
' from the snippet authored by LPChip Interactive.
Private Function GetXfromAngle(Angle, Length, Buiging)
    Select Case Angle
        Case 0 To 89, 360
            GetXfromAngle = Length - (90 - Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging
        Case 90 To 179
            GetXfromAngle = Length - (Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging
        Case 180 To 269
            GetXfromAngle = (90 - Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging - Length
        Case 270 To 359
            GetXfromAngle = (Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging - Length
    End Select
End Function


Private Function GetYfromAngle(Angle, Length, Buiging)
    Select Case Angle
        Case 0 To 89, 360
            GetYfromAngle = (Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging
        Case 90 To 179
            GetYfromAngle = Length - (90 - Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging + Length
        Case 180 To 269
            GetYfromAngle = Length - (Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging + Length
        Case 270 To 360
            GetYfromAngle = (90 - Angle Mod 90) ^ Buiging * Length / 90 ^ Buiging
   End Select
End Function



Private Sub doPrintChr(ByVal inAngle As Integer, ByVal inText As String)
    On Error Resume Next
    Dim L As LOGFONT
    Dim mFont As Long
    Dim mOrigFont As Long
    Dim mFontSize As Integer
    mresult = SetGraphicsMode(Picture1.hdc, GM_ADVANCED)
    mFontSize = 24
    L.lfFaceName = "Times New Roman" & Space(1) & "Bold" & Chr(0)
    L.lfEscapement = inAngle * 10
    L.lfHeight = mFontSize * -20 / Screen.TwipsPerPixelY
    mFont = CreateFontIndirect(L)
    mOrigFont = SelectObject(Picture1.hdc, mFont)
    Picture1.Print inText
    mresult = DeleteObject(mFont)
End Sub



Private Sub Command3_Click()
    Picture1.Cls
End Sub

