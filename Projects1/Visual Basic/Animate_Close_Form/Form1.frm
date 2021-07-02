VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tile Image/Animate Form/Make Tanslucent"
   ClientHeight    =   5250
   ClientLeft      =   3840
   ClientTop       =   2730
   ClientWidth     =   7725
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7725
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Translucency"
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   1590
      Begin VB.Label lblMake 
         BackStyle       =   0  'Скрытый
         Caption         =   "Make Translucent"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label lblClear 
         BackStyle       =   0  'Скрытый
         Caption         =   "Make Opaque"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Animation Effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1710
      Begin VB.OptionButton opt 
         BackColor       =   &H00000000&
         Caption         =   "Collapse Inward"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   750
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00000000&
         Caption         =   "Slide Diagonally"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   495
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00000000&
         Caption         =   "Slide Downward"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   3600
      Picture         =   "Form1.frx":0442
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbWindows98orHigher As Boolean
'
' Constants used by AnimateWindow
'
'Duration of animation in milliseconds
Const AW_DURATION_DEFAULT = 200
'Animate from left to right
Const AW_HOR_POSITIVE = &H1
'Animate from right to left
Const AW_HOR_NEGATIVE = &H2
'Animate from top to bottom
Const AW_VER_POSITIVE = &H4
'Animate from bottom to top
Const AW_VER_NEGATIVE = &H8
'Collapse window inward when used with
'  AW_HIDE or outward otherwise
Const AW_CENTER = &H10
'Hides the window
Const AW_HIDE = &H10000
'Activates the window
Const AW_ACTIVATE = &H20000
'Slide animation. Cannot use with AW_CENTER
Const AW_SLIDE = &H40000
'Fade window. Only works with top level windows
Const AW_BLEND = &H80000

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function AnimateWindow Lib "user32" ( _
    ByVal hWnd As Long, ByVal dwTime As Long, _
    ByVal dwFlags As Long) As Boolean

Private Declare Function GetVersionEx Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'hWnd    - handle to window to layer.
'crKey   - specifies the color key
'bAlpha  - value for the blend function
'dwFlags - action
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&

Private Declare Function GetParent Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Private Declare Function IsWindowVisible Lib "user32" _
    (ByVal hWnd As Long) As Long


Private Sub Form_Load()
'
' See if we are on Win98 or Win2000.
'
mbWindows98orHigher = fGetOSVersion()
Frame1.Enabled = mbWindows98orHigher
opt(0) = True
End Sub

Private Function fGetOSVersion()
Dim os As OSVERSIONINFO
'
' Returns True if Win98 or Win2000
'
fGetOSVersion = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)

    ' Windows 2000
    If .dwMajorVersion > 4 Then fGetOSVersion = True

    If .dwMajorVersion = 4 And _
       .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And _
       .dwMinorVersion > 0 Then
        fGetOSVersion = True
    End If
End With
End Function



Private Sub Form_Paint()
Dim X As Integer
Dim Y As Integer
'
' Note, the visible property of the image should be set false,
'       and the Form's Auto-Redraw property must be False.
'
' The Form_Paint event is used since it occurs whenever the Form
' is repainted (when the form is restored from minimized form, form
' has been resized, maximized etc.)
'
For X = 0 To Me.Width Step Image1.Width
    For Y = 0 To Me.Height Step Image1.Height
        PaintPicture Image1, X, Y
    Next Y
Next X
End Sub


Private Sub Form_Unload(Cancel As Integer)
'
' Animate the closing of the form. This can be done
' in the form load also. Just use AW_ACTIVATE instead
' of AW_HIDE.
'
' Typically you would use AW_DURATION_DEFAULT for the
' time period. Here I used 500 milliseconds (.5 second)
' to exagerate the efects.
'
If Not mbWindows98orHigher Then Exit Sub

If opt(0) Then
    Call AnimateWindow(Me.hWnd, 500, _
                AW_SLIDE Or AW_VER_POSITIVE Or AW_HIDE)
ElseIf opt(1) Then
    Call AnimateWindow(Me.hWnd, 500, _
                AW_VER_NEGATIVE Or AW_HOR_NEGATIVE Or AW_HIDE)
ElseIf opt(2) Then
    Call AnimateWindow(Me.hWnd, 500, _
                AW_CENTER Or AW_HIDE)
End If
Set Form1 = Nothing
End Sub



Private Function fSetTranslucency(ByVal hWnd As Long, ByVal alpha As Byte) As Boolean
Dim lStyle As Long

'
' Layering only works with Win2K or above.
'
If fIsWin2000 Then
    '
    ' Only a top level window can be translucent.
    '
    hWnd = fGetTopLevel(hWnd)
    '
    ' Make the window translucent by setting its
    ' extended style.
    '
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(hWnd, GWL_EXSTYLE, lStyle) Then
        fSetTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(alpha), LWA_ALPHA))
    End If
End If
End Function
Private Function fClearTranslucency(ByVal hWnd As Long) As Boolean
Dim lStyle As Long

'
' Layering only works with Win2K or above.
'
If fIsWin2000 Then
    '
    ' Only a top level window can be translucent.
    '
    hWnd = fGetTopLevel(hWnd)
    '
    ' Clear translucency - make the window opaque.
    '
    Call SetLayeredWindowAttributes(hWnd, 0, 255&, LWA_ALPHA)
    '
    ' Clear the extended style bit.
    '
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
    fClearTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, lStyle))
End If
End Function


Private Function fIsWin2000() As Boolean
Dim os As OSVERSIONINFO
'
' Returns True if Win98 or Win2000
'
fIsWin2000 = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)

    ' Windows 2000
    If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
        fIsWin2000 = (.dwMajorVersion > 4)
    End If
End With

End Function

Private Function fGetTopLevel(ByVal hChild As Long) As Long
Dim hWnd As Long

hWnd = hChild
Do While IsWindowVisible(GetParent(hWnd))
    hWnd = GetParent(hChild)
    hChild = hWnd
Loop
fGetTopLevel = hWnd
End Function

Private Sub lblClear_Click()

Call fClearTranslucency(Me.hWnd)
End Sub

Private Sub lblMake_Click()
'
' Try values between 0 (completely invisible)
' to 255 (fully opaque).
'
Call fSetTranslucency(Me.hWnd, 160)
End Sub


