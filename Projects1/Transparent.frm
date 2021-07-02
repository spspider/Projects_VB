VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbWindows98orHigher As Boolean
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
Private Function fGetOSVersion()
Dim os As OSVERSIONINFO
fGetOSVersion = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)
    If .dwMajorVersion > 4 Then fGetOSVersion = True
    If .dwMajorVersion = 4 And _
       .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And _
       .dwMinorVersion > 0 Then
        fGetOSVersion = True
    End If
End With
End Function
Private Function fSetTranslucency(ByVal hWnd As Long, ByVal alpha As Byte) As Boolean
Dim lStyle As Long
If fIsWin2000 Then
    hWnd = fGetTopLevel(hWnd)
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(hWnd, GWL_EXSTYLE, lStyle) Then
        fSetTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(alpha), LWA_ALPHA))
    End If
End If
End Function
Private Function fClearTranslucency(ByVal hWnd As Long) As Boolean
Dim lStyle As Long
If fIsWin2000 Then
    hWnd = fGetTopLevel(hWnd)
    Call SetLayeredWindowAttributes(hWnd, 0, 255&, LWA_ALPHA)
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
    fClearTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, lStyle))
End If
End Function
Private Function fIsWin2000() As Boolean
Dim os As OSVERSIONINFO
fIsWin2000 = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)
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
Private Sub Command1_Click()
Call fSetTranslucency(Me.hWnd, 200)
End Sub
Private Sub Command2_Click()
Call fClearTranslucency(Me.hWnd)
End Sub
