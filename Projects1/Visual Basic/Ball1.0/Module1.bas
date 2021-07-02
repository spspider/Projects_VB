Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------
'------------------------------------------
'всегда впереди
'------------------------------------------
'------------------------------------------
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'------------------------------------------
'------------------------------------------
'всегда впереди
'------------------------------------------





'------------------------------------------
'------------------------------------------
'Иконка в трее
'------------------------------------------
'Спасибо Антонову Д.А.   mailto: vsdad@pisem.net
'                         http://vsdad.pisem.net
'------------------------------------------
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public TreyIcon As NOTIFYICONDATA
'------------------------------------------
'------------------------------------------
'Иконка в трее
'------------------------------------------
'------------------------------------------



