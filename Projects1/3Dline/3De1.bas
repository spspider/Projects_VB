Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Sub SetFormPosition(frmHandl As Long, TopPosition As Boolean)
If TopPosition Then
SetWindowPos frmHandl, HWND_TOPMOST, 0, 0, 0, 0, _
SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Else
SetWindowPos frmHandl, HWND_NOTOPMOST, 0, 0, 0, 0, _
SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub


