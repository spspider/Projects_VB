Attribute VB_Name = "Module1"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean '��������� ��� ����������, �������� � ����������� ����� ������:
Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
'��������� ������������� �� ������� ������������ ������ ������ ������, ������������� � Traybar:
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4
'��������� ������������� �� ������� ��������� ����� ������������ ������ ������ ������, ' ������������� � Traybar:
'��� ����� ������� �����:
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
'��� ������ ������� �����:
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
'��� ������� ������� �����:
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
'��������� ���������� ������������ �������������:
Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type

