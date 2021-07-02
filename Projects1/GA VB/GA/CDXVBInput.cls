VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VERY VERY VERY CLOSE TO COMPLETION!
' Alls thats missing is Joystick support... But I dont have a joystick anyway so how
' am I going to add it? :)
'
' Just DONT FORGET TO ADD GUIDDEFS.BAS ALONG WITH THIS CLASS!!!
' GUIDDEFS.BAS contains code you need to a) Create this class and b) to use any input
' data!

Public m_lpDI As IDirectInputA
Public m_lpDIDKeyboard As IDirectInputDeviceA
Public m_lpDIDMouse As IDirectInputDeviceA

Private Sub Class_Terminate()
      m_lpDIDKeyboard.UnAcquire
      Set m_lpDIDKeyboard = Nothing
      
      m_lpDIDMouse.UnAcquire
      Set m_lpDIDMouse = Nothing
      
      Set m_lpDI = Nothing
End Sub

Public Sub Create(hInst As Long, hWnd As Long)
      Dim res As Long

      DirectInputCreateA ByVal hInst, &H500, m_lpDI, Nothing
      
      m_lpDI.CreateDevice GUID_SysKeyboard, m_lpDIDKeyboard, Nothing
      m_lpDIDKeyboard.SetDataFormat c_dfDIKeyboard
      m_lpDIDKeyboard.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND

      m_lpDI.CreateDevice GUID_SysMouse, m_lpDIDMouse, Nothing
      m_lpDIDMouse.SetDataFormat c_dfDIMouse
      m_lpDIDMouse.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
End Sub

Public Sub ReAcquire()
      m_lpDIDKeyboard.Acquire
      m_lpDIDMouse.Acquire
End Sub

Public Sub UnAcquire()
      m_lpDIDKeyboard.UnAcquire
      m_lpDIDMouse.UnAcquire
End Sub

Public Sub UpdateMouse()
      m_lpDIDMouse.GetDeviceState Len(MouseState), MouseState
      If Err <> 0 Then m_lpDIDMouse.Acquire
End Sub

Public Sub UpdateKeyboard()
      m_lpDIDKeyboard.GetDeviceState 256, Keys(0)
      If Err <> 0 Then m_lpDIDKeyboard.Acquire
End Sub
