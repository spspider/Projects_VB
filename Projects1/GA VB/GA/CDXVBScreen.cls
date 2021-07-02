VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ALMOST working...

Private Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private video_buffer() As Byte
Private sa As SAFEARRAY2D

Public m_lpdd As IDirectDraw2
Private m_ddsd As DDSURFACEDESC
Public m_lpDDSFront As IDirectDrawSurface2
Public m_lpDDSBack As IDirectDrawSurface2

Public m_PixelWidth As Integer
Public m_PixelHeight As Integer
Public m_BPP As Integer
Public m_HWND As Long
Public m_HDC As Long
Public m_Font As Long

Private ScreenRect As RECT

Public Function CreateFullScreen(hWnd As Long, Width As Integer, Height As Integer, BPP As Integer, Optional bVGA As Boolean = False) As Boolean
      Dim result As Long
      Dim dwFlags As Long
      Dim ddsCaps1 As DDSCAPS
      Dim ddsd As DDSURFACEDESC
      
      m_PixelWidth = Width
      m_PixelHeight = Height
      m_HWND = hWnd
      m_BPP = BPP

      dwFlags = DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT Or DDSCL_ALLOWMODEX

      DirectDrawCreate ByVal 0&, m_lpdd, Nothing
      
      m_lpdd.SetCooperativeLevel hWnd, dwFlags
      
      If bVGA = True Then
            m_lpdd.SetDisplayMode Width, Height, BPP, 0, DDSDM_STANDARDVGAMODE
      Else
            m_lpdd.SetDisplayMode Width, Height, BPP, 0, 0
      End If

      ddsd.dwSize = Len(ddsd)
      ddsd.dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
      ddsd.DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
      ddsd.dwBackBufferCount = 1

      m_lpdd.CreateSurface ddsd, m_lpDDSFront, Nothing
      
      ddsCaps1.dwCaps = DDSCAPS_BACKBUFFER
      
      m_lpDDSFront.GetAttachedSurface ddsCaps1, m_lpDDSBack
      
      With ScreenRect
            .top = 0
            .left = 0
            .bottom = Height
            .right = Width
      End With
End Function

Public Function Flip() As Long
      m_lpDDSFront.Flip Nothing, DDFLIP_WAIT
End Function

Public Sub CloseCDXVBScreen()
      m_lpdd.FlipToGDISurface
      m_lpdd.SetCooperativeLevel 0, DDSCL_NORMAL
      m_lpdd.RestoreDisplayMode

      Set m_lpDDSBack = Nothing
      Set m_lpDDSFront = Nothing
      Set m_lpdd = Nothing
End Sub

Public Sub ClearBack()
      Dim ClearFX As DDBLTFX

      With ClearFX
            .dwSize = Len(ClearFX)
            .dwFillColor = 0
      End With

      m_lpDDSBack.Blt ScreenRect, Nothing, ScreenRect, DDBLT_COLORFILL Or DDBLT_WAIT, ClearFX
End Sub

Public Sub HideMouse()
      ShowCursor False
End Sub

Public Sub ShowMouse()
      ShowCursor True
End Sub

Public Sub SurfGetBackDC()
      m_lpDDSBack.GetDC m_HDC
End Sub

Public Sub SurfReleaseBackDC()
      m_lpDDSBack.ReleaseDC m_HDC
End Sub

Private Sub Class_Terminate()
      Call CloseCDXVBScreen
End Sub

Public Sub LockMe()
      CopyMemory m_ddsd, ByVal 0&, Len(m_ddsd)
      m_ddsd.dwSize = Len(m_ddsd)

      m_lpDDSBack.Lock ByVal 0&, m_ddsd, DDLOCK_WAIT Or DDLOCK_SURFACEMEMORYPTR, 0

      With sa
            .cbElements = 1
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = m_ddsd.dwHeight
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = m_ddsd.dwWidth
            .pvData = m_ddsd.lpSurface
      End With
      CopyMemory ByVal VarPtrArray(video_buffer), VarPtr(sa), 4
End Sub

Public Sub Pixel(X As Integer, Y As Integer, Color As Integer)
      video_buffer(X, Y) = Color
End Sub

Public Sub UnLockMe()
      m_lpDDSBack.Unlock m_ddsd.lpSurface

      CopyMemory ByVal VarPtrArray(video_buffer), ByVal 0&, 4
End Sub

