VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WORKING! AT THE MOMENT...

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Const LR_LOADFROMFILE = &H10
Private Const LR_CREATEDIBSECTION = &H2000
Private Const SRCCOPY = &HCC0020

Private Const TRANSPARENT = 1
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const VARIABLE_PITCH = 2

Private Const FW_NORMAL = 400
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_SEMIBOLD = 600
Private Const FW_THIN = 100


Private Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type


Public m_lpDDS As IDirectDrawSurface2
Private m_ddsd As DDSURFACEDESC

Public m_PixelWidth As Integer
Public m_PixelHeight As Integer
Public Filename As String

Public m_HDC As Long
Public m_Font As Long
Public m_ColorKey As Integer

Private m_SrcRect As RECT
Private m_DstRect As RECT


Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

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

Private Function LoadBitmap(DXObject As IDirectDraw2, ByVal BMPFile As String) As IDirectDrawSurface2
    Dim hBitmap As Long                 ' Handle on bitmap
    Dim dBitmap As BITMAP               ' Handle on bitmap descriptor
    Dim TempDXD As DDSURFACEDESC        ' Surface description
    Dim TempDXS As IDirectDrawSurface3   ' Created surface
    Dim dcBitmap As Long                ' Handle on image
    Dim dcDXS As Long                   ' Handle on surface context
    Dim ddck As DDCOLORKEY
    
    ddck.dwColorSpaceLowValue = 0
    ddck.dwColorSpaceHighValue = 0
    
    ' Load bitmap
    hBitmap = LoadImage(ByVal 0&, BMPFile, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    
    ' Get bitmap descriptor
    GetObject hBitmap, Len(dBitmap), dBitmap
    
    ' Fill DX surface description
    With TempDXD
        .dwSize = Len(TempDXD)
        .dwflags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .dwWidth = dBitmap.bmWidth
        .dwHeight = dBitmap.bmHeight
    End With
    
      m_SrcRect.top = 0
      m_SrcRect.left = 0
      m_SrcRect.bottom = dBitmap.bmHeight
      m_SrcRect.right = dBitmap.bmWidth
      
      m_PixelWidth = dBitmap.bmWidth
      m_PixelHeight = dBitmap.bmHeight
    
    ' Create DX surface
    DXObject.CreateSurface TempDXD, TempDXS, Nothing
    
    ' Create API memory DC
    dcBitmap = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap into API memory DC
    SelectObject dcBitmap, hBitmap
    
    ' Restore DX surface
    TempDXS.Restore
    ' Get DX surface API DC
    TempDXS.GetDC dcDXS
    ' Blit BMP from API DC into DX DC using standard API BitBlt
    StretchBlt dcDXS, 0, 0, TempDXD.dwWidth, TempDXD.dwHeight, dcBitmap, 0, 0, dBitmap.bmWidth, dBitmap.bmHeight, SRCCOPY

    ' Cleanup
    TempDXS.ReleaseDC dcDXS
    DeleteDC dcBitmap
    DeleteObject hBitmap
    
    TempDXS.SetColorKey DDCKEY_SRCBLT, ddck
    ' Return created DX surface
    Set LoadBitmap = TempDXS
End Function

Public Function Create(FN As String, DDraw As CDXVBScreen) As Boolean
      Set m_lpDDS = LoadBitmap(DDraw.m_lpdd, FN)

      Create = True
End Function

Private Sub Class_Terminate()
      Set m_lpDDS = Nothing
End Sub

Public Sub Blit(x As Long, y As Long, Back As IDirectDrawSurface2)
      Back.BltFast x, y, m_lpDDS, m_SrcRect, DDBLTFAST_SRCCOLORKEY
End Sub

Public Sub SurfGetDC()
      m_lpDDS.GetDC m_HDC
End Sub

Public Sub SurfReleaseDC()
      m_lpDDS.ReleaseDC m_HDC
End Sub

Public Sub TextXY(x As Integer, y As Integer, Color As Long, Text As String)
      SetBkMode m_HDC, TRANSPARENT
      SetTextColor m_HDC, Color
      TextOut m_HDC, x, y, Text, Len(Text)
End Sub

Public Sub SetSrc(t As Long, l As Long, b As Long, r As Long)
      With m_SrcRect
            .top = t
            .left = l
            .bottom = b
            .right = r
      End With
End Sub

Public Sub SetDest(t As Long, l As Long, b As Long, r As Long)
      With m_DstRect
            .top = t
            .left = l
            .bottom = b
            .right = r
      End With
End Sub

Public Sub PrivateBlit(x As Integer, y As Integer, Back As IDirectDrawSurface2, Flags As Long)
      m_lpDDS.BltFast x, y, Back, m_SrcRect, DDBLTFAST_WAIT
End Sub

Public Sub LockMe()
      CopyMemory m_ddsd, ByVal 0&, Len(m_ddsd)
      m_ddsd.dwSize = Len(m_ddsd)

      m_lpDDS.Lock ByVal 0&, m_ddsd, DDLOCK_WAIT Or DDLOCK_SURFACEMEMORYPTR, 0

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

Public Sub Pixel(x As Integer, y As Integer, Color As Integer)
      video_buffer(x, y) = Color
End Sub

Public Sub UnLockMe()
      m_lpDDS.Unlock m_ddsd.lpSurface

      CopyMemory ByVal VarPtrArray(video_buffer), ByVal 0&, 4
End Sub
