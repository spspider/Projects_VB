VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_XOffset As Long
Public m_YOffset As Long

Public m_PixelWidth As Long
Public m_PixelHeight As Long

Public m_ScreenX As Long
Public m_ScreenY As Long

Public m_lpDDS As New CDXVBSurface
Private m_Debug As New CDXVBINI
Private m_rcRect As RECT

Public Sub Create(FN As String, Screen As CDXVBScreen, ScreenX As Long, ScreenY As Long)
      m_lpDDS.Create FN, Screen

      m_XOffset = 1
      m_YOffset = 1
      
      m_ScreenX = ScreenX
      m_ScreenY = ScreenY
      
      m_PixelWidth = CLng(m_lpDDS.m_PixelWidth)
      m_PixelHeight = CLng(m_lpDDS.m_PixelHeight)
      
      m_Debug.Create "DEBUG INFO", App.Path & "\Layer.ini"
End Sub

Public Sub Draw(Back As IDirectDrawSurface2)
      With m_rcRect
            .top = m_PixelHeight - m_YOffset
            .left = m_PixelWidth - m_XOffset
            .bottom = m_PixelHeight
            .right = m_PixelWidth
      End With
      If m_rcRect.bottom = 0 Then GoTo No2
      If m_rcRect.right = 0 Then GoTo No2
      If Not m_rcRect.bottom > m_rcRect.top Then GoTo No2
      If Not m_rcRect.right > m_rcRect.left Then GoTo No2
      Back.BltFast m_ScreenX, m_ScreenY, m_lpDDS.m_lpDDS, m_rcRect, DDBLTFAST_WAIT
      
No2:
      
      With m_rcRect
            .top = m_PixelHeight - m_YOffset
            .left = 0
            .bottom = m_PixelHeight
            .right = m_PixelWidth - m_XOffset
      End With
      If m_rcRect.bottom = 0 Then GoTo No3
      If m_rcRect.right = 0 Then GoTo No3
      If Not m_rcRect.bottom > m_rcRect.top Then GoTo No3
      If Not m_rcRect.right > m_rcRect.left Then GoTo No3
      Back.BltFast m_ScreenX + m_XOffset, m_ScreenY, m_lpDDS.m_lpDDS, m_rcRect, DDBLTFAST_WAIT
      
No3:
      
      With m_rcRect
            .top = 0
            .left = m_PixelWidth - m_XOffset
            .bottom = m_PixelHeight - m_YOffset
            .right = m_PixelWidth
      End With
      If m_rcRect.bottom = 0 Then GoTo No4
      If m_rcRect.right = 0 Then GoTo No4
      If Not m_rcRect.bottom > m_rcRect.top Then GoTo No4
      If Not m_rcRect.right > m_rcRect.left Then GoTo No4
      Back.BltFast m_ScreenX, m_ScreenY + m_YOffset, m_lpDDS.m_lpDDS, m_rcRect, DDBLTFAST_WAIT
      
No4:
      
      With m_rcRect
            .top = 0
            .left = 0
            .bottom = m_PixelHeight - m_YOffset
            .right = m_PixelWidth - m_XOffset
      End With
      If m_rcRect.bottom = 0 Then Exit Sub
      If m_rcRect.right = 0 Then Exit Sub
      If Not m_rcRect.bottom > m_rcRect.top Then Exit Sub
      If Not m_rcRect.right > m_rcRect.left Then Exit Sub
      Back.BltFast m_ScreenX + m_XOffset, m_ScreenY + m_YOffset, m_lpDDS.m_lpDDS, m_rcRect, DDBLTFAST_WAIT
End Sub

Public Sub ScrollDown(Amt As Long)
      m_YOffset = m_YOffset - Amt
      
      If m_YOffset < 0 Then m_YOffset = m_PixelHeight + m_YOffset
End Sub

Public Sub ScrollUp(Amt As Long)
      m_YOffset = m_YOffset + Amt
      
      If m_YOffset > m_PixelHeight Then m_YOffset = m_YOffset - m_PixelHeight
End Sub

Public Sub ScrollLeft(Amt As Integer)
      m_XOffset = m_XOffset - Amt
      
      If m_XOffset < 0 Then m_XOffset = m_XOffset + m_PixelWidth
End Sub

Public Sub ScrollRight(Amt As Integer)
      m_XOffset = m_XOffset + Amt
      
      If m_XOffset > m_PixelWidth Then m_XOffset = m_XOffset - m_PixelWidth
End Sub
