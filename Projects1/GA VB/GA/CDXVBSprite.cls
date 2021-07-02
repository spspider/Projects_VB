VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDXVBSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Long
Public y As Long
Public XVel As Long
Public YVel As Long

Public m_lpDDS As New CDXVBSurface
Private Boundry As RECT

Private m_lpdd As CDXVBScreen

Public Sub Create(FileName As String, DDraw As CDXVBScreen, top As Integer, bottom As Integer, left As Integer, right As Integer)
      x = 0
      y = 0

      m_lpDDS.Create FileName, DDraw

      Set m_lpdd = DDraw

      With Boundry
            .top = top
            .bottom = bottom
            .left = left
            .right = right
      End With
End Sub

Private Sub Class_Terminate()
      Set m_lpDDS = Nothing
End Sub

Public Sub Draw()
      Call m_lpDDS.Blit(x, y, m_lpdd.m_lpDDSBack)
End Sub

Public Sub SetXYVel(nXVel As Long, nYVel As Long)
      XVel = nXVel
      YVel = nYVel
End Sub

Public Sub Move()
      If (x + XVel < Boundry.left) Then
            x = Boundry.left
      ElseIf (x + XVel > Boundry.right) Then
            x = Boundry.right
      Else
            x = x + XVel
      End If

      If (y + YVel < Boundry.top) Then
            y = Boundry.top
      ElseIf (y + YVel > Boundry.bottom) Then
            y = Boundry.bottom
      Else
            y = y + YVel
      End If
End Sub

Public Sub SetXY(nX As Long, nY As Long)
      x = nX
      y = nY
End Sub
