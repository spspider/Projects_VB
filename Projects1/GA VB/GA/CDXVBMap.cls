VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Only works for 64*64 tiles at the moment, and only a screen res of 640*480

Public px, py As Integer
Public bx, by As Integer

Public X_BOUND As Long
Public Y_BOUND As Long

Private Map() As Integer
Private Tiles() As New CDXVBSurface

Public m_TileWidth As Long
Public m_TileHeight As Long

Public Sub Create(FN As String, MapWidth As Integer, MapHeight As Integer, ScreenWidth As Integer, ScreenHeight As Integer, NoOfTiles As Integer, TileAdds() As String, DDraw As CDXVBScreen, TileW As Long, TileH As Long)
      Dim char As String
      
      ' Resize the map to given size
      ReDim Map(MapWidth, MapHeight)
      ' Resize the tiles array to the number of tiles there are
      ReDim Tiles(1 To NoOfTiles)
      
      m_TileWidth = TileW
      m_TileHeight = TileH
      
      ' Load map file into array
      Open FN For Binary As #1
            For Y = 0 To MapHeight
                  For X = 0 To MapWidth
                        char = " "
                        Get #1, 1 + Y * (MapWidth + 2) + X, char
                        
                        If IsNumeric(char) Then Map(X, Y) = char
                  Next X
            Next Y
      Close #1
      
      ' Set map boundaries
      X_BOUND = MapWidth * m_TileWidth
      Y_BOUND = MapHeight * m_TileHeight
      
      ' Set screen boundaries
      px = 0
      py = 0
      bx = ScreenWidth
      by = ScreenHeight
      
      ' Load up bitmaps into tiles
      For i = 1 To UBound(Tiles, 1)
            Tiles(i).Create TileAdds(i), DDraw
      Next i
End Sub

Public Sub Draw(Back As IDirectDrawSurface2, ClipLeft As Long, ClipTop As Long, ClipRight As Long, ClipBottom As Long)
      ' Height in tiles
      For Y = Int(py / m_TileHeight) To Int((by - 1) / m_TileHeight)
            ' Width in tiles
            For X = Int(px / m_TileWidth) To Int((bx - 1) / m_TileWidth)
                  If Not Tiles(Map(X, Y)) Is Nothing Then
                        DrawTile Back, Tiles(Map(X, Y)), X * m_TileWidth - px, Y * m_TileHeight - py, ClipLeft, ClipTop, ClipRight, ClipBottom
                  End If
            Next X
      Next Y
End Sub

Private Sub DrawTile(Dest As IDirectDrawSurface2, Tile As CDXVBSurface, X As Long, Y As Long, ClipLeft As Long, ClipTop As Long, ClipRight As Long, ClipBottom As Long)
      Dim rcRect As RECT
      Dim Diff As Long
      
      With rcRect
            .top = 0
            .left = 0
            .bottom = m_TileHeight
            .right = m_TileWidth
      End With
      
      ' Check to see if the tile is on the screen
      If X + rcRect.right < ClipLeft Then Exit Sub
      If X > ClipRight Then Exit Sub
      If Y + rcRect.bottom < ClipTop Then Exit Sub
      If Y > ClipBottom Then Exit Sub
      
      ' Left side
      If X < ClipLeft Then
            Diff = Abs(X - ClipLeft)
            rcRect.left = rcRect.left + Diff
            X = ClipLeft
      End If
      ' Top side
      If Y < ClipTop Then
            Diff = Abs(Y - ClipTop)
            rcRect.top = rcRect.top + Diff
            Y = ClipTop
      End If
      ' Right side
      If X + rcRect.right > ClipRight Then rcRect.right = ClipRight - X + rcRect.left
      ' Bottom side
      If Y + rcRect.bottom > ClipBottom Then rcRect.bottom = ClipBottom - Y + rcRect.top

      ' Error fixing
      If rcRect.bottom = 0 Then Exit Sub
      If rcRect.right = 0 Then Exit Sub
      If Not rcRect.bottom > rcRect.top Then Exit Sub
      If Not rcRect.right > rcRect.left Then Exit Sub

      ' Set clipping rectangle
      Tile.SetSrc rcRect.top, rcRect.left, rcRect.bottom, rcRect.right

      Tile.Blit X, Y, Dest
End Sub

Public Sub MoveUp(Amt As Integer)
      py = py - Amt
      by = by - Amt

      If (py < 0) Then py = 0: by = by + Amt
End Sub

Public Sub MoveDown(Amt As Integer)
      py = py + Amt
      by = by + Amt

      If (by > Y_BOUND) Then by = Y_BOUND: py = py - Amt
End Sub

Public Sub MoveLeft(Amt As Integer)
      px = px - Amt
      bx = bx - Amt

      If (px < 0) Then px = 0: bx = bx + Amt
End Sub

Public Sub MoveRight(Amt As Integer)
      px = px + Amt
      bx = bx + Amt

      If (bx > X_BOUND) Then bx = X_BOUND: px = px - Amt
End Sub

Public Sub Clear()
      For X = 0 To UBound(Map, 1)
            For Y = 0 To UBound(Map, 2)
                  Map(X, Y) = 0
            Next Y
      Next X
End Sub

Public Sub Fill(TileNo As Integer)
      For X = 0 To UBound(Map, 1)
            For Y = 0 To UBound(Map, 2)
                  Map(X, Y) = TileNo
            Next Y
      Next X
End Sub

Public Sub PutTile(X As Integer, Y As Integer, TileNo As Integer)
      Map(X, Y) = TileNo
End Sub

Public Function GetTile(X As Integer, Y As Integer) As Integer
      GetTile = Map(X, Y)
End Function

