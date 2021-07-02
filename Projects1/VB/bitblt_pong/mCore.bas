Attribute VB_Name = "mCore"
Option Explicit
Global ShotEnabled
'Types
    Type tPlayer
        x As Long 'Position X
        y As Long 'Position Y
        w As Long 'Size X
        h As Long 'Size Y
        
        DC As Long 'hDC of PictureBox
    End Type
    
    Type tPlayer2
        x As Long 'Position X
        y As Long 'Position Y
        w As Long 'Size X
        h As Long 'Size Y
        
        DC As Long 'hDC of PictureBox
    End Type
    
    Type tTile
        w As Long 'Size X
        h As Long 'Size Y
        
        DC As Long 'hDC of PictureBox
    End Type
    
    Type tBall
        x As Long 'Position X
        y As Long 'Position Y
        w As Long 'Size X
        h As Long 'Size Y
        
        DC As Long 'hDC of PictureBox
    End Type

'API
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function GetTickCount Lib "kernel32" () As Long
'Vars
    Global BackDC As Long 'hDC of PictureBox
    Global FrontDC As Long 'hDC of PictureBox
    
    Global Tile As tTile 'Tiles... in our case just grass
    Global Player As tPlayer 'The player
    Global Ball As tBall 'The Ball
    Global Player2 As tPlayer2
    'Customizable vars
    Global DX As Integer
    Global DY As Integer
    Global Life As Integer
    Global Difficulty As String
    Global Score As String
