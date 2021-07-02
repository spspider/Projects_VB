Attribute VB_Name = "modVariables"


Public X As Integer
Public Y As Integer
Public Map(100) As String
Public PositionMap As String
Public MapLoaded As String

Public Midi As String
Public OutsideMidi As String

Public Message(100)
Public Speaker As String
Public NewLine As String
Public SpeechLoaded As String

Public PassToNext As Integer

Public CharX As Integer
Public CharY As Integer
Public CharFacing As Integer

Public FloorMode As String

'Game Variables
Public SpellCut As Boolean
Public SpellDestroy As Boolean

Public Item As Integer
Public Wood As Integer
Public Coin As Integer
Public Magic As Integer
Public Ticket As Integer
Public Toast As Integer

Public ItemHouseFound1 As Boolean
Public ItemHouseFound2 As Boolean
Public ItemHouseFound3 As Boolean

Public ItemTownFound1 As Boolean
Public ItemTownFound2 As Boolean
Public ItemTownFound3 As Boolean

'******************
'Letter Definition.
'******************

'G = Grass
'B = Bush
'0 = Sign  'Cero
'Q = Top Out Left Border
'A = Bottom Out Left Border
'W = Top Out Right Border
'S = Bottom Out Right Border
'Z = Inn Left Border
'X = Inn Right Border
'z = Inn Top Border
'x = Inn Bottom Border
'- = Water
'_ = Nothing

