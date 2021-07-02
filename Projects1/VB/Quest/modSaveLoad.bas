Attribute VB_Name = "modSaveLoad"
    
  Public SAVE_MapLoaded As String
  Public SAVE_Midi As String
  Public SAVE_OutsideMidi As String
  Public SAVE_SpeechLoaded As String
  Public SAVE_CharX As Integer
  Public SAVE_CharY As Integer
  Public SAVE_CharFacing As Integer
  Public SAVE_Wood As Integer
  Public SAVE_Coin As Integer
  Public SAVE_Magic As Integer
  Public SAVE_Ticket As Integer
  Public SAVE_Toast As Integer
  Public SAVE_SpellCut As String
  Public SAVE_SpellDestroy As String
  

  

Public Sub SaveGame()
  
  SAVE_MapLoaded = MapLoaded
  SAVE_Midi = Midi
  SAVE_OutsideMidi = OutsideMidi
  SAVE_SpeechLoaded = SpeechLoaded
  SAVE_CharX = CharX
  SAVE_CharY = CharY
  SAVE_CharFacing = CharFacing
  SAVE_Wood = Wood
  SAVE_Coin = Coin
  SAVE_Magic = Magic
  SAVE_Ticket = Ticket
  SAVE_Toast = Toast
  SAVE_SpellCut = SpellCut
  SAVE_SpellDestroy = SpellDestroy
  
  
  Open App.Path & "\SaveData.txt" For Output As 1
  Write #1, SAVE_MapLoaded, SAVE_Midi, SAVE_OutsideMidi, SAVE_SpeechLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_SpellCut, SAVE_SpellDestroy
  Close
  MsgBox "Game Saved"
  
End Sub

Public Sub LoadGame()
  
  Open App.Path & "\SaveData.txt" For Input As 1
  Input #1, SAVE_MapLoaded, SAVE_Midi, SAVE_OutsideMidi, SAVE_SpeechLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_SpellCut, SAVE_SpellDestroy
  
    MapLoaded = SAVE_MapLoaded
    OutsideMidi = SAVE_OutsideMidi
    SpeechLoaded = SAVE_SpeechLoaded
    CharX = SAVE_CharX
    CharY = SAVE_CharY
    CharFacing = SAVE_CharFacing
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    Ticket = SAVE_Ticket
    Toast = SAVE_Toast
    SpellCut = SAVE_SpellCut
    SpellDestroy = SAVE_SpellDestroy
    frmGame.Show
  Close
  
End Sub

