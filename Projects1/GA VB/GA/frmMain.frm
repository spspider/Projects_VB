VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "APE Example"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If key ESC is pressed
    If KeyAscii = 27 Then
        ' If playing then pause the game
        If Game.GameState = Game.GS_PLAYING Then Game.GameState = Game.GS_MENU: MIDIMusic.PlayIntro: DSoundCode.LoadMenuSounds: DSoundCode.PlayMenuAmbient
        ' If paused, then resume the game
        If Game.GameState = Game.GS_PAUSED Then Game.GameState = Game.GS_PLAYING
        ' If we are exiting the game, quit completely
        If Game.GameState = Game.GS_EXITING Then Unload Me
    End If
End Sub

Private Sub Form_Load()
    ' Current time and next time, frame limiter
    Dim curTime As Long, nextTime As Long

    ' Make sure the form is shown first
    Me.Show
    
    ' Initialise DirectDraw
    DDrawCode.InitDDraw
    ' Set up the credits scroller
    DDrawCode.myCredits.PosY = DDrawCode.myScreen.m_PixelHeight
    ' Load the credits text file
    Game.LoadCreditsTXT
    ' Create fixedsys font
    Credits.myFont.CreateNewFont "Fixedsys"
    
    ' Initialise GUID's for DInput
    GUIDDEFS.GUID_Initialize
    ' Initialise DirectInput
    DInputCode.InitDInput
    
    ' Initialise DirectSound
    DSoundCode.InitDSound
    ' Load the sounds for the menu
    DSoundCode.LoadMenuSounds
    ' Play ambient music in the menu
    DSoundCode.PlayMenuAmbient
    
    ' Set the game's current state to MENU
    Game.GameState = Game.GS_MENU
    ' Set the menu state to being on the "new" button
    Game.MenuState = Game.MS_NEW_ON
    
    ' Start up the MIDI music
    MIDIMusic.PlayIntro
    
    ' No level yet
    Game.levelnumber = 0
    
    ' Menu time delay
    Game.TimeAddOn = 200
    
    ' Frame limiting
    nextTime = timeGetTime() + Game.TimeAddOn
    
    ' Main game loop
    Do While DoEvents()
        ' Get the current time
        curTime = timeGetTime()
        
        ' Compare current time to when loop should be executed
        If curTime >= nextTime Then
        
            ' Calculate FPS
            myFPS.GetFPS
            
            ' Render the game according to the game's state
            Select Case Game.GameState
                ' Render the menu
                Case Game.GS_MENU:
                    Game.TimeAddOn = 200
                    DDrawCode.RenderMenu
                ' Render the paused screen
                Case Game.GS_PAUSED:
                    DDrawCode.RenderPaused
                ' Render the actual game (map and stuff)
                Case Game.GS_PLAYING:
                    Game.TimeAddOn = 10
                    DDrawCode.RenderGame
                ' Render exit screen
                Case Game.GS_EXITING:
                    Game.TimeAddOn = 50
                    DDrawCode.RenderExit
                ' Load the next level
                Case Game.GS_LOADNEXTLEVEL:
                    Game.levelnumber = Game.levelnumber + 1
                    DDrawCode.RenderChangeLevel
                    MIDIMusic.PlayLevelMusic Game.levelnumber
                    Game.GameState = Game.GS_PLAYING
                ' Set up an entirely new game
                Case Game.GS_NEWGAME:
                    Game.TimeAddOn = 100
                    DDrawCode.InitMap
                    Game.levelnumber = 0
                    Game.GameState = Game.GS_LOADNEXTLEVEL
            End Select
            
            ' Calculate next time loop is executed
            nextTime = curTime + Game.TimeAddOn
        End If
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Close down all of the DirectX stuff
    DDrawCode.CloseDDraw
    DInputCode.CloseDInput
    DSoundCode.CloseDSound
    MIDIMusic.myMusic.StopPlaying
End Sub
