Attribute VB_Name = "DInputCode"
Dim myInput As New CDXVBInput

Public Sub InitDInput()
    myInput.Create App.hInstance, frmMain.hWnd
    myInput.UnAcquire
    myInput.ReAcquire
End Sub

Public Sub CloseDInput()

End Sub

Public Sub UpdateMenuInput()
    myInput.UpdateKeyboard
    
    If Keys(DIK_UP) Then
        Game.MenuState = Game.MenuState - 1
        If Game.MenuState < Game.MENU_LOWEST Then Game.MenuState = Game.MENU_HIGHEST
        DSoundCode.PlayMenuMoveSound
    End If
    
    If Keys(DIK_DOWN) Then
        Game.MenuState = Game.MenuState + 1
        If Game.MenuState > Game.MENU_HIGHEST Then Game.MenuState = Game.MENU_LOWEST
        DSoundCode.PlayMenuMoveSound
    End If
    
    If Keys(DIK_RETURN) Then
        If Game.MenuState = Game.MS_NEW_ON Then Game.GameState = Game.GS_NEWGAME: DSoundCode.StopMenuSounds: MIDIMusic.myMusic.StopPlaying
        If Game.MenuState = Game.MS_QUIT_ON Then
            Game.GameState = Game.GS_EXITING
            MIDIMusic.PlayOuttro
            DSoundCode.StopMenuSounds
        End If
        DSoundCode.PlayMenuMoveSound
    End If
End Sub

Public Sub UpdateGameInput()
    myInput.UpdateKeyboard
    
    If Keys(DIK_P) Then
        Game.GameState = Game.GS_PAUSED
    End If
    
    If Keys(DIK_UP) Then
        myMap.MoveUp 4
        LevelBack.ScrollUp 2
    End If
    If Keys(DIK_DOWN) Then
        myMap.MoveDown 4
        LevelBack.ScrollDown 2
    End If
    If Keys(DIK_LEFT) Then
        myMap.MoveLeft 4
        LevelBack.ScrollRight 2
    End If
    If Keys(DIK_RIGHT) Then
        myMap.MoveRight 4
        LevelBack.ScrollLeft 2
    End If
End Sub
