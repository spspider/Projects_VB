Attribute VB_Name = "DDrawCode"
Public myScreen As New CDXVBScreen
Public myFPS As New CDXVBFPS
Dim MenuBack As New CDXVBSurface
Dim MenuItems(1 To 8) As New CDXVBSurface

Public myMap As CDXVBMap
Public HealthBar As CDXVBSurface
Public LevelBack As CDXVBLayer

Public PausedBack As New CDXVBSurface

Public noise(1) As New CDXVBSurface
Private b1 As Boolean

Public myCredits As New CDXVBCredits

Public myTyper As New CDXVBTypewriter

Public Sub InitDDraw()
    myScreen.HideMouse

    myScreen.CreateFullScreen frmMain.hWnd, 320, 200, 16
    
    MenuBack.Create App.Path & Game.IMAGES_DIR & "back_menu.bmp", myScreen
    MenuItems(1).Create App.Path & Game.IMAGES_DIR & "menu_new_on.bmp", myScreen
    MenuItems(2).Create App.Path & Game.IMAGES_DIR & "menu_new_off.bmp", myScreen
    MenuItems(3).Create App.Path & Game.IMAGES_DIR & "menu_load_on.bmp", myScreen
    MenuItems(4).Create App.Path & Game.IMAGES_DIR & "menu_load_off.bmp", myScreen
    MenuItems(5).Create App.Path & Game.IMAGES_DIR & "menu_save_on.bmp", myScreen
    MenuItems(6).Create App.Path & Game.IMAGES_DIR & "menu_save_off.bmp", myScreen
    MenuItems(7).Create App.Path & Game.IMAGES_DIR & "menu_quit_on.bmp", myScreen
    MenuItems(8).Create App.Path & Game.IMAGES_DIR & "menu_quit_off.bmp", myScreen

    PausedBack.Create App.Path & Game.IMAGES_DIR & "paused.bmp", myScreen

    noise(0).Create App.Path & Game.IMAGES_DIR & "noise1.bmp", myScreen
    noise(1).Create App.Path & Game.IMAGES_DIR & "noise2.bmp", myScreen
    b1 = True
End Sub

Public Sub CloseDDraw()
    myScreen.ShowMouse
End Sub

Public Sub RenderMenu()
    DInputCode.UpdateMenuInput

    myScreen.ClearBack
    MenuBack.Blit 0, 0, myScreen.m_lpDDSBack
    If Game.MenuState = MS_NEW_ON Then
        MenuItems(1).Blit 0, 84, myScreen.m_lpDDSBack
    Else
        MenuItems(2).Blit 0, 84, myScreen.m_lpDDSBack
    End If
    If Game.MenuState = MS_LOAD_ON Then
        MenuItems(3).Blit 0, 114, myScreen.m_lpDDSBack
    Else
        MenuItems(4).Blit 0, 114, myScreen.m_lpDDSBack
    End If
    If Game.MenuState = MS_SAVE_ON Then
        MenuItems(5).Blit 0, 144, myScreen.m_lpDDSBack
    Else
        MenuItems(6).Blit 0, 144, myScreen.m_lpDDSBack
    End If
    If Game.MenuState = MS_QUIT_ON Then
        MenuItems(7).Blit 0, 174, myScreen.m_lpDDSBack
    Else
        MenuItems(8).Blit 0, 174, myScreen.m_lpDDSBack
    End If
    myScreen.Flip
End Sub

Public Sub RenderExit()
    myScreen.ClearBack
        myScreen.SurfGetBackDC
            SetTextAlign myScreen.m_HDC, TA_CENTER
            SetBkMode myScreen.m_HDC, TRANSPARENT
            Credits.myFont.SetFont myScreen.m_HDC
            myCredits.Draw myScreen.m_HDC, myScreen.m_PixelWidth, myScreen.m_PixelHeight
        myScreen.SurfReleaseBackDC
    myScreen.Flip
End Sub

Public Sub InitMap()
    Dim tileadds(1 To 5) As String
    
    For i = LBound(tileadds) To UBound(tileadds)
        tileadds(i) = App.Path & Game.IMAGES_DIR & "tile" & CStr(i) & ".bmp"
    Next i

    Set myMap = New CDXVBMap
    myMap.Create App.Path & Game.LEVELS_DIR & "level1.txt", 50, 10, 320, 200, 5, tileadds, myScreen, 32, 32

    Set HealthBar = New CDXVBSurface
    HealthBar.Create App.Path & Game.IMAGES_DIR & "player_health.bmp", myScreen

    Set LevelBack = New CDXVBLayer
    LevelBack.Create App.Path & Game.IMAGES_DIR & "level1_back.bmp", myScreen, 0, 0
End Sub

Public Sub RenderGame()
    DInputCode.UpdateGameInput

    myScreen.ClearBack
        LevelBack.Draw myScreen.m_lpDDSBack
        myMap.Draw myScreen.m_lpDDSBack, 0, 0, 320, 200
        HealthBar.Blit 5, 5, myScreen.m_lpDDSBack
        
        myScreen.SurfGetBackDC
            SetBkMode myScreen.m_HDC, TRANSPARENT
            SetTextColor myScreen.m_HDC, RGB(0, 0, 180)
            myFont.SetFont myScreen.m_HDC
            TextOut myScreen.m_HDC, 0, 0, CStr(myFPS.ReturnFPS), Len(CStr(myFPS.ReturnFPS))
        myScreen.SurfReleaseBackDC
    myScreen.Flip
End Sub

Public Sub RenderPaused()
    PausedBack.Blit 0, 0, myScreen.m_lpDDSBack
    myScreen.Flip
End Sub

Public Sub RenderChangeLevel()
    Dim nextTime As Long
    Dim curTime As Long
    Dim bFinished As Boolean
    Dim add As Integer
    
    DSoundCode.LoadChangeLevelSounds
    
    bFinished = False
    add = 50

    For i = 0 To 10
        b1 = Not b1
        myScreen.ClearBack
        If b1 Then
            noise(0).Blit 0, 0, myScreen.m_lpDDSBack
        Else
            noise(1).Blit 0, 0, myScreen.m_lpDDSBack
        End If
        myScreen.Flip
        myScreen.m_lpdd.WaitForVerticalBlank DDWAITVB_BLOCKBEGIN, 0
    Next i
    
    nextTime = timeGetTime() + add
    
    Do Until bFinished
        curTime = timeGetTime()
        
        If curTime >= nextTime Then
            myScreen.ClearBack
            myScreen.SurfGetBackDC
                SetBkMode myScreen.m_HDC, TRANSPARENT
                Credits.myFont.SetFont myScreen.m_HDC
                
                res = myTyper.Draw(myScreen.m_HDC, "Prepare for level " & Game.LevelNumber & "...", RGB(0, 255, 0))
        
                DSoundCode.PlayChangeLevelType
        
                If myTyper.CurCharNo = 18 + Len(Game.LevelNumber) Then add = 750
        
                If res = TW_DONETYPING Then
                    bFinished = True
                End If
        
                nextTime = curTime + add
            myScreen.SurfReleaseBackDC
            
            myScreen.Flip
        End If
    Loop
End Sub
