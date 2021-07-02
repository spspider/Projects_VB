Attribute VB_Name = "Game"
Public Const GS_MENU = 0
Public Const GS_PLAYING = 1
Public Const GS_PAUSED = 2
Public Const GS_NEWGAME = 3
Public Const GS_LOADNEXTLEVEL = 4
Public Const GS_EXITING = 5
Public GameState As Integer

Public Const MS_NEW_ON = 0
Public Const MS_LOAD_ON = 1
Public Const MS_SAVE_ON = 2
Public Const MS_QUIT_ON = 3
Public Const MENU_LOWEST = 0
Public Const MENU_HIGHEST = 3
Public MenuState As Integer

Public Const IMAGES_DIR = "\StaticBitmaps\"
Public Const CONFIG_DIR = "\Config\"
Public Const LEVELS_DIR = "\Levels\"
Public Const MUSIC_DIR = "\Music\"
Public Const SOUND_DIR = "\Sounds\"

Public LevelNumber As Integer

Public TimeAddOn As Integer

Public Sub LoadCreditsTXT()
    ' Resize the strings array to no strings
    ReDim Credits.Strings(0)

    ' Open the text file
    Open App.Path & CONFIG_DIR & "\credits.txt" For Input As #1
        ' Repeat until the end of the file
        Do Until EOF(1)
            ' Resize the array preserving the current entries
            ReDim Preserve Credits.Strings(UBound(Credits.Strings) + 1)
            ' Load in the string
            Input #1, Credits.Strings(UBound(Credits.Strings))
        Loop
    Close #1
End Sub

Public Sub SaveGame()

End Sub

Public Sub LoadGame()

End Sub
