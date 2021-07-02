Attribute VB_Name = "MIDIMusic"
Public myMusic As New CDXVBMusic
Public CurrentFile As String

Public Sub PlayIntro()
    Dim temp(1) As String
    Dim lineo As Integer
    
    lineo = 0

    Open App.Path & Game.CONFIG_DIR & "music.txt" For Input As #1
        Do Until EOF(1)
            Input #1, temp(lineno)
            lineno = lineo + 1
        Loop
    Close #1
    
    myMusic.Init frmMain.hWnd
    myMusic.Play App.Path & Game.MUSIC_DIR & temp(0)
    CurrentFile = temp(0)
End Sub

Public Sub PlayOuttro()
    Dim temp(1) As String
    Dim lineo As Integer
    
    lineo = 0

    Open App.Path & Game.CONFIG_DIR & "music.txt" For Input As #1
        Do Until EOF(1)
            Input #1, temp(lineno)
            lineno = lineo + 1
        Loop
    Close #1
    
    myMusic.StopPlaying
    myMusic.Play App.Path & Game.MUSIC_DIR & temp(1)
    CurrentFile = temp(1)
End Sub

Public Sub PlayLevelMusic(levelnumber As Integer)
    Dim temp As String

    Open App.Path & Game.LEVELS_DIR & "level" & levelnumber & "def.txt" For Input As #1
        Input #1, temp
    Close #1
    
    myMusic.StopPlaying
    myMusic.Play App.Path & Game.MUSIC_DIR & temp
    curentfile = temp
End Sub
