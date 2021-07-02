Attribute VB_Name = "DSoundCode"
Public mySound As New CDXVBSound
Public mySndBuffs(0 To 7) As CDXVBSoundBuffer

Public Sub InitDSound()
    mySound.Create frmMain.hWnd
    
    For i = LBound(mySndBuffs) To UBound(mySndBuffs)
        Set mySndBuffs(i) = New CDXVBSoundBuffer
    Next i
End Sub

Public Sub CloseDSound()

End Sub

Public Sub LoadMenuSounds()
    mySndBuffs(0).LoadFromDisk App.Path & Game.SOUND_DIR & "move1.wav", frmMain.hWnd
    mySndBuffs(1).LoadFromDisk App.Path & Game.SOUND_DIR & "ambient1.wav", frmMain.hWnd
End Sub

Public Sub PlayMenuAmbient()
    mySndBuffs(1).Play DSBPLAY_LOOPING
End Sub

Public Sub PlayMenuMoveSound()
    mySndBuffs(0).StopSound
    mySndBuffs(0).Play 0
End Sub

Public Sub StopMenuSounds()
    mySndBuffs(0).StopSound
    mySndBuffs(1).StopSound
End Sub

Public Sub LoadChangeLevelSounds()
    mySndBuffs(0).LoadFromDisk App.Path & Game.SOUND_DIR & "type1.wav", frmMain.hWnd
End Sub

Public Sub PlayChangeLevelType()
    mySndBuffs(0).StopSound
    mySndBuffs(0).Play 0
End Sub
