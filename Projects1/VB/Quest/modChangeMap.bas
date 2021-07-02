Attribute VB_Name = "modChangeMap"

Public Sub SelectPlace()
  'I really dont have idea of improving this.
  
  'Enter
  If PositionMap = "#" Then
    Call Door1
    OutsideMidi = Midi
    Midi = "House.mid"
    Call InitMusic
  End If
  
  'Exit
  If PositionMap = "M" Then
    Call Exit1
    Midi = OutsideMidi
    Call InitMusic
    ItemTownFound1 = False
    ItemTownFound2 = False
    ItemTownFound3 = False
  End If
  
  'Star Gate
  If PositionMap = "@" Then
    Call StarGate1
    Midi = "Town2.mid"
    Call InitMusic
  End If
  
End Sub

Public Sub Door1()

  If MapLoaded = "A1" Then
    If CharY = 7 Then
      If CharX = 10 Then
        Call Map_A2
        MapLoaded = "A2"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 29 Then
        Call Map_A3
        MapLoaded = "A3"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 35 Then
        Call Map_A4
        MapLoaded = "A4"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
  End If
  
  If MapLoaded = "B1" Then
    If CharY = 7 Then
      If CharX = 7 Then
        Call Map_B2
        MapLoaded = "B2"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 10 Then
      If CharX = 8 Then
        Call Map_B3
        MapLoaded = "B3"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 8 Then
      If CharX = 36 Then
        Call Map_B4
        MapLoaded = "B4"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
    If CharY = 9 Then
      If CharX = 21 Then
        Call Map_B5
        MapLoaded = "B5"
        CharY = 20
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
        ItemHouseFound3 = False
      End If
    End If
  End If
  
  If MapLoaded = "B1" Then
    If CharY = 23 Then
      If CharX = 59 Then
        Dim Final(10) As String
        Final(1) = "Guess what theres no temple!! This is the last version of this game."
        Final(2) = "Now you can make your RPG game."
        Final(3) = "And one last thing. If you make a game and you got inspired by this one"
        Final(4) = "try putting me on your credits :P"
        Final(5) = "Thanks for downloading this code :p"
        MsgBox Final(1) & vbCrLf & Final(2) & vbCrLf & Final(3) & vbCrLf & Final(4) & vbCrLf & Final(5) & vbCrLf & "Andres Zacarias" & vbCrLf & "Zacarias Software" & "www.zacarias.mainpage.net"
      End If
    End If
  End If
  
End Sub

Public Sub Exit1()

  If MapLoaded = "A2" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 10
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A3" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 29
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A4" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 35
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B2" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 7
        CharX = 7
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B3" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 10
        CharX = 8
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B4" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 8
        CharX = 36
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "B5" Then
    If CharY = 20 Then
      If CharX = 7 Then
        Call Map_B1
        MapLoaded = "B1"
        CharY = 9
        CharX = 21
        ItemTownFound1 = False
      End If
    End If
  End If

End Sub

Public Sub StarGate1()
  If MapLoaded = "A1" Then
    If CharY = 10 Then
      If CharX = 64 Then
        Call Map_B1
        MapLoaded = "B1"
        SpeechLoaded = "B1"
        CharY = 41
        CharX = 15
        ItemTownFound1 = False
      End If
    End If
  End If
End Sub
