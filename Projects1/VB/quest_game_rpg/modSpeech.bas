Attribute VB_Name = "modSpeech"

Public Sub InitializeSpeech()
  
  If CharFacing = 1 Then
    If MapLoaded = "A1" Then
      Call Character_A1
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "A3" Then
      Call Character_A3
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "A4" Then
      Call Character_A4
      SpeechLoaded = "A1"
    End If
    If MapLoaded = "B1" Then
      Call Character_B1
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B2" Then
      Call Character_B2
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B3" Then
      Call Character_B3
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B4" Then
      Call Character_B4
      SpeechLoaded = "B1"
    End If
    If MapLoaded = "B5" Then
      Call Character_B5
      SpeechLoaded = "B1"
    End If
  End If
  
End Sub


Public Sub Speech_A1()

  Message(0) = "Sheik's House."
  Message(1) = "To make spells you will need magic, i think theres some magic in the house."
  Message(2) = "Hey brother, the White Wizard whants to talk to you, you can find him in the mini island."
  Message(3) = "Yep. I can fix this bridge but i will need some wood. Bring it to me and i will do the rest."
  Message(4) = "You whont be able to leave this island without magic. You should find the black wizards hidden in the island's to learn new spells."
  Message(5) = "You have learned the Cut spell. Now you can use this spell by pressing the Z key. Now head east to start your Quest."
  Message(6) = "Feew!!, this weed is to hard. I wish someone could help me."
  Message(7) = "To the mountin of Wizards."
  Message(8) = "To mini Island."
  Message(9) = "To Centurion Island."
 Message(10) = "Lot of weed grows in the bridge. Only those with special powers will leave this island."
 Message(11) = "You found some wood! This thing doesnt has any use to you but you could try to give it to someone."
 Message(12) = "Good. I see that you got wood, ill fix this bridge in no time."
 Message(13) = "Nothing in there."
 Message(14) = "There should be some wood in one of those jars, go ahead and take a peek."
 Message(15) = "My housband left the island to look for a job, but he hasnt come back. Im really worried about this."
 Message(16) = "You found a Coin! With coins you can buy things to get equiped."
 Message(17) = "Thanks Sheik, i really needed your help. Please recive 10 coins as my gratitud."
 Message(18) = "Thanks Sheik, i really needed your help."
 Message(19) = "Star Gate to Centurion Island."
 Message(20) = "You found a bag of Magic! With this your magic increments by 10 and you can make spells."
 Message(21) = "You got no magic, you need to get some more to make your spells."
 
End Sub

Public Sub Speech_B1()

  Message(0) = "Good to see you. In this island you will find the first temple, learn the spell and then you will be able to enter to the temple."
  Message(1) = "Entering the forest of kings."
  Message(2) = "Lost Village."
  Message(3) = "This village is lost. No one has ever found us except for the Black Wizard and you!!"
  Message(4) = "It whont be easy to get with the Black Wizard, there are some obstalces you must reach first to get to him."
  Message(5) = "How the hell you got here!!. O well you are welcome any ways to the Lost Village."
  Message(6) = "Strange. If you look to the right you will see another room, but i dont know how to get there."
  Message(7) = "Me and my troops got lost on the forest and by luck we were rescue by the villagers from this town. As our gratitud we offer them our protection."
  Message(8) = "I wonder how the king is."
  Message(9) = "Were are you heading to??. Ohhh!!. You are the chosen one, you whont be able to pass to the temple without something to prove it."
 Message(10) = "Ahhhh!! im so hungry, i whant some food."
 Message(11) = "You found some wood! This thing doesnt has any use to you but you could try to give it to someone."
 Message(12) = "Give me 50 coins and i will use my sword to cut that tree thats blocking your way."
 Message(13) = "Nothing in there."
 Message(14) = "Thanks!!"
 Message(15) = "Hey. Give me 20 of wood and i will fix the hole of water in the floor."
 Message(16) = "You found a Coin! With coins you can buy things to get equiped."
 Message(17) = "Well thanks for cutting the weeds. Just kidding, you have now learned the Destroy Spell, now you can destroy hills!!. Press the X for this."
 Message(18) = "That rock is not leting my weeds grow. Take it off and i will give you something."
 Message(19) = "Thanks for removing it. Take this Bread as my gratitud."
 Message(20) = "You found a bag of Magic! With this your magic increments by 10 and you can make spells."
 Message(21) = "You got no magic, you need to get some more to make your spells."
 Message(22) = "Temple of Glory."
 Message(23) = "Only royal family may pass to the Temple of Glory."
 Message(24) = "Hey i got my own shop but i cant sell you anything right now. I wonder why nobody comes!!."
 Message(25) = "Thanks for removing it."
 Message(26) = "Wow, thanks kid for the bread!!. Here you will need this Gold Ticket to pass to the Temple of Glory."
 Message(27) = "Ohh!!. This ticket is from the lost troop!!. So they are alive!!, i will let you pass to so you can tell the king. You may stay with the ticket as a prove."
 Message(28) = "ZzzZzz... ZzzZZzz..."
 Message(29) = "Thats good, the troop is alive!!!."
 Message(30) = "How did you find them??."
 
End Sub

Public Sub Character_A1()

  If CharY = 7 Then
    If CharX = 8 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(0)
    End If
  End If
  If CharY = 9 Then
    If CharX = 13 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Mother: " & Message(1)
    End If
  End If
  If CharY = 8 Then
    If CharX = 6 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Brother: " & Message(2)
    End If
  End If
  If CharY = 18 Then
    If CharX = 9 Then
      frmGame.txtSpeech.Visible = True
      If Wood > 0 Then
        frmGame.txtSpeech.Text = "Carpinter: " & Message(12)
        Mid(Map(25), 13, 1) = "G"
        Wood = Wood - 1
      Else
        frmGame.txtSpeech.Text = "Carpinter: " & Message(3)
      End If
    End If
  End If
  If CharY = 19 Then
    If CharX = 30 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "White Wizard: " & Message(4)
    End If
  End If
  If CharY = 35 Then
    If CharX = 31 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Black Wizard: " & Message(5)
      SpellCut = True
    End If
  End If
  If CharY = 10 Then
    If CharX = 32 Then
      If Mid(Map(13), 30, 1) = "G" Then
        If ItemTownFound1 = False Then
          frmGame.txtSpeech.Visible = True
          frmGame.txtSpeech.Text = "Lady: " & Message(17)
          Coin = Coin + 10
          ItemTownFound1 = True
        Else
          frmGame.txtSpeech.Visible = True
          frmGame.txtSpeech.Text = "Lady: " & Message(18)
        End If
      Else
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = "Lady: " & Message(6)
        
      End If
    End If
  End If
  If CharY = 18 Then
    If CharX = 11 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(7)
    End If
  End If
  If CharY = 11 Then
    If CharX = 21 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(8)
    End If
  End If
  If CharY = 8 Then
    If CharX = 38 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(9)
    End If
  End If
  If CharY = 10 Then
    If CharX = 43 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(10)
    End If
  End If
  If CharY = 9 Then
    If CharX = 59 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(19)
    End If
  End If
  
End Sub

Public Sub Character_A3()
  If CharY = 10 Then
    If CharX = 9 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(14)
    End If
  End If
End Sub

Public Sub Character_A4()
  If CharY = 9 Then
    If CharX = 12 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(15)
    End If
  End If
End Sub

Public Sub Character_B1()
  If CharY = 39 Then
    If CharX = 18 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "White Wizard: " & Message(0)
    End If
  End If
  If CharY = 31 Then
    If CharX = 15 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(1)
    End If
  End If
  If CharY = 11 Then
    If CharX = 19 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(2)
    End If
  End If
  If CharY = 6 Then
    If CharX = 10 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(3)
    End If
  End If
  If CharY = 7 Then
    If CharX = 5 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(4)
    End If
  End If
  If CharY = 10 Then
    If CharX = 31 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Boy: " & Message(5)
    End If
  End If
  If CharY = 32 Then
    If CharX = 61 Then
      frmGame.txtSpeech.Visible = True
      If Mid(Map(34), 56, 1) = ">" Then
        frmGame.txtSpeech.Text = "Boy: " & Message(18)
      Else
        If ItemTownFound1 = False Then
          Toast = Toast + 1
          frmGame.txtSpeech.Text = "Boy: " & Message(19)
        Else
          frmGame.txtSpeech.Text = "Boy: " & Message(25)
        End If
      End If
    End If
  End If
  If CharY = 30 Then
    If CharX = 61 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(22)
    End If
  End If
  If CharY = 29 Then
    If CharX = 59 Then
      frmGame.txtSpeech.Visible = True
      If Ticket < 1 Then
        frmGame.txtSpeech.Text = "Soldier: " & Message(23)
      Else
        frmGame.txtSpeech.Text = "Soldier: " & Message(27)
        Mid(Map(31), 61, 1) = "G"
      End If
    End If
  End If
  If CharY = 29 Then
    If CharX = 60 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(28)
    End If
  End If
  If CharY = 26 Then
    If CharX = 58 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(29)
    End If
  End If
  If CharY = 26 Then
    If CharX = 61 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(30)
    End If
  End If
  
End Sub

Public Sub Character_B2()
  If CharY = 10 Then
    If CharX = 9 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(6)
    End If
  End If
End Sub

Public Sub Character_B3()
  If CharY = 10 Then
    If CharX = 6 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(7)
    End If
  End If
  If CharY = 6 Then
    If CharX = 17 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(8)
    End If
  End If
  If CharY = 6 Then
    If CharX = 19 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Soldier: " & Message(9)
    End If
  End If
  If CharY = 6 Then
    If CharX = 21 Then
      frmGame.txtSpeech.Visible = True
      If Toast < 1 Then
        frmGame.txtSpeech.Text = "Soldier: " & Message(10)
      Else
        Ticket = Ticket + 1
        Toast = Toast - 1
        frmGame.txtSpeech.Text = "Soldier: " & Message(26)
      End If
    End If
  End If

End Sub

Public Sub Character_B4()
  
  If CharY = 9 Then
    If CharX = 11 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "White Wizard: " & Message(24)
    End If
  End If

End Sub

Public Sub Character_B5()

  If CharY = 19 Then
    If CharX = 27 Then
      frmGame.txtSpeech.Visible = True
      If Coin < 50 Then
        frmGame.txtSpeech.Text = "Soldier: " & Message(12)
      Else
        Mid(Map(20), 13, 1) = "G"
        Coin = Coin - 50
        frmGame.txtSpeech.Text = "Soldier: " & Message(14)
      End If
    End If
  End If
  
  If CharY = 14 Then
    If CharX = 6 Then
      frmGame.txtSpeech.Visible = True
      If Wood < 20 Then
        frmGame.txtSpeech.Text = "Carpinter: " & Message(15)
      Else
        Mid(Map(16), 15, 1) = "G"
        Coin = Coin - 50
        frmGame.txtSpeech.Text = "Carpinter: " & Message(14)
      End If
    End If
  End If
  
  If CharY = 7 Then
    If CharX = 6 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Black Wizard: " & Message(17)
      SpellDestroy = True
    End If
  End If
  
End Sub
