Attribute VB_Name = "modSettings"
Option Explicit

Global Const conDefName = "Неизвестен"
Global Const conDefTime = 999

Private Type TRecord
  Time As Integer
  Name As String
End Type

Global SizeX As Integer
Global SizeY As Integer
Global KolMines As Integer
Global Rang As Integer
Global Rec(2) As TRecord

Public Sub SaveSettings()
  Dim i As Integer
  
  SaveSetting App.EXEName, "Settings", "SizeX", Str$(SizeX)
  SaveSetting App.EXEName, "Settings", "SizeY", Str$(SizeY)
  SaveSetting App.EXEName, "Settings", "KolMines", Str$(KolMines)
  SaveSetting App.EXEName, "Settings", "Rang", Str$(Rang)
  For i = 0 To 2
    SaveSetting App.EXEName, "Records", "Name" & Trim$(Str$(i)), Rec(i).Name
    SaveSetting App.EXEName, "Records", "Time" & Trim$(Str$(i)), Str$(Rec(i).Time)
  Next
End Sub

Public Sub GetSettings()
  Dim i As Integer
  
  SizeX = Int(Val(GetSetting(App.EXEName, "Settings", "SizeX", "8")))
  SizeY = Int(Val(GetSetting(App.EXEName, "Settings", "SizeY", "8")))
  KolMines = Int(Val(GetSetting(App.EXEName, "Settings", "KolMines", "10")))
  Rang = Int(Val(GetSetting(App.EXEName, "Settings", "Rang", "0")))
  If SizeX < 8 Or SizeX > 30 Then SizeX = 8
  If SizeY < 8 Or SizeY > 20 Then SizeY = 8
  If KolMines < 10 Or KolMines > (SizeX - 1) * (SizeY - 1) Then KolMines = 10
  If Rang < 0 Or Rang > 3 Then Rang = 0
  
  For i = 0 To 2
    Rec(i).Name = GetSetting(App.EXEName, "Records", "Name" & Trim$(Str$(i)), conDefName)
    Rec(i).Time = Int(Val(GetSetting(App.EXEName, "Records", "Time" & Trim$(Str$(i)), _
    Str$(conDefTime))))
    If Rec(i).Time < 0 Or Rec(i).Time > 999 Then Rec(i).Time = conDefTime
  Next
End Sub
