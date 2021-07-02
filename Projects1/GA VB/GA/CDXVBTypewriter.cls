VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBTypewriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Typewriter class. Also requires VBTypeWriter.bas

' X and Y offset of text on the hDC
Public X As Long, Y As Long
' The current character number. Should only be read
Public CurCharNo As Integer

Public Sub NextLine()
    ' Add 12 to Y, moves onto next line for typing
    Y = Y + 12
End Sub

Public Function Draw(hDC As Long, text As String, RGBcolor As Long) As Integer
    ' The actual string to be typed
    Dim theString As String
    ' Number of times the function has needs to be executed
    Static times As Integer
    
    ' Set the typewriter is still typing by default
    Draw = TW_STILLTYPING
    
    ' Increase times the function has been called
    times = times + 1
    ' If we have finished typing the line of text, exit
    If times > Len(text) Then times = 1: Draw = TW_DONETYPING: Exit Function
    
    ' Fill the string to be typed
    theString = Mid(text, 1, times)
    
    ' Set the current character number
    CurCharNo = times
    
    ' Set the text color
    SetTextColor hDC, RGBcolor
    
    ' Output the text onto the hDC
    TextOut hDC, X, Y, theString & Chr(10), Len(theString & Chr(10))
End Function

Private Sub Class_Initialize()
    ' Set the default values
    X = Y = 0
    CurCharNo = 0
End Sub
