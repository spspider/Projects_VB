Attribute VB_Name = "modMusicEngine"
Option Explicit

Public Const MCI_NOTIFY_ABORTED = &H4
Public Const MCI_NOTIFY_FAILURE = &H8
Public Const MCI_NOTIFY_SUCCESSFUL = &H1
Public Const MCI_NOTIFY_SUPERSEDED = &H2
Public Const MM_MCINOTIFY = &H3B9  '  MCI
Public Const MM_MCISIGNAL = &H3CB

Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Declare Function mciGetErrorString Lib "winmm.dll" _
    Alias "mciGetErrorStringA" _
    (ByVal dwError As Long, _
    ByVal lpstrBuffer As String, _
    ByVal uLength As Long) As Long

Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)
    
Global lpPrevWndProc As Long 'Used for subclassing
Global gHW As Long 'Handle of window to subclass
Global glFrom As Long 'Start time
Global glTo As Long 'End time
Global gsMidiFile As String 'Filename of Midi file to play
Global gbPlaying As Boolean 'Set to true when a midi file is playing
Global gbRepeat As Boolean 'Flag to know whether to repeat the file when it is done playing


Function MidiNotify(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case MM_MCINOTIFY
            If wParam = MCI_NOTIFY_SUCCESSFUL Then
                If gbRepeat Then
                    PlayMidi
                Else
                    CloseMidi
                End If
                
                'Insert Completed code here.
                Debug.Print "Midi Completed"
                
            End If
    End Select
    
    'All other messages need to be passed back
    'so they can be processed.  Your program will
    'lock up if you don't
    MidiNotify = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

End Function
Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf MidiNotify)
End Sub
Public Sub UnHook()
    Dim Temp As Long
    Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub
Public Sub PlayMidi()
    Dim sFile As String
    Dim sShortFile As String * 67
    Dim lResult As Integer
    Dim sError As String * 255
    Dim sCommand As String
    
    'Set the path and filename to open.  I am using the
    'mcitest.mid which I found in my VB5 directory in
    'the sub folders samples\comptool\mci
    'I just copied it to this projects folder.
    
    'The mciSendString API call doesn't seem to like'
    'long filenames that have spaces in them, so we
    'will make another API call to get the short
    'filename version.
    lResult = GetShortPathName(gsMidiFile, sShortFile, _
                               Len(sShortFile))
    gsMidiFile = Left$(sShortFile, lResult)
    
    If Not gbPlaying Then
        'Make the call to open the midi file and assign
        'it an alias
        lResult = mciSendString("open " & gsMidiFile & " type sequencer", ByVal 0&, 0, gHW) '0&) 'AddressOf MidiNotify)
        
        'Check to see if there was an error
        If lResult Then
            lResult = mciGetErrorString(lResult, sError, 255)
            Debug.Print "open: " & CStr(lResult) & ": " & sError
            Exit Sub
        End If
    End If
    
    'Make the call to start playing the midi
    If glFrom = 0 And glTo = 0 Then
        sCommand = "play " & gsMidiFile & " from 0 notify"
    Else
        sCommand = "play " & gsMidiFile & " from " & CStr(glFrom) & " to " & CStr(glTo) & " notify"
    End If
    lResult = mciSendString(sCommand, ByVal 0&, 0, gHW)
    
    'Check to see if there were any errors
    If lResult Then
        lResult = mciGetErrorString(lResult, sError, 255)
        Debug.Print "play: " & CStr(lResult) & ": " & sError
    Else
        gbPlaying = True
    End If

End Sub

Public Sub CloseMidi()
    Dim lResult As Integer
    Dim sError As String * 255
    
    'Make the call to close the midi file
    lResult = mciSendString("close " & gsMidiFile, " notify", 0&, 0&)
    
    gbPlaying = False
End Sub



Public Sub InitMusic()
    
    CloseMidi
    UnHook

    'Set your global variables. You can find these
    'declared in the MidiAPI module
    'gHW = frmGame.hWnd
    gsMidiFile = App.Path & "\Music\" & Midi
    gbRepeat = True  'Set to true if you want it to repeat
    
    'You can specify the start and stop times with
    'these variables.  To Play the whole song, just
    'set both glFrom and glTo to 0.
    glFrom = 0
    'glTo = 15 'Will play for a couple of seconds
    
    'Go ahead and make sure no other midi file is open
    CloseMidi
    UnHook
    DoEvents
    
    'It is important to hook the form after
    'you begin playing the Midi.  Not sure why.
    PlayMidi
    Hook
   
End Sub

