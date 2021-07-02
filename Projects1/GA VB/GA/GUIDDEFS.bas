Attribute VB_Name = "GUIDDEFS"
' Game controls objects
Public GUID_XAxis As Win32.GUID
Public GUID_YAxis As Win32.GUID
Public GUID_ZAxis As Win32.GUID
Public GUID_RxAxis As Win32.GUID
Public GUID_RyAxis As Win32.GUID
Public GUID_RzAxis As Win32.GUID
Public GUID_Slider As Win32.GUID
Public GUID_Button As Win32.GUID
Public GUID_Key As Win32.GUID
Public GUID_POV As Win32.GUID
Public GUID_Unknown As Win32.GUID

' Input devices
Public GUID_SysMouse As Win32.GUID
Public GUID_SysKeyboard As Win32.GUID
Public GUID_Joystick As Win32.GUID

' Data formats
Public c_dfDIMouse As DIDATAFORMAT
Public c_dfDIKeyboard As DIDATAFORMAT
Public c_dfDIJoystick As DIDATAFORMAT
Public c_dfDIJoystick2 As DIDATAFORMAT

' Private
Private odfMouse(0 To 6) As DIOBJECTDATAFORMAT
Private odfKeyBoard(0 To 255) As DIOBJECTDATAFORMAT
Private odfJoystick(0 To 43) As DIOBJECTDATAFORMAT
Private odfJoystick2(0 To 163) As DIOBJECTDATAFORMAT

' These are needed by CDXVBInput... Look into these structures for info!
Public MouseState As DIMOUSESTATE
Public JoystickState As DIJOYSTATE
Public Keys(0 To 255) As Byte

' Fills a GUID
Private Sub DefineGUID(GUID As Win32.GUID, ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, ByVal Data40 As Byte, ByVal Data41 As Byte, ByVal Data42 As Byte, ByVal Data43 As Byte, ByVal Data44 As Byte, ByVal Data45 As Byte, ByVal Data46, ByVal Data47 As Byte)
    With GUID
        .Data1 = Data1
        .Data2 = Data2
        .Data3 = Data3
        .Data4(0) = Data40
        .Data4(1) = Data41
        .Data4(2) = Data42
        .Data4(3) = Data43
        .Data4(4) = Data44
        .Data4(5) = Data45
        .Data4(6) = Data46
        .Data4(7) = Data47
    End With
End Sub

Public Sub GUID_Initialize()
    Dim i As Long
    ' Game controls
    DefineGUID GUID_XAxis, &HA36D02E0, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_YAxis, &HA36D02E1, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_ZAxis, &HA36D02E2, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_RxAxis, &HA36D02F4, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_RyAxis, &HA36D02F5, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_RzAxis, &HA36D02E3, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_Slider, &HA36D02E4, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_Button, &HA36D02F0, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_Key, &H55728220, &HD33C, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_POV, &HA36D02F2, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_Unknown, &HA36D02F3, &HC9F3, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    
    ' Input devices
    DefineGUID GUID_SysMouse, &H6F1D2B60, &HD5A0, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_SysKeyboard, &H6F1D2B61, &HD5A0, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    DefineGUID GUID_Joystick, &H6F1D2B70, &HD5A0, &H11CF, &HBF, &HC7, &H44, &H45, &H53, &H54, 0, 0
    
    ' Mouse objects array
    With odfMouse(0)
        .dwOfs = 0
        .dwType = &HFFFF00 + DIDFT_AXIS
        .dwflags = 0
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfMouse(1)
        .dwOfs = 4
        .dwType = &HFFFF00 + DIDFT_AXIS
        .dwflags = 0
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfMouse(2)
        .dwOfs = 8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 0
        .pGUID = VarPtr(GUID_ZAxis)
    End With
    For i = 0 To 3
        With odfMouse(3 + i)
            .dwOfs = 12 + i
            .dwType = &HFFFF00 + DIDFT_BUTTON
            If i = 2 Or i = 3 Then .dwType = &H80000000 Or .dwType
            .dwflags = 0
            .pGUID = VarPtr(GUID_Button)
        End With
    Next
    
    ' Mouse data format
    With c_dfDIMouse
        .dwSize = &H18
        .dwObjSize = &H10
        .dwflags = 2
        .dwDataSize = &H10
        .dwNumObjs = 7
        .rgodf = VarPtr(odfMouse(0))
    End With
    
    ' Keyboard objects array
    For i = 0 To 255
        With odfKeyBoard(i)
            .dwOfs = i
            .dwType = &H80000000 + DIDFT_BUTTON + i * 256&
            .dwflags = &H0
            .pGUID = VarPtr(GUID_Key)
        End With
    Next
    
    ' Keyboard data format
    With c_dfDIKeyboard
        .dwSize = &H18
        .dwObjSize = &H10
        .dwflags = 2
        .dwDataSize = &H100
        .dwNumObjs = &H100
        .rgodf = VarPtr(odfKeyBoard(0))
    End With
    
    ' Joystick objects array
    With odfJoystick(0)
        .dwOfs = 0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfJoystick(1)
        .dwOfs = 4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick(2)
        .dwOfs = 8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_ZAxis)
    End With
    With odfJoystick(3)
        .dwOfs = &HC
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_RxAxis)
    End With
    With odfJoystick(4)
        .dwOfs = &H10
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_RyAxis)
    End With
    With odfJoystick(5)
        .dwOfs = &H14
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_RzAxis)
    End With
    With odfJoystick(6)
        .dwOfs = &H18
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick(7)
        .dwOfs = &H1C
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick(8)
        .dwOfs = &H20
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick(9)
        .dwOfs = &H24
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick(10)
        .dwOfs = &H28
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick(11)
        .dwOfs = &H2C
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    For i = 0 To 31
    With odfJoystick(12 + i)
        .dwOfs = &H30 + i
        .dwType = &H80FFFF00 + DIDFT_BUTTON
        .dwflags = 0
        .pGUID = VarPtr(GUID_Button)
    End With
    Next
    ' Joystick data format
    With c_dfDIJoystick
        .dwSize = &H18
        .dwObjSize = &H10
        .dwflags = 1
        .dwDataSize = &H50
        .dwNumObjs = &H2C
        .rgodf = VarPtr(odfJoystick(0))
    End With
    
    ' Joystick2
    With odfJoystick2(0)
        .dwOfs = 0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfJoystick2(1)
        .dwOfs = 4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(2)
        .dwOfs = 8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(3)
        .dwOfs = 12
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(4)
        .dwOfs = 16
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(5)
        .dwOfs = 20
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(6)
        .dwOfs = 24
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(7)
        .dwOfs = 28
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 1
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(8)
        .dwOfs = 32
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick2(9)
        .dwOfs = 36
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick2(10)
        .dwOfs = 40
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    With odfJoystick2(11)
        .dwOfs = 44
        .dwType = &H80FFFF00 + DIDFT_POV
        .dwflags = 0
        .pGUID = VarPtr(GUID_POV)
    End With
    For i = 0 To 127
        With odfJoystick2(12 + i)
            .dwOfs = &H30 + i
            .dwType = &H80FFFF00 + DIDFT_BUTTON
            .dwflags = 0
            .pGUID = VarPtr(GUID_Key)
        End With
    Next
    With odfJoystick2(140)
        .dwOfs = &HB0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfJoystick2(141)
        .dwOfs = &HB4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(142)
        .dwOfs = &HB8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_ZAxis)
    End With
    With odfJoystick2(143)
        .dwOfs = &HBC
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_RxAxis)
    End With
    With odfJoystick2(144)
        .dwOfs = &HC0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_RyAxis)
    End With
    With odfJoystick2(145)
        .dwOfs = &HC4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_RzAxis)
    End With
    With odfJoystick2(146)
        .dwOfs = &H18
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick2(147)
        .dwOfs = &H1C
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 2
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick2(148)
        .dwOfs = &HD0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfJoystick2(149)
        .dwOfs = &HD4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(150)
        .dwOfs = &HD8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_ZAxis)
    End With
    With odfJoystick2(151)
        .dwOfs = &HDC
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_RxAxis)
    End With
    With odfJoystick2(152)
        .dwOfs = &HE0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_RyAxis)
    End With
    With odfJoystick2(153)
        .dwOfs = &HE4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_RzAxis)
    End With
    With odfJoystick2(154)
        .dwOfs = &H18
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick2(155)
        .dwOfs = &H1C
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 3
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick2(156)
        .dwOfs = &HF0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_XAxis)
    End With
    With odfJoystick2(157)
        .dwOfs = &HF4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_YAxis)
    End With
    With odfJoystick2(158)
        .dwOfs = &HF8
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_ZAxis)
    End With
    With odfJoystick2(159)
        .dwOfs = &HFC
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_RxAxis)
    End With
    With odfJoystick2(160)
        .dwOfs = &H0
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_RyAxis)
    End With
    With odfJoystick2(161)
        .dwOfs = &H4
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_RzAxis)
    End With
    With odfJoystick2(162)
        .dwOfs = &H18
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_Slider)
    End With
    With odfJoystick2(163)
        .dwOfs = &H1C
        .dwType = &H80FFFF00 + DIDFT_AXIS
        .dwflags = 4
        .pGUID = VarPtr(GUID_Slider)
    End With
    With c_dfDIJoystick2
        .dwSize = &H18
        .dwObjSize = &H10
        .dwflags = 1
        .dwDataSize = &H110
        .dwNumObjs = &HA4
        .rgodf = VarPtr(odfJoystick2(0))
    End With
End Sub
