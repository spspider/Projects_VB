VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDXVBINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function kpGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
Private Declare Function kpGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Private Declare Function kpWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpAppName As Any, ByVal lpKeyName As Any, ByVal lpString As Any) As Integer
Private Declare Function kpGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Private Declare Function kpGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Private Declare Function kpWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
Private Declare Function kpSendMessage Lib "User32" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Private Declare Function kpGetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectory" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Public smSectionName As String   'Current section in private Ini file
Public smIniFileName As String   'Fully qualified path/name of current private Ini file
Public nmPrivInit As Integer     'Flag to indicate that Private.Ini is initialized

Private Const Max_SectionBuffer = 4096
Private Const Max_EntryBuffer = 255

Public Function ExtractName$(sSpecIn$, nBaseOnly%)
  Dim nCnt%, nDot%, sSpecOut$

  On Local Error Resume Next
  
  If InStr(sSpecIn, "\") Then
    For nCnt = Len(sSpecIn) To 1 Step -1
      If Mid$(sSpecIn, nCnt, 1) = "\" Then
        sSpecOut = Mid$(sSpecIn, nCnt + 1)
        Exit For
      End If
    Next nCnt
  
  ElseIf InStr(sSpecIn, ":") = 2 Then
    sSpecOut = Mid$(sSpecIn, 3)
    
  Else
    sSpecOut = sSpecIn
  End If
    
  If nBaseOnly Then
    nDot = InStr(sSpecOut, ".")
    If nDot Then
      sSpecOut = Left$(sSpecOut, nDot - 1)
    End If
  End If

  ExtractName$ = UCase$(sSpecOut)

End Function

Public Function ExtractPath$(sSpecIn$)

  Dim nCnt%, sSpecOut$
  
  On Local Error Resume Next

  If InStr(sSpecIn, "\") Then
    For nCnt = Len(sSpecIn) To 1 Step -1
      If Mid$(sSpecIn, nCnt, 1) = "\" Then
        sSpecOut = Left$(sSpecIn, nCnt)
        Exit For
      End If
    Next nCnt
  
  ElseIf InStr(sSpecIn, ":") = 2 Then
    sSpecOut = CurDir$(sSpecIn)
    If Len(sSpecOut) = 0 Then sSpecOut = CurDir$

  Else
    sSpecOut = CurDir$
  End If
    
  If Right$(sSpecOut, 1) <> "\" Then
    sSpecOut = sSpecOut + "\"
  End If
  ExtractPath$ = UCase$(sSpecOut)

End Function

Public Sub PrivClearEntry(sEntryName As String)

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Sets a specific entry in Private.Ini to Nothing or Blank
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, sEntryName, "", smIniFileName)

End Sub

Public Sub PrivDeleteEntry(sEntryName As String)

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Deletes a specific entry in Private.Ini
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, sEntryName, 0&, smIniFileName)

End Sub

Public Sub PrivDeleteSection()

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Deletes an *entire* [Section] and all its Entries in Private.Ini
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, 0&, 0&, smIniFileName)

  'Now Private.Ini needs to be reinitialized
    smSectionName = ""
    nmPrivInit = False

End Sub

Public Function PrivGetInt(sEntryName As String, nDefaultValue As Integer) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Retrieves an Integer value from Private.Ini, range: 0-32767
    PrivGetInt = kpGetPrivateProfileInt(smSectionName, sEntryName, nDefaultValue, smIniFileName)

End Function

Public Function PrivGetString(sEntryName As String, ByVal sDefaultValue As String) As String

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Retrieves Specific Entry from Private.Ini
    Dim sTemp As String * Max_EntryBuffer
    Dim nRetVal As Integer
    nRetVal = kpGetPrivateProfileString(smSectionName, sEntryName, sDefaultValue, sTemp, Len(sTemp), smIniFileName)
    If nRetVal Then
      PrivGetString = Left$(sTemp, nRetVal)
    End If

End Function

Public Function PrivGetTF(sEntryName As String, nDefaultValue As Integer)
  
  'Retrieves Specific Entry as either True/False from Private.Ini
  'local vars
    Dim sTF As String
    Dim sDefault As String

  'get string value from INI
    If nDefaultValue Then
      sDefault = "true"
    Else
      sDefault = "false"
    End If
    sTF = PrivGetString(sEntryName, sDefault)

  'interpret return string
    Select Case Trim$(UCase$(sTF))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
        PrivGetTF = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
        PrivGetTF = False
      Case Else
        PrivGetTF = False
    End Select

End Function

Private Sub PrivIniNotReg()
  
  'Warn *PROGRAMMER* that there's a logic error!
    MsgBox "[Section] and FileName Not Registered in Private.Ini!", 16, "IniFile Logic Error"

End Sub

Public Sub PrivIniRead(SectionName$, KeyName$, nDefault%, ByVal DefaultStr$, ReturnStr$, Numeric%, IniFileName$)

  'One-shot read from Private.Ini, more *work* than it's worth
    Dim nRetVal As Integer
    Dim RetStr As String * Max_EntryBuffer 'Create an empty string to be filled

    If Numeric% Then    'we are looking for integer input
      Numeric% = kpGetPrivateProfileInt(SectionName$, KeyName$, nDefault%, IniFileName$)
    Else
      nRetVal = kpGetPrivateProfileString(SectionName$, KeyName$, DefaultStr$, RetStr$, Len(RetStr$), IniFileName$)
      If nRetVal Then
        ReturnStr$ = Left$(RetStr$, nRetVal)
      End If
    End If

End Sub

Public Sub Create(sSectionName As String, sIniFileName As String)

  'Store module-level values for future reference
    smSectionName = Trim$(sSectionName)
    smIniFileName = Trim$(sIniFileName)
    nmPrivInit = True

End Sub

Public Sub PrivIniWrite(SectionName$, IniFileName$, EntryName$, ByVal NewVal$)
    
  'One-shot write to Private.Ini, more *work* than it's worth
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(SectionName$, EntryName$, NewVal$, IniFileName$)
    
End Sub

Public Function PrivPutInt(sEntryName As String, nValue As Integer) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Write an integer to Private.Ini
    PrivPutInt = kpWritePrivateProfileString(smSectionName, sEntryName, Format$(nValue), smIniFileName)

End Function

Public Function PrivPutString(sEntryName As String, ByVal sValue As String) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Write a string to Private.Ini
    PrivPutString = kpWritePrivateProfileString(smSectionName, sEntryName, sValue, smIniFileName)

End Function

Public Function PrivPutTF(sEntryName As String, nValue As Integer)

  'Set an entry in Private.Ini to True/False
  'local vars
    Dim sTF As String

  'create INI string
    If nValue Then
      sTF = "true"
    Else
      sTF = "false"
    End If

  'write new value
    PrivPutTF = PrivPutString(sEntryName, sTF)

End Function

Public Sub ChangeINIFile(nPathFlag%)

  'Store module-level flag for future reference
    nmSysPath = nPathFlag

End Sub
