VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBSoundBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WORKING!

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)

Public FileName As String
Public m_DS As IDirectSound
Public m_lpDSB As IDirectSoundBuffer

Private Sub Class_Terminate()
      Set m_lpDSB = Nothing
      Set m_DS = Nothing
End Sub

Public Sub LoadFromDisk(fname As String, hWnd As Long)
      LoadWav hWnd, m_DS, fname, m_lpDSB
End Sub

Private Sub LoadWav(hWnd As Long, Lds As IDirectSound, ByVal fname As String, Ldsb As IDirectSoundBuffer)
      Dim hWave As Long
      Dim pcmwave As WAVEFORMATEX
      Dim lngSize As Long
      Dim lngPosition As Long
      Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
      Dim aByte() As Byte

      ReDim aByte(1 To FileLen(fname))
      hWave = FreeFile
      Open fname For Binary As hWave
            Get hWave, , aByte
      Close hWave
      lngPosition = 1

      While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
            lngPosition = lngPosition + 1
      Wend

      CopyMemory VarPtr(pcmwave), VarPtr(aByte(lngPosition + 8)), Len(pcmwave)

      While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
            lngPosition = lngPosition + 1
      Wend

      CopyMemory VarPtr(lngSize), VarPtr(aByte(lngPosition + 4)), Len(lngSize)

      Dim dsbd As DSBUFFERDESC
      With dsbd
            .dwSize = Len(dsbd)
            .dwFlags = DSBCAPS_CTRLDEFAULT
            .dwBufferBytes = lngSize
            .lpwfxFormat = VarPtr(pcmwave)
      End With

      DirectSoundCreate ByVal 0&, Lds, Nothing
      Lds.SetCooperativeLevel hWnd, DSSCL_NORMAL

      Lds.CreateSoundBuffer dsbd, Ldsb, Nothing
      Ldsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&

      CopyMemory ptr1, VarPtr(aByte(lngPosition + 4 + 4)), lng1

      If lng2 <> 0 Then
            CopyMemory ptr2, VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
      End If
End Sub

Public Sub Play(dwFlags As Long)
      ' An example of dwFlags may be DSBPLAY_LOOPING
      m_lpDSB.Play 0, 0, dwFlags
End Sub

Public Sub StopSound()
      m_lpDSB.Stop
      m_lpDSB.SetCurrentPosition 0
End Sub

Public Sub SetVolume(VOLUME As Long)
      m_lpDSB.SetVolume VOLUME
End Sub

Public Function GetVolume() As Long
      Dim VOL As Long

      m_lpDSB.GetVolume VOL
      
      GetVolume = VOL
End Function

Public Sub SetFrequency(FREQUENCY As Long)
      m_lpDSB.SetFrequency FREQUENCY
End Sub

Public Function GetFrequency() As Long
      Dim FREQ As Long
      
      m_lpDSB.GetFrequency FREQ
      
      GetFrequency = FREQ
End Function

Public Sub SetPan(PAN As Long)
      m_lpDSB.SetPan PAN
End Sub

Public Function GetPan() As Long
      Dim PAN As Long
      
      m_lpDSB.GetPan PAN
      
      GetPan = PAN
End Function
