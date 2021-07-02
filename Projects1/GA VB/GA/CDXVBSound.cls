VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WORKING!

Public m_lpDS As IDirectSound
Private m_DSCaps As DSCAPS

Private Sub Class_Terminate()
      Set m_lpDS = Nothing
End Sub

Public Function Create(hWnd As Long) As Boolean
      Dim result As Integer

      DirectSoundCreate ByVal 0&, m_lpDS, Nothing

      m_lpDS.SetCooperativeLevel hWnd, DSSCL_NORMAL

      Create = True
End Function

Public Function GetCaps() As Boolean
      Dim result As Long

      m_lpDS.GetCaps m_DSCaps

      GetCaps = True
End Function
