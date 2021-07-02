VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDXVBCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const Yspeed = 1 ' Move the credits at the speed of 1 pixel each loop
Const ConstColors = 50 'how many shades of gray do you want to fade in & out
Private C(ConstColors) As Long ' Array of colours (For fading)
Public PosY As Long ' Y starting text position

Private Sub Class_Initialize()
  For i = 0 To ConstColors
    ' Greyscale
    'C(i) = RGB(i * Int(256 / ConstColors), i * Int(256 / ConstColors), i * Int(256 / ConstColors))
    
    'green
    'C(i) = RGB(0, i * Int(256 / ConstColors), 0)
    
    'red
    C(i) = RGB(i * Int(256 / ConstColors), 0, 0)
    
    'blue
    'C(i) = RGB(0, 0, i * Int(256 / ConstColors))
  Next i
End Sub

Public Sub Draw(hDC As Long, SWidth As Long, SHeight As Long)
  Dim n As Long
  If PosY < -(14 * (UBound(Strings) + 1)) Then PosY = SHeight
  For i = 0 To UBound(Strings)
    If (PosY + (i * 14)) < Int(SHeight / 2) Then
      n = Int(SHeight / 2) - (Int(SHeight / 2) - (PosY + (i * 14)))
      If n >= 0 And n < ConstColors Then
        SetTextColor hDC, C(n)
      Else
        If n < 0 Then
          SetTextColor hDC, C(0)
        Else
          SetTextColor hDC, C(ConstColors)
        End If
      End If
    Else
      n = Int(SHeight - (PosY + (i * 14)))
      If n > 0 And n < ConstColors Then
        SetTextColor hDC, C(n)
      Else
        SetTextColor hDC, C(ConstColors)
      End If
    End If
    TextOut hDC, frmMain.ScaleWidth / 2, (i * 14) + PosY, Strings(i), Len(Strings(i))
  Next i
  PosY = PosY - Yspeed
End Sub
