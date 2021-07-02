VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ScaleHeight     =   975
   ScaleWidth      =   15360
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   480
      Picture         =   "Pic2.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Xm As Long
Dim Ym As Long
Dim X As Long
Dim Y As Long
Dim a2 As Long
Dim q As Long

Dim Xver As Long
Dim Yver As Long

Dim naz As Boolean
Dim e As Long
Dim v1 As Long
Dim v2 As Long

Dim op1L As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long

Dim op1L1 As Long
Dim op1T1 As Long
Dim op2L1 As Long
Dim op2T1 As Long
Private Sub Timer1_Timer()
UserControl.Cls
a2 = a2 + 1
For q = 0 To 60
X = (Tan(-a2 / 100 + q / 10) * Width / 2 + Width / 2)
Y = (Sin(-a2 / 12 + q / 4) * Height / 3 + Height / 2 - 100)
PaintPicture Image1, X, Y
Next
If X > Width Then
X = 0
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'v1 = v1 - 1
'For a1 = 1 To 30
'op1T = Sin(((-v1 / 20) - a1 * 400)) * Cos(v1 / 60 - a1 / 100) * 600 + Height - 2100
'op1L = Cos(((-v1 / 20) - a1 * 400)) * Cos(v1 / 60 - a1 / 100) * 600 + 2100
'Line (op1L, op1T)-(op2L, op2T), RGB(255, 255 - a1 * 3, 255), B
'op2T = Height - 2100
'op2L = 2100
'Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


