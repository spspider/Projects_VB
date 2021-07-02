VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VS3 
      Height          =   4455
      Left            =   5760
      Max             =   255
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.VScrollBar VS1 
      Height          =   4455
      Left            =   5400
      Max             =   255
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.VScrollBar VS2 
      Height          =   4455
      Left            =   5040
      Max             =   255
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lY As Long
Dim lScaleHeight As Long
Dim lScaleWidth As Long
Private Sub Form_load()

AutoRedraw = False

End Sub

Private Sub Form_Paint()
ScaleMode = vbPixels
lScaleHeight = ScaleHeight
lScaleWidth = ScaleWidth

DrawStyle = vbInvisible
FillStyle = vbFSSolid
End Sub

Private Sub Timer1_Timer()
For lY = 0 To lScaleHeight

FillColor = RGB(255 - VS1.Value + lY / 4, 255 - lY / 4 + VS2.Value, 255 - lY / 4 + VS3.Value)
Line (-1, lY - 1)-(lScaleWidth, lY + 1), , B
Next lY

VS1.Left = ScaleWidth - 100
VS2.Left = ScaleWidth - 126
VS1.Top = ScaleHeight - 300
VS2.Top = ScaleHeight - 300
VS3.Left = ScaleWidth - 152
VS3.Top = ScaleHeight - 300
End Sub
