VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "AutoRoad"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   5520
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   1
      Top             =   5460
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.PictureBox picRoad 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6915
      Left            =   120
      ScaleHeight     =   461
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   331
      TabIndex        =   0
      Top             =   240
      Width           =   4965
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DrawRoad()

    picRoad.Line (0, 0)-(picRoad.Width * 0.05, picRoad.Height), vbGreen, BF

    picRoad.Line (picRoad.Width * 0.95, 0)-(picRoad.Width, picRoad.Height), _
    vbGreen, BF

    picRoad.Line (picRoad.Width * 0.05, 0)-(picRoad.Width * 0.95, _
    picRoad.Height), vbBlack, BF

    picRoad.Line (picRoad.Width * 0.5, 0)-(picRoad.Width * 0.5, _
    picRoad.Height), vbYellow, BF

End Sub
Public Type POINTAPI

    x As Long

    y As Long

End Type


Public Auto As POINTAPI
Public Sub DrawAuto()

    StretchBlt picRoad.hdc, Auto.x, Auto.y, _
        23, 31, picSrc.hdc, 0, 0, 23, 31, SRCCOPY

End Sub

