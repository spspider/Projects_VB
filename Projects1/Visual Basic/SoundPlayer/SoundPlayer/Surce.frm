VERSION 5.00
Begin VB.Form Surce 
   Caption         =   "Form3"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "Surce.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PLBTN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox ACV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1080
      ScaleHeight     =   240
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   2760
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox Xt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2400
      ScaleHeight     =   240
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   120
   End
   Begin VB.PictureBox PLST2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   2160
      ScaleHeight     =   765
      ScaleWidth      =   90
      TabIndex        =   4
      Top             =   1920
      Width           =   90
   End
   Begin VB.PictureBox PLST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   2040
      Width           =   900
   End
   Begin VB.PictureBox PP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1920
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   2
      Top             =   1800
      Width           =   105
   End
   Begin VB.PictureBox Buttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   120
      ScaleHeight     =   810
      ScaleWidth      =   1650
      TabIndex        =   1
      Top             =   1800
      Width           =   1650
   End
   Begin VB.PictureBox BGD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   0
      Width           =   4125
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   495
   End
End
Attribute VB_Name = "Surce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
