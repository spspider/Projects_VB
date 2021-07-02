VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1500
      Top             =   180
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   2700
      Top             =   1140
      _ExtentX        =   1693
      _ExtentY        =   33867
      _Version        =   327680
      Picture         =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2580
      TabIndex        =   0
      Top             =   300
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Top             =   660
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
        
    Timer1.Enabled = Not (Timer1.Enabled)

    
End Sub






Private Sub Form_Load()
    
    With PictureClip1
        .Cols = 1
        .Rows = 40
    End With
    
End Sub


Private Sub Timer1_Timer()
    
    Static Index1 As Integer, Index2 As Integer
    Image1.Visible = False
    Image2.Visible = False
    Index1 = Index1 + 1
    If Index1 = 40 Then Index1 = 0
    Image1.Picture = PictureClip1.GraphicCell(Index1)
    Index2 = Index1
    If Index1 = 39 Then Index2 = -1
    Image2.Picture = PictureClip1.GraphicCell(Index2 + 1)
    Image1.Visible = True
    Image2.Visible = True
    
End Sub


