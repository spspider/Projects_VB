VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   15  'Merge Pen Not
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ACV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3480
      ScaleHeight     =   120
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   1320
      Width           =   450
   End
   Begin VB.PictureBox EXT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3960
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      Top             =   45
      Width           =   120
   End
   Begin VB.PictureBox BUNS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   1680
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   840
      Width           =   330
   End
   Begin VB.PictureBox BUNS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   1320
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   5
      Top             =   840
      Width           =   330
   End
   Begin VB.PictureBox BUNS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   960
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   840
      Width           =   330
   End
   Begin VB.PictureBox BUNS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   600
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   840
      Width           =   330
   End
   Begin VB.PictureBox BUNS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   240
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      Top             =   840
      Width           =   330
   End
   Begin VB.PictureBox Scroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   240
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1320
      Width           =   3000
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   1
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BTN(0 To 4)
Dim XtB
Dim SNDScrll, MaxSS
Dim PLState

Private Sub BUNS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If BTN(N) <> 2 Then
For N = 0 To 4
BTN(N) = 0
Next
BTN(Index) = 1
End If
End Sub

Private Sub BUNS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If BTN(Index) <> 2 Then
For N = 0 To 4
BTN(N) = 0
Next
End If

If X >= 0 And Y >= 0 And X <= 22 And Y <= 18 And BTN(Index) <> 2 Then

If Index = 1 Then
MClose
OpenFirst Label2.Caption & Label1.Caption
PlayFirst
MaxSS = GetFirstLen
PLState = 1
End If


If Index = 2 Then
StopFirst
PLState = 2
End If

If Index = 3 Then
MClose
PLState = 0
End If

If Index = 4 Then NextSing
If Index = 0 Then BackSing
End If
End Sub

Private Sub EXT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XtB = 1
End Sub



Private Sub EXT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And Y >= 0 And X <= 8 And Y <= 8 Then
Unload Form1
Unload Form2
Unload Surce
End
Else
XtB = 0
End If
End Sub

Private Sub Form_Activate()
Surce.Label1.Caption = 0
End Sub


Private Sub Form_Load()
Surce.BGD.Picture = LoadPicture(App.Path & "\Skins\Main.bmp")
Surce.Buttons.Picture = LoadPicture(App.Path & "\Skins\Buttons.bmp")
Surce.ACV.Picture = LoadPicture(App.Path & "\Skins\Active.bmp")
Surce.PLST.Picture = LoadPicture(App.Path & "\Skins\Title_P.bmp")
Surce.PP.Picture = LoadPicture(App.Path & "\Skins\T_M.bmp")
Surce.PLST2.Picture = LoadPicture(App.Path & "\Skins\T_P.bmp")
Surce.Xt.Picture = LoadPicture(App.Path & "\Skins\Ext.bmp")
Surce.PLBTN.Picture = LoadPicture(App.Path & "\Skins\SV_PL.bmp")
Scroll.Picture = LoadPicture(App.Path & "\Skins\Scroll.bmp")
Pic.Picture = LoadPicture(App.Path & "\Skins\Scroll_Box.BMP")
Form2.LeftB.Picture = LoadPicture(App.Path & "\Skins\LeftT_P.bmp")
Form2.RightB.Picture = LoadPicture(App.Path & "\Skins\RightT_P.bmp")

Form1.Left = Screen.Width / 2 - Form1.Width / 2
Form1.Top = Screen.Height / 2 - Form1.Height / 2
W = Surce.BGD.Width
H = Surce.BGD.Height
Form1.Width = W * 15
Form1.Height = H * 15
Form2.Width = W * 15
SNDScrll = 0
MaxSS = 100
Form2.Visible = True
XtB = 0
PLState = 0
Form2.Left = Form1.Left
Form2.Top = Form1.Height + Form1.Top
End Sub





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mXp = X
mYp = Y
S = 0
If Button = 1 Then
a = Form2.Left
B = Form2.Top
nx = Form1.Left + Form1.Width
nx2 = Form1.Left
ny = Form1.Top + Form1.Height
ny2 = Form1.Top
If RZN(a / 15, nx / 15) < 1 Or RZN(nx2 / 15, (a + Form2.Width) / 15) < 2 Or RZN(B / 15, ny / 15) < 2 Or RZN(ny2 / 15, (B + Form2.Height) / 15) < 2 Then S = 1

Form1.Left = Form1.Left + X - Form1.CurrentX
Form1.Top = Form1.Top + Y - Form1.CurrentY

If S = 1 Then
Form2.Left = Form2.Left + X - Form1.CurrentX
Form2.Top = Form2.Top + Y - Form1.CurrentY
End If



End If
End Sub

Private Sub Label1_Click()
MClose
OpenFirst Label2.Caption & Label1.Caption
PlayFirst
MaxSS = GetFirstLen
PLState = 1

End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic.CurrentX = X
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = Pic.Left + X - Pic.CurrentX
If a < 0 Then a = 0
If a > 190 Then a = 190
Pic.Left = a
N = Int(GetSS)
PlayFrom N
End If
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
N = Int(GetSS)
PlayFrom N
End Sub

Private Sub Scroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X - 5
If a < 0 Then a = 0
If a > 190 Then a = 190
Pic.Left = a
N = Int(GetSS)
PlayFrom N
End If
End Sub

Private Sub Scroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X - 5
If a < 0 Then a = 0
If a > 190 Then a = 190
Pic.Left = a
N = Int(GetSS)
PlayFrom N
End If
End Sub

Private Sub Timer1_Timer()
If Form2.List1.ListCount = 0 Then
For N = 0 To 4
BTN(N) = 2
Next
Else

For N = 0 To 4
If BTN(N) = 2 Then BTN(N) = 0
Next

End If

BitBlt Form1.hdc, 0, 0, Surce.BGD.Width, Surce.BGD.Height, Surce.BGD.hdc, 0, 0, vbSrcCopy

If PLState <> 0 Then
PZNT = Int(Pic.Left / 1.9 * 10) / 10
txt$ = PZNT & ":100"
TextOut Form1.hdc, 160, 55, txt$, Len(txt$)
End If

Form1.Refresh
For N = 0 To 4
BitBlt BUNS(N).hdc, 0, 0, 22, 18, Surce.Buttons.hdc, N * 22, BTN(N) * 18, vbSrcCopy
BUNS(N).Refresh
Next
BitBlt EXT.hdc, 0, 0, 8, 8, Surce.Xt.hdc, 0, XtB * 8, vbSrcCopy
EXT.Refresh
BitBlt ACV.hdc, 0, 0, 30, 8, Surce.ACV.hdc, 0, Surce.Label1 * 8, vbSrcCopy
ACV.Refresh
If PLState <> 0 Then SetSS GetFirstPos
If GetFirstPos >= MaxSS Then NextSing
End Sub

Sub NextSing()
L = Form2.List1.ListIndex
Mx = Form2.List1.ListCount
If Mx > 0 Then
If Mx = 1 Then Form2.List1.ListIndex = 0
If L = Mx - 1 Then Form2.List1.ListIndex = 0
If Mx > 1 And L < Mx - 1 Then Form2.List1.ListIndex = Form2.List1.ListIndex + 1
MClose
OpenFirst Label2.Caption & Label1.Caption
PlayFirst
MaxSS = GetFirstLen
PLState = 1
SetSS 0
End If
End Sub

Sub BackSing()
L = Form2.List1.ListIndex
Mx = Form2.List1.ListCount
If Mx > 0 Then
If Mx = 1 Then Form2.List1.ListIndex = 0
If L = 0 Then Form2.List1.ListIndex = Mx - 1
If Mx > 1 And L > 0 Then Form2.List1.ListIndex = Form2.List1.ListIndex - 1
MClose
OpenFirst Label2.Caption & Label1.Caption
PlayFirst
MaxSS = GetFirstLen
PLState = 1
SetSS 0
End If

End Sub

Sub SetSS(State)
N = State
If N < 0 Then N = 0
If N > MaxSS Then N = MaxSS
If MaxSS < 1 Then MaxSS = 1
K = 190 / MaxSS
Pic.Left = N * K
End Sub

Function GetSS()
N = Pic.Left
If N < 0 Then N = 0
If N > MaxSS Then N = MaxSS
If MaxSS < 1 Then MaxSS = 1
K = 190 / MaxSS
GetSS = N / K
End Function
