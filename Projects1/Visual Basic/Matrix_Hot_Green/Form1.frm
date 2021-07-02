VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HotMatrix"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox img 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   9
      Top             =   0
      Width           =   9600
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6000
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   255
      Left            =   9960
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   9720
      Max             =   50
      Min             =   1
      TabIndex        =   5
      Top             =   360
      Value           =   5
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   9720
      Max             =   500
      Min             =   2
      TabIndex        =   4
      Top             =   1080
      Value           =   100
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   9720
      Pattern         =   "*.bmp"
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Потухание"
      Height          =   255
      Left            =   9720
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   9720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   3840
      Width           =   795
   End
   Begin VB.PictureBox Pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      Height          =   1500
      Left            =   8640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "FPS:"
      Height          =   255
      Left            =   9960
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Кол-во симв."
      Height          =   255
      Left            =   9720
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dat(0 To 52, 0 To 47) As Byte
Dim kDat(0 To 52, 0 To 47) As Byte

Dim FPS

Sub RUN()
1
DoEvents
For X = 0 To 600 Step 100
For Y = 0 To 400 Step 100
BitBlt img.hdc, X, Y, 200, 200, Pc.hdc, 0, 0, vbSrcCopy
Next
Next



For X = 0 To 52
For Y = 46 To 0 Step -1
I = GetPixel(Im.hdc, X, Y) / 65536
If I < 0 Then I = 0
If I > 255 Then I = 255
If I > 32 Then
kDat(X, Y) = I
End If

I = Dat(X, Y)
Dat(X, Y + 1) = I
I = kDat(X, Y)
kDat(X, Y + 1) = I
Next
Next

For X = 0 To 52
For Y = 1 To 45
I = Dat(X, Y)
If I = 32 Then kDat(X, Y - 1) = 255
Next
Next

For X = 0 To 52
I = Int(Rnd * 256)
If Int(Rnd * HScroll1) = Int(Rnd * HScroll1) Then I = 32
Dat(X, 0) = I
For Y = 0 To 47
img.ForeColor = RGB(0, kDat(X, Y), 0)
If Check1.Value <> 1 Then
Yb = Y + 1
If Yb > 47 Then Yb = 47
SD = kDat(X, Y)
SD = kDat(X, Yb) - (HScroll2 + 10)
'Sd = Sd * 29 / 30
Else
SD = kDat(X, Y)
SD = SD - HScroll2
If SD < 0 Then SD = 0
kDat(X, Y) = SD
End If

If SD < 0 Then SD = 0
kDat(X, Y) = SD
I = Dat(X, Y)
S = Chr$(I)
TextOut img.hdc, X * 12, Y * 10, S, 1
Next
Next

img.Refresh
FPS = FPS + 1
GoTo 1

End Sub

Private Sub Command1_Click()
RUN
End Sub

Private Sub File1_Click()
Im.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
A = Int(X / 12)
B = Int(Y / 10)
If A < 0 Then A = 0
If B < 0 Then B = 0
If A > 52 Then A = 52
If B > 47 Then B = 47
kDat(A, B) = 255
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form1
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
End
End Sub

Private Sub Timer1_Timer()
For X = 0 To 600 Step 100
For Y = 0 To 400 Step 100
BitBlt Form1.hdc, X, Y, 200, 200, Pc.hdc, 0, 0, vbSrcCopy
Next
Next
For X = 0 To 52
For Y = 46 To 0 Step -1
I = GetPixel(Im.hdc, X, Y) / 65536
If I < 0 Then I = 0
If I > 255 Then I = 255
If I > 32 Then
kDat(X, Y) = I
End If
I = Dat(X, Y)
Dat(X, Y + 1) = I
I = kDat(X, Y)
kDat(X, Y + 1) = I
Next
Next
For X = 0 To 52
For Y = 1 To 45
I = Dat(X, Y)
If I = 32 Then kDat(X, Y - 1) = 255
Next
Next
For X = 0 To 52
I = Int(Rnd * 256)
If Int(Rnd * HScroll1) = Int(Rnd * HScroll1) Then I = 32
Dat(X, 0) = I
For Y = 0 To 47
Form1.ForeColor = RGB(0, kDat(X, Y), 0)
If Check1.Value <> 1 Then
Yb = Y + 1
If Yb > 47 Then Yb = 47
SD = kDat(X, Y)
SD = kDat(X, Yb) - (HScroll2 + 10)
'Sd = Sd * 29 / 30
Else
SD = kDat(X, Y)
SD = SD - HScroll2
If SD < 0 Then SD = 0
kDat(X, Y) = SD
End If

If SD < 0 Then SD = 0
kDat(X, Y) = SD
I = Dat(X, Y)
S = Chr$(I)
TextOut Form1.hdc, X * 12, Y * 10, S, 1
Next
Next

Form1.Refresh
End Sub


Private Sub Timer2_Timer()
Label2.Caption = "FPS: " & FPS
FPS = 0
End Sub
