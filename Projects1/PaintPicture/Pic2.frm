VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "SpecialEdition"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   Icon            =   "Pic2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   8760
      Pattern         =   "*.exe"
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   960
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   8040
      Picture         =   "Pic2.frx":5A4A
      ScaleHeight     =   1515
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   8040
      Picture         =   "Pic2.frx":14746
      ScaleHeight     =   1515
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8040
      Picture         =   "Pic2.frx":23442
      ScaleHeight     =   1575
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   0
      Picture         =   "Pic2.frx":3213E
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub File1_DblClick()
Shell ("Install\" + File1.FileName), vbNormalFocus
'Shell (Left(App.Path, 3) + "ZOthers\" + File1.FileName), vbNormalFocus
End Sub
Private Sub Form_Load()
On Error GoTo er1:
File1.Left = Width
'File1.FileName = Left(App.Path, 3) + "ZOthers"
File1.FileName = "Install\"
er1:
On Error GoTo er:
Call mciExecute("play data\menu.mp3")
'Shell ("data\install.exe")
er:
naz = True
End Sub

Private Sub Picture1_Click()
If naz = False Then
Timer2.Enabled = True
End If
If naz = True Then
Timer3.Enabled = True
End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Picture2.Picture
End Sub

Private Sub Timer1_Timer()
Form1.Cls
a2 = a2 + 1
For q = 0 To 40
X = Tan(a2 / 20 + q / 8) * 3000 + Width / 2
Y = Cos(a2 / 60 + q / 4) * 3000 + Height / 2 * q / 20
PaintPicture Image1, X, Y
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
v1 = v1 - 1
For a1 = 1 To 40
op1T = Sin(((-v1 / 20) - a1 * 400)) * Cos(v1 / 60 - a1 / 100) * 1600 + Height - 2100
op1L = Cos(((-v1 / 20) - a1 * 400)) * Cos(v1 / 60 - a1 / 100) * 1600 + 2100
Line (op1L, op1T)-(op2L, op2T), RGB(255, 255 - a1 * 3, 255), B
'Circle (op1L, op1T), RGB(15, 5, 0), RGB(255, 255 - a1 * 10, 255)
op2T = Height - 2100 '- op1T3
op2L = 2100 '- op1L3
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
Picture1.Picture = Picture3.Picture
End Sub
Private Sub Form_Resize()
Picture1.Left = Width - 3200
Picture1.Top = Height - 2200
If naz = False Then
Timer2.Enabled = True
End If
If naz = True Then
Timer3.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
e = e + 1
File1.Left = Width - Sin(e / 30) * 2600
If File1.Left < Width - (File1.Width + 200) Then
Timer2.Enabled = False
naz = True
e = 0
End If
End Sub
Private Sub Timer3_Timer()
e = e + 1
File1.Left = Width - Cos(e / 30) * 2600
If File1.Left > Width + 500 Then
Timer3.Enabled = False
naz = False
e = 0
End If
End Sub

