VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome To SP"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8895
   Icon            =   "SP_3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSpr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   4800
   End
   Begin VB.Timer TimerOPE 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   4200
   End
   Begin VB.Timer TimerOP 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   4560
   End
   Begin VB.Timer TimerOpt 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   4200
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Установить...>"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7830
      Width           =   1335
   End
   Begin VB.Timer Timer14 
      Interval        =   1000
      Left            =   720
      Top             =   5760
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   240
      Top             =   5760
   End
   Begin VB.OptionButton Op4 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op3 
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op2 
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op1 
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00800000&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   5400
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   2040
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   1680
   End
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   120
      Top             =   1680
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   960
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00808000&
      Height          =   1845
      Left            =   5640
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Привет всем от Серёги!!!!"
      BeginProperty Font 
         Name            =   "Courier New CYR"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   945
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   3600
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2280
      X2              =   3480
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   3480
      X2              =   1440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2280
      X2              =   1800
      Y1              =   960
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1800
      X2              =   3480
      Y1              =   2520
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1800
      X2              =   1440
      Y1              =   2520
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2280
      X2              =   1440
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      X1              =   7440
      X2              =   7440
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000A&
      X1              =   7560
      X2              =   7560
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line9 
      X1              =   7680
      X2              =   7680
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line10 
      X1              =   7800
      X2              =   7800
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line11 
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   8040
      X2              =   8040
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line18 
      X1              =   8160
      X2              =   8160
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line13 
      X1              =   8280
      X2              =   8280
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line14 
      X1              =   8400
      X2              =   8400
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line15 
      X1              =   8520
      X2              =   8520
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line16 
      X1              =   8640
      X2              =   8640
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line17 
      X1              =   8760
      X2              =   8760
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Menu file 
      Caption         =   "Файл"
      Begin VB.Menu music 
         Caption         =   "*    Музыка"
      End
      Begin VB.Menu sled 
         Caption         =   "      Выжигание"
      End
      Begin VB.Menu opt 
         Caption         =   "Опции"
      End
      Begin VB.Menu exit 
         Caption         =   "выход"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Справка"
      Begin VB.Menu prog 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Long
Dim Form1LE As Long
Dim Spr As Long
Dim Form3l As Currency
Dim FormL As Long
Dim ScreenH As Integer
Dim ScreenW As Integer
Dim Xver As Long
Dim Yver As Long
Dim c2 As Long
Dim e As Long
Dim b As Long
Dim le1 As Integer
Dim sledval As Integer
Dim musval As Integer
Dim mus As String
Dim le2 As Integer
Dim le3 As Integer
Dim le4 As Integer
Dim com1 As Integer
Dim fi As Integer
Dim li1 As Integer
Dim fi2 As Integer
Dim fi3 As Integer
Dim h1 As Long
Dim h2 As Long
Dim w As Integer
Dim s As Integer
Dim s1 As Integer
Dim i As Long
Dim a As Integer
Dim k As Currency
Dim a1 As Long
Dim er As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub Command1_Click()
If com1 = 0 Then
Timer11.Enabled = True
Timer12.Enabled = False
Command1.Caption = "<"
com1 = 1
Else
Command1.Caption = ">"
Timer12.Enabled = True
Timer11.Enabled = False
com1 = 0
End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HC00000
File1.BackColor = &H800000
End Sub

Private Sub Command2_Click()
On Error GoTo er:
Shell "VB\SETUP.exe", vbNormalFocus
er:
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &H80FF80
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub File1_DblClick()
musval = 0
music.Caption = "* Музыка"
Call mciExecute("close all")
li1 = File1.ListIndex
mus = "play " + "Data\music\" & File1.FileName
Call mciExecute(mus)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''
Private Sub file1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
musval = 0
music.Caption = "* Музыка"
Call mciExecute("close all")
li1 = File1.ListIndex
mus = "play " + "Data\music\" & File1.FileName
Call mciExecute(mus)
End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H800000
File1.BackColor = &HC00000
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H800000
Command2.BackColor = &HFF00&
Label1.BorderStyle = 0
File1.BackColor = &H800000
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 1
End Sub

Private Sub Form_Click()
a1 = 5000
End Sub
Private Sub Form_Load() ''''''''''''музыка'''''




Form2.Top = -Form2.Width
ScreenW = Screen.Width / 15
ScreenH = Screen.Height / 15

Op1.Top = -100
Op1.Left = -100
Op2.Top = -100
Op2.Left = -100
Op3.Top = -100
Op3.Left = -100
Op4.Top = -100
Op4.Left = -100

Command2.Top = 0 - Command2.Height

Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Line13.Visible = False
Line14.Visible = False
Line15.Visible = False
Line16.Visible = False
Line17.Visible = False
Line18.Visible = False

Line7.BorderColor = RGB(180, 180, 255)
Line8.BorderColor = RGB(180, 180, 255)
Line9.BorderColor = RGB(180, 180, 255)
Line10.BorderColor = RGB(180, 180, 255)
Line11.BorderColor = RGB(180, 180, 255)
Line12.BorderColor = RGB(180, 180, 255)
Line13.BorderColor = RGB(180, 180, 255)
Line14.BorderColor = RGB(180, 180, 255)
Line15.BorderColor = RGB(180, 180, 255)
Line16.BorderColor = RGB(180, 180, 255)
Line17.BorderColor = RGB(180, 180, 255)
Line18.BorderColor = RGB(180, 180, 255)

li1 = 0 '''''''li - номер файла в ListIndex-е
Command1.Top = 6400
Command1.Left = -500
Timer9.Enabled = True
File1.Top = Height + 10
w = 0
On Error GoTo er1:
File1.FileName = "Data\music\"
File1.ListIndex = 0

Call mciExecute("close all")
li1 = File1.ListIndex
mus = "play " + "Data\music\" & File1.FileName
Call mciExecute(mus)

er1:
End Sub

Private Sub music_Click()
If musval = 0 Then
paus = "pause " + "Data\music\" & File1.FileName
Call mciExecute(paus)
music.Caption = "      Музыка"
musval = 1
Else
pla = "play " + "Data\music\" & File1.FileName
Call mciExecute(pla)
music.Caption = "*    Музыка"
musval = 0
End If
End Sub

Private Sub opt_Click()
Form3.Show
TimerOpt.Enabled = True
TimerOP.Enabled = True
End Sub

Private Sub prog_Click()
On Error GoTo er:
TimerSpr.Enabled = True
Form2.Show

musval = 0
music.Caption = "* Музыка"
Call mciExecute("close all")
mus = "play " + "Data\music\KVvolUP.wav"
Call mciExecute(mus)
er:
End Sub

Private Sub sled_Click()
If sledval = 0 Then
Timer16.Enabled = True
sled.Caption = "*    Выжигание"
sledval = 1
Else
Timer16.Enabled = False
sled.Caption = "     Выжигание"
sledval = 0
End If
End Sub

Private Sub Timer1_Timer()
If a > 32000 Then
a = 0
End If

e = Height - Math.Tan(a / 200) * (Height / 3) + (Height / 3)
b = Math.Sin(a / 30) * (Height / 3) + (Height / 3)

If h1 + 30 > Width Then
Timer2.Enabled = True

a1 = a1 + 1
End If
If a1 > 5000 Then
a1 = 1
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Line13.Visible = False
Line14.Visible = False
Line15.Visible = False
Line16.Visible = False
Line17.Visible = False
Line18.Visible = False
Timer2.Enabled = False
k = Math.Rnd * 100
h = 0
h1 = 0
End If

If k > 49 Then
'''''''''''''''''''''''''''эффект шторки'''''
Timer3.Enabled = True
Timer4.Enabled = False
Else
'''''''''''''''''''''''''''эффект крест''''''
Timer4.Enabled = True
Timer3.Enabled = False
End If

If w <= 180 Then
Timer5.Enabled = True
Timer6.Enabled = False
End If
If w >= 270 Then
Timer6.Enabled = True
Timer5.Enabled = False
End If

End Sub
Private Sub Timer10_Timer()
fi2 = fi2 + 1
Command1.Left = Math.Cos(fi2 / 70) * (-4800) + 1000
If Command1.Left > 5390 Then
Timer10.Enabled = False
le2 = Command1.Left
End If
End Sub

Private Sub Timer11_Timer()
le1 = le1 + 1
Command1.Left = Math.Sin(le1 / 25) * 3300 + le2
File1.Left = Command1.Left + Command1.Width
If Command1.Left > 8640 Then
Timer11.Enabled = False
le4 = Command1.Left
le1 = 0
End If
End Sub

Private Sub Timer12_Timer()
le3 = le3 + 1
Command1.Left = le4 - Math.Sin(le3 / 25) * 3300
File1.Left = Command1.Left + Command1.Width
If Command1.Left < 5400 Then
Timer12.Enabled = False
le3 = 0
End If
End Sub

Private Sub Timer13_Timer()
Label1.Visible = False
Timer13.Enabled = False
End Sub

Private Sub Timer14_Timer()
Label1.Visible = True
Timer13.Enabled = True
End Sub

Private Sub Timer15_Timer()
c2 = c2 + 1
Command2.Left = Math.Tan(c2 / 100) * (Height / 2) - 4000
Command2.Top = Math.Sin(c2 / 99) * Height
If Command2.Top > 7830 Then
Timer15.Enabled = False
End If
End Sub

Private Sub Timer16_Timer()
sledE
End Sub

Private Sub Timer2_Timer()
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
Line9.Visible = True
Line10.Visible = True
Line11.Visible = True
Line12.Visible = True
Line13.Visible = True
Line14.Visible = True
Line15.Visible = True
Line16.Visible = True
Line17.Visible = True
Line18.Visible = True

er = a1
If er > 1000 Then
er = 1000
End If
On Error GoTo er:
Xver = Math.Tan(a / Form3.HST1.Value) * (Math.Cos(a / Form3.HSC.Value)) * Form3.HSRR.Value + 2500
Yver = Math.Tan(a / Form3.HST2.Value) * (Math.Sin(a / Form3.HSS.Value)) * Form3.HSRR.Value + 2500

Op1.Top = Math.Sin(a1 / (Form3.HSP.Value + 4)) * (Form3.HSR.Value + er + Math.Rnd * 10) + Xver
Op1.Left = Math.Cos(a1 / (Form3.HSP.Value + 1)) * (Form3.HSR.Value + er + Math.Rnd * 10) + Yver

Op2.Top = Math.Sin(a1 / (Form3.HSP.Value + 3)) * (Form3.HSR.Value + er + Math.Rnd * 10) + (Xver + 1000)
Op2.Left = Math.Cos(a1 / (Form3.HSP.Value + 6)) * (Form3.HSR.Value + er + Math.Rnd * 10) + (Yver + 1000)

Op3.Top = Math.Sin(a1 / (Form3.HSP.Value + 8)) * (Form3.HSR.Value + er + Math.Rnd * 10) + (Xver + 2000)
Op3.Left = Math.Cos(a1 / (Form3.HSP.Value + 2)) * (Form3.HSR.Value + er + Math.Rnd * 10) + (Yver + 2000)

Op4.Top = Math.Sin(a1 / (Form3.HSP.Value + 3)) * (Form3.HSR.Value + Math.Rnd * 10) + (Xver + 2800)
Op4.Left = Math.Cos(a1 / Form3.HSP.Value) * (Form3.HSR.Value + Math.Rnd * 10) + (Yver + 2700)

er:
'''''''''''рисование линии и её изменение цвета в цикле(255) на 1 миллисекунду''''
i = 255



Line1.X1 = Op1.Left
Line1.Y1 = Op1.Top
Line1.X2 = Op3.Left
Line1.Y2 = Op3.Top

Line2.X1 = Op1.Left
Line2.Y1 = Op1.Top
Line2.X2 = Op2.Left
Line2.Y2 = Op2.Top

Line3.X1 = Op1.Left
Line3.Y1 = Op1.Top
Line3.X2 = Op4.Left
Line3.Y2 = Op4.Top

Line4.X1 = Op2.Left
Line4.Y1 = Op2.Top
Line4.X2 = Op3.Left
Line4.Y2 = Op3.Top

Line5.X1 = Op2.Left
Line5.Y1 = Op2.Top
Line5.X2 = Op4.Left
Line5.Y2 = Op4.Top

Line6.X1 = Op3.Left
Line6.Y1 = Op3.Top
Line6.X2 = Op4.Left
Line6.Y2 = Op4.Top

Line1.BorderColor = RGB(w - 40, 180, 255)
Line2.BorderColor = RGB(w - 40, 180, 255)
Line3.BorderColor = RGB(w - 40, 180, 255)
Line4.BorderColor = RGB(w - 40, 180, 255)
Line5.BorderColor = RGB(w - 40, 180, 255)
Line6.BorderColor = RGB(w - 40, 180, 255)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
a = a + 1

Line7.X1 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line7.Y1 = Math.Cos(a / 10) * 200 + (e + 1600)
Line7.X2 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line7.Y2 = Math.Cos(a / 10) * 100 + (e + 1600)

Line8.X1 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line8.Y1 = Math.Cos(a / 10) * 200 + (e + 1000)
Line8.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line8.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)

Line9.X1 = Math.Sin(a / 10) * 1500 + (b + 400)
Line9.Y1 = Math.Cos(a / 10) * 100 + e
Line9.X2 = Math.Sin(a / 10) * 1500 + (b + 1900)
Line9.Y2 = Math.Cos(a / 10) * 100 + e

Line10.X1 = Math.Sin(a / 10) * 1500 + b
Line10.Y1 = Math.Cos(a / 10) * 100 + (e + 600)
Line10.X2 = Math.Sin(a / 10) * 1500 + (b + 1500)
Line10.Y2 = Math.Cos(a / 10) * 100 + (e + 600)

Line11.X1 = Math.Sin(a / 10) * 1500 + b
Line11.Y1 = Math.Cos(a / 10) * 100 + (e + 600)
Line11.X2 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line11.Y2 = Math.Cos(a / 10) * 200 + (e + 1600)

Line12.X1 = Math.Sin(a / 10) * 1500 + (b + 400)
Line12.Y1 = Math.Cos(a / 10) * 100 + e
Line12.X2 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line12.Y2 = Math.Cos(a / 10) * 200 + (e + 1000)

Line13.X1 = Math.Sin(a / 10) * 1500 + (b + 1900)
Line13.Y1 = Math.Cos(a / 10) * 100 + e
Line13.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line13.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)

Line14.X1 = Math.Sin(a / 10) * 1500 + (b + 1500)
Line14.Y1 = Math.Cos(a / 10) * 100 + (e + 600)
Line14.X2 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line14.Y2 = Math.Cos(a / 10) * 100 + (e + 1600)

Line15.X1 = Math.Sin(a / 10) * 1500 + b
Line15.Y1 = Math.Cos(a / 10) * 100 + (e + 600)
Line15.X2 = Math.Sin(a / 10) * 1500 + (b + 400)
Line15.Y2 = Math.Cos(a / 10) * 100 + e

Line16.X1 = Math.Sin(a / 10) * 1500 + (b + 1500)
Line16.Y1 = Math.Cos(a / 10) * 100 + (e + 600)
Line16.X2 = Math.Sin(a / 10) * 1500 + (b + 1900)
Line16.Y2 = Math.Cos(a / 10) * 100 + e

Line17.X1 = Math.Sin(a / 10) * 1000 + (b + 1000)
Line17.Y1 = Math.Cos(a / 10) * 200 + (e + 1600)
Line17.X2 = Math.Sin(a / 10) * 1000 + (b + 1400)
Line17.Y2 = Math.Cos(a / 10) * 200 + (e + 1000)

Line18.X1 = Math.Sin(a / 10) * 1000 + (b + 2500)
Line18.Y1 = Math.Cos(a / 10) * 100 + (e + 1600)
Line18.X2 = Math.Sin(a / 10) * 1000 + (b + 2900)
Line18.Y2 = Math.Cos(a / 10) * 100 + (e + 1000)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
Private Sub Timer3_Timer() '''''''''''''шторки''''''
If h < Height Then
h = h + 15
Line (h1 - h, 0)-(Width, h), RGB(0, 0, h / 70)
End If
If h1 < Width Then
h1 = h1 + 15
Line (0, h - h1)-(h1, Height), RGB(0, 0, h1 / 70)
End If
End Sub
Private Sub Timer4_Timer() ''''''''''''крест'''''''''''''''''''
If h < Height Then
h = h + 15
Line (0, h)-(Width, h), RGB(0, 0, h / 70)
End If
If h1 < Width Then
h1 = h1 + 15
Line (h1, 0)-(h1, Height), RGB(0, 0, h / 70)
End If
End Sub

Function sledE()
Line (Op1.Left, Op1.Top)-(Op3.Left, Op3.Top), RGB(0, 255 - i, 355 - i)
Line (Op1.Left, Op1.Top)-(Op2.Left, Op2.Top), RGB(0, 255 - i, 355 - i)
Line (Op1.Left, Op1.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
Line (Op2.Left, Op2.Top)-(Op3.Left, Op3.Top), RGB(0, 255 - i, 355 - i)
Line (Op2.Left, Op2.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
Line (Op3.Left, Op3.Top)-(Op4.Left, Op4.Top), RGB(0, 255 - i, 355 - i)
End Function

Private Sub Timer5_Timer()
w = w + 1
End Sub
Private Sub Timer6_Timer()
w = w - 1
End Sub
Private Sub Timer7_Timer()
fi = fi + 1
File1.Top = Math.Cos(fi / 40) * 3000 + 8000
If File1.Top = 6400 Then
Timer7.Enabled = False
End If
End Sub
Private Sub Timer9_Timer()
Timer7.Enabled = True
Timer10.Enabled = True
Timer15.Enabled = True
Timer9.Enabled = False
End Sub

Private Sub TimerOP_Timer()
Form3l = 1 + Form3l
Form3.Left = Math.Cos(Form3l / 30) * Screen.Width / 2 + Screen.Width / 2 + 10
If Form3.Left < ((Screen.Width - Form3.Width) + 100) Then
TimerOP.Enabled = False
Form3l = 0
End If
End Sub

Private Sub TimerOPE_Timer()
Form1LE = 1 + Form1LE
Form1.Left = Math.Sin(Form1LE / 40) * (((Screen.Width / 2) - (Form1.Height / 2)) + 100)
If Form1.Left > ((Screen.Width / 2) - (Form1.Height / 2)) Then
Form1LE = 0
TimerOPE.Enabled = False
End If
End Sub

Private Sub TimerOpt_Timer()
Form1.Left = Math.Cos(FormL / 40) * ((Screen.Width / 2) - (Form1.Height / 2))
FormL = 1 + FormL
If Form1.Left < 80 Then
FormL = 0
TimerOpt.Enabled = False
End If
End Sub

Private Sub TimerSpr_Timer()
Form2.Top = Math.Cos(Spr / 40) * Screen.Width / 30 + Screen.Height / 2
Spr = 1 + Spr
If Form2.Top = ((Screen.Width / 2) - (Form2.Width / 2)) Then
Spr = 0
TimerSpr.Enabled = False
End If
End Sub
