VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eraser Electro3Ddance"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "sp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7560
      Top             =   1080
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6960
      Top             =   600
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6960
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      Height          =   1845
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   255
   End
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   8040
      Top             =   1080
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   600
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7560
      Top             =   240
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFF00&
      Height          =   1845
      Left            =   5760
      TabIndex        =   5
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8040
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8040
      Top             =   240
   End
   Begin MCI.MMControl MM1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   240
   End
   Begin VB.OptionButton Op4 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op3 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      MaskColor       =   &H80000004&
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op1 
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1200
      X2              =   1680
      Y1              =   1560
      Y2              =   1080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1680
      X2              =   2160
      Y1              =   1080
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   1080
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1680
      X2              =   2160
      Y1              =   1080
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1200
      X2              =   1680
      Y1              =   600
      Y2              =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Long
Dim le1 As Integer
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
Dim k As Currency
Dim a As Long
Dim er As Long

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

Private Sub File1_DblClick()
li1 = File1.ListIndex
MM1.Command = "stop"
MM1.Command = "close"
MM1.FileName = "Data\music\" & File1.FileName
MM1.Command = "open"
MM1.Command = "play"
Timer8.Enabled = True
End Sub

Private Sub MM1_Done(NotifyCode As Integer)
If Timer8.Enabled = False Then
    If (li1 + 1) = File1.ListCount Then
    li1 = -1
    End If
li1 = li1 + 1
On Error GoTo er1:      ''''''''''' если музыку не находит - не играет''''
MM1.Command = "stop"
MM1.Command = "close"
File1.ListIndex = li1
MM1.FileName = "Data\music\" & File1.FileName
MM1.Command = "open"
MM1.Command = "play"
End If
er1:
End Sub
Private Sub Form_Click()
a = 5000
End Sub
Private Sub Form_Load() ''''''''''''музыка'''''
Command1.Top = 6400
Command1.Left = -500
Timer9.Enabled = True
File1.Top = Height + 10
w = 0
On Error GoTo er1:
File1.FileName = "Data\music\"
File1.ListIndex = 0
MM1.EjectEnabled = True
MM1.FileName = "Data\music\" & File1.FileName
MM1.Command = "open"
MM1.Command = "play"
er1:
End Sub
Private Sub Timer1_Timer()

If h1 + 30 > Width Then
Timer2.Enabled = True
a = a + 1
End If
If a > 5000 Then
a = 1
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Timer2.Enabled = False
k = Math.Rnd * 100
h = 0
h1 = 0
End If

If k > 55 Then
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
If Command1.Left > 5520 Then
Timer10.Enabled = False
le2 = Command1.Left
End If
End Sub

Private Sub Timer11_Timer()
le1 = le1 + 40
Command1.Left = le1 + le2
File1.Left = Command1.Left + Command1.Width
If Command1.Left > 8640 Then
Timer11.Enabled = False
le4 = Command1.Left
le1 = 0
End If
End Sub

Private Sub Timer12_Timer()
le3 = le3 + 40
Command1.Left = le4 - le3
File1.Left = Command1.Left + Command1.Width
If Command1.Left < 5520 Then
Timer12.Enabled = False
le3 = 0
End If
End Sub

Private Sub Timer2_Timer()
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
er = a
If er > 2000 Then
er = 2000
End If

Op1.Top = Math.Sin(a / 26) * (1000 + er + Math.Rnd * 10) + 2900
Op1.Left = Math.Cos(a / 23) * (1000 + er + Math.Rnd * 10) + 2900

Op2.Top = Math.Sin(a / 25) * (1000 + er + Math.Rnd * 10) + 3900
Op2.Left = Math.Cos(a / 28) * (1000 + er + Math.Rnd * 10) + 3900

Op3.Top = Math.Sin(a / 30) * (1000 + er + Math.Rnd * 10) + 4900
Op3.Left = Math.Cos(a / 24) * (1000 + er + Math.Rnd * 10) + 4900

Op4.Top = Math.Sin(a / 25) * (1000 + Math.Rnd * 10) + 5700
Op4.Left = Math.Cos(a / (22)) * (1000 + Math.Rnd * 10) + 5600

'''''''''''рисование линии и её изменение цвета в цикле(255) на 1 миллисекунду''''
i = 255

sled

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

Line1.BorderColor = RGB(0, w, w - 40)
Line2.BorderColor = RGB(0, w, w - 40)
Line3.BorderColor = RGB(0, w, w - 40)
Line4.BorderColor = RGB(0, w, w - 40)
Line5.BorderColor = RGB(0, w, w - 40)
Line6.BorderColor = RGB(0, w, w - 40)

End Sub
Private Sub Timer3_Timer() '''''''''''''шторки''''''
If h < Height Then
h = h + 15
Line (h1 - h, 0)-(Width, h), RGB(0, 0, h / 13)
End If
If h1 < Width Then
h1 = h1 + 15
Line (0, h - h1)-(h1, Height), RGB(0, 0, h1 / 13)
End If
End Sub
Private Sub Timer4_Timer() ''''''''''''крест'''''''''''''''''''
If h < Height Then
h = h + 15
Line (0, h)-(Width, h), RGB(0, 0, h / 13)
End If
If h1 < Width Then
h1 = h1 + 15
Line (h1, 0)-(h1, Height), RGB(0, 0, h1 / 13)
End If
End Sub
Function sled()
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
Private Sub Timer8_Timer()
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
Timer7.Enabled = True
Timer10.Enabled = True
Timer9.Enabled = False
End Sub

