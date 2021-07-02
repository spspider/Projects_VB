VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   2745
   ClientTop       =   3930
   ClientWidth     =   14715
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   14715
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   1800
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.FileListBox File4 
      Height          =   1455
      Left            =   4320
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.FileListBox File3 
      Height          =   1455
      Left            =   3120
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.FileListBox File2 
      Height          =   1455
      Left            =   1920
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   480
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   8280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   500
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2640
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2040
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   1800
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   3615
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   8445
      Left            =   13920
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   7320
      TabIndex        =   16
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "bonus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   11760
      TabIndex        =   15
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   12240
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Str, Fstr1, NStr, CodeS, CodeS_1, CodeS_2, Str1 As String
Dim CodeMass(3), Fstr(3) As String
Dim InStrC, X, x1, Push(250), a, shoot, bonus, qstr, Ym, Xm, Timer_2 As Integer
Dim StopPlay, stRic, Timer_3 As Integer
Dim redCheck, redL, shootG As Integer
Dim timer_1 As Long
Dim q, q1, Push2, Push3, w1 As Integer
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Dim GameDir As String
Dim op1L2, op1T2, op2L2, op2T2 As Long

Private Sub Combo1_Change()
Text2.Text = Combo1.Index
End Sub

Private Sub Command1_Click()

Dim i As Integer
Dim a1 As Integer
i = 0
'MSComm1.DTREnable = True
'MSComm1.RTSEnable = True
'For i = 0 To 500
MSComm1.Output = "1"
'Next i
End Sub

Private Sub Command2_Click()
Str = ""
Str1 = ""
Text1.Text = ""
List1.Clear
End Sub
Function openP()

' Fire Rx Event Every Two Bytes
MSComm1.RThreshold = 1
Text1.MaxLength = 900
' When Inputting Data, Input 2 Bytes at a time
MSComm1.InputLen = 3
MSComm1.Settings = "9600,N,8,1"
' Disable DTR
MSComm1.DTREnable = False
MSComm1.RTSEnable = True

' Open COM1
'On Error GoTo er1:
MSComm1.CommPort = (Combo1.ListIndex + 1)
'MSComm1.CommPort = 5
MSComm1.PortOpen = True
'er1:
End Function

Function CodeS1()

Fstr(0) = "#"
If (InStr(Str, Fstr(0)) <> 0) Then 'поиск n в строке
List1.AddItem (Str)
If List1.ListCount >= 50 Then
List1.Clear
End If
'Text5.Text = Mid(Str, 1, Len(Str) - 1)
On Error GoTo er1:
Push(q) = Val(Mid(Str, 1, Len(Str) - 1))
Push3 = Push(q) - Push2

On Error GoTo er:
ProgressBar1.Max = 25
ProgressBar1.Value = Push3
er:
er1:
q = q + 1
Str = ""

End If
End Function
Function PushMast()
If q = 100 Then
q = 0
For q1 = 0 To 99
    If w1 = 1 Then
           
        If Push(q) >= Push(q1 + 1) - 4 And Push(q) <= Push(q1 + 1) + 4 Then
        
        Push2 = Push(q1)
        Else
        w1 = 0
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    End If
Next q1
w1 = 1
End If
'Text4.Text = Push2

qstr = 100 - q
Command5.Caption = "Wait " & qstr
If Push2 <> "" Then
Command5.Visible = False
Timer1.Enabled = True
Picture1.Visible = True
Timer3.Enabled = True
End If

'Text5.Text = Str
End Function
Function CodeS2()
Fstr(0) = "d1"
Fstr(1) = "d2"
Fstr(2) = "d0"
Fstr(3) = "n"

'CodeMass(0) = Text2.Text
'CodeMass(1) = Text3.Text
'CodeMass(2) = Text4.Text

NStr = "#"
Str = Text1.Text
X = 0
For X = 0 To 3
'For x1 = 0 To 3 'обнуление всего диапазона массива
'CodeMass(x1) = 0
'Next x1
If (InStr(Str, Fstr(X)) <> 0) And (InStr(Str, NStr) <> 0) Then
Str = Mid(Str, (InStr(Str, Fstr(X))), 255)
CodeMass(X) = Val(Mid(Str, InStr(Str, Fstr(X)) + 3, InStr(Str, NStr) - InStr(Str, Fstr(X)) + 255)) 'считает длину символов с первой строки
Str = Mid(Str, (InStr(Str, NStr)) + 1, 255)
Else
Str1 = Str
End If
Next X
If (InStr(Str, Fstr(3)) <> 0) Then 'поиск n в строке
Str = Mid(Str, (InStr(Str, Fstr(3))) + 1, 255)
Else
End If

Text2.Text = CodeMass(0)
Text3.Text = CodeMass(1)
Text4.Text = CodeMass(2)

Text1.Text = Str
Text6.Text = Str1
'ProgressBar1.Value =
End Function
Private Sub Command4_Click()
Fstr(0) = "d0"
Fstr(1) = "d1"
Fstr(2) = "d2"
CodeMass(0) = Text2.Text
CodeMass(1) = Text3.Text
CodeMass(2) = Text4.Text

NStr = "#"
'Str = Text5.Text
X = 0
For X = 0 To 3
CodeMass(X) = 0
If (InStr(Str, Fstr(X)) <> 0) And (InStr(Str, NStr) <> 0) Then
Str1 = Mid(Str, (InStr(Str, Fstr(X))), 255)
CodeMass(X) = Val(Mid(Str1, InStr(Str1, Fstr(X)) + 3, InStr(Str1, NStr) - InStr(Str1, Fstr(X)) + 255)) 'считает длину символов с первой строки
Str1 = Mid(Str1, (InStr(Str1, NStr)) + 1, 255)
Str = Mid(Str, (InStr(Str, NStr)) + 1, 255)
Else
Str1 = Str
End If


Next X
Text2.Text = CodeMass(0)
Text3.Text = CodeMass(1)
Text4.Text = CodeMass(2)

'Text5.Text = Str
'Text6.Text = Str1
End Sub



Private Sub Command5_Click()
openP

'Command5.Visible = False

End Sub




'Private Sub MMControl1_PlayCompleted(Errorcode As Long)
'MMControl1.FileName = GameDir + "menu.mp3" ' Открываем выбранный файл
'MMControl1.Command = "open" ' Запускаем
'MMControl1.Command = "play"
'End Sub
Private Sub Command7_Click()
Push3 = 4
End Sub

Private Sub Command8_Click()
Push3 = 0
End Sub

Private Sub Form_Load()
Combo1.AddItem ("com1")
Combo1.AddItem ("com2")
Combo1.AddItem ("com3")
Combo1.AddItem ("com4")
Combo1.AddItem ("com5")
Combo1.ListIndex = 0
redCheck = 0
GameDir = "D:\Data\"
File1.ListIndex = 0
stRic = 0
'WindowsMediaPlayer1
'Timer3.Enabled = True

'If MMControl1.TrackLength = MMControl1.Position Then
MMControl1.Command = "stop"
MMControl1.Command = "close"
'MMControl1.Command = "prev"
'End If

File4.Path = GameDir + "music\"
'File4.ListIndex = Round(Abs(Cos(Rnd * 1000) * File4.ListCount), 0)
File4.ListIndex = 0
MMControl1.FileName = File4.Path + "\" + File4.FileName ' Открываем выбранный файл
'Text1.Text = File4.Path + "\" + File4.FileName ' Открываем выбранный файл
MMControl1.Command = "open" ' Запускаем
MMControl1.Command = "play"
'Text3.Text = MMControl1.TrackLength


End Sub

Private Sub Form_Resize()
List1.Left = Form1.Width - List1.Width - 10
List1.Height = Form1.Height
Command5.Left = Form1.Width / 2 - Command5.Width / 2
Command5.Top = Form1.Height / 2 - Command5.Top / 2
'HScroll1.Top = Form1.Height - HScroll1.Height - 1000
ProgressBar1.Top = Form1.Height - 1000
MMControl1.Top = Form1.Height - 1400
'MMControl1.Left = -MMControl1.Width
Label1.Left = Form1.Width - Label1.Width - 1000
Label3.Top = Form1.Height - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:
MSComm1.PortOpen = False
er:
'Call mciExecute("stop d:\Data\menu.mp3")
End Sub


Private Sub MMControl1_Done(NotifyCode As Integer)

If MMControl1.TrackLength = MMControl1.Position Then
'MMControl1.Command = "prev"
MMControl1.Command = "stop"
MMControl1.Command = "close"

File4.Path = GameDir + "music\"
On Error GoTo er:
File4.ListIndex = Round(Abs(Cos(Rnd * 1000) * File4.ListCount), 0)
er:
'File4.ListIndex = 0
MMControl1.FileName = File4.Path + "\" + File4.FileName ' Открываем выбранный файл
MMControl1.Command = "open" ' Запускаем

'Text1.Text = File4.Path + "\" + File4.FileName ' Открываем выбранный файл
MMControl1.Command = "play"
End If

End Sub

Private Sub MSComm1_OnComm()
Dim sData As String ' Holds our incoming data
If MSComm1.CommEvent = comEvReceive Then
sData = MSComm1.Input ' Get data (2 bytes)
lWord = (sData)
Str = Str & CStr(lWord)
'Text1.Text = Str
Text1.Text = Str
CodeS1
PushMast
End If
End Sub



Function expl()

'''''''''''''''''''''''''''''''''''''''''''''''''
File1.Path = GameDir + "Sounds\exp\"
On Error GoTo er3:
If File1.FileName <> "" Then
Call mciExecute("stop " + File1.Path + "\" + File1.FileName)
End If
er3:
On Error GoTo er2:
File1.ListIndex = Round(Abs(Cos(Rnd * 1000) * File1.ListCount), 0)
'Text3.Text = File1.ListIndex
er2:
On Error GoTo er12:
Call mciExecute("play " + File1.Path + "\" + File1.FileName)
er12:
'''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Function ric()
'''''''''''''''''''''''''''''''''''''''''''''''''
StopPlay = 1
File3.Path = GameDir + "Sounds\ric\"
On Error GoTo er3:
If File3.FileName <> "" Then
Call mciExecute("stop " + File3.Path + "\" + File3.FileName)
End If
File3.ListIndex = 0
er3:
On Error GoTo er2:
File3.ListIndex = Round(Abs(Cos(Rnd * 1000) * File3.ListCount), 0)
'Text3.Text = File1.ListIndex
er2:
On Error GoTo er12:
Call mciExecute("play " + File3.Path + "\" + File3.FileName)
er12:
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
Picture1.BorderStyle = 0
End Sub
Function ShootGT()

Timer5.Enabled = True



End Function

Function find()
'''''''''''''''''''''''''''''''''''''''''''''''''

If redL = 1 Then
File2.Path = GameDir + "Sounds\Finded\"
On Error GoTo er3:
If File2.FileName <> "" Then
Call mciExecute("stop " + File2.Path + "\" + File2.FileName)
End If
er3:
On Error GoTo er2:
File2.ListIndex = Round(Abs(Cos(Rnd * 1000) * File2.ListCount), 0)
'Text3.Text = File1.ListIndex
er2:
On Error GoTo er12:
Call mciExecute("play " + File2.Path + "\" + File2.FileName)
er12:
'''''''''''''''''''''''''''''''''''''''''''''''''''
redL = 0
redCheck = 1
End If
End Function
Private Sub Timer1_Timer()



timer_1 = timer_1 + 1
If PushN = 0 Then
Timer_3 = Timer_3 + 1
If Timer_3 = 10 Then
Timer_3 = 0
PushN = 1
End If
End If

If timer_1 = 2000 Then
timer_1 = 0
End If
Picture1.Left = Tan(timer_1 / 70) * 1500 + Height / 2 + Cos(timer_1 / 70) * 1500
Picture1.Top = Tan(timer_1 / 50) * 1500 + Height / 2 + Sin(timer_1 / 70) * 1500
Form1.Cls

d1 = Tan(timer_1 / 50) * 1000
Y1 = Tan(timer_1 / 50) * 1000

'd1_1 = Picture1.Top / (Form1.Height / (d1 / 2))
'y1_1 = Picture1.Left / (Form1.Width / (Y1 / 2))
d1_1 = Picture1.Width / 2 + d1
y1_1 = Picture1.Height / 2 + Y1

Label1.Caption = shoot
        
If -10 < d1_1 And y1_1 <= 800 Then
    red1 = Abs((Sin(y1_1 / 2) + Sin(d1_1)) * 255)
    green1 = 0
    blue1 = 0
    
    If redCheck = 0 Then
    redL = 1
    redCheck = 1
    find
    End If
    'Call mciExecute("play d:\Data\Sounds\menu.mp3") 'перекрестие найдено


        If Push3 >= 10 Then
       
        
        Picture1.BorderStyle = 1
        expl
        timer_1 = Abs(Rnd * 1000)
        shoot = shoot + 1
        shootG = shootG + 1
        If shootG = 2 Then
        ShootGT
        shootG = 0
        End If
        Call mciExecute("play " + GameDir + "Sounds\Others\shootG.mp3")
        On Error GoTo er4:
        Form1.BackColor = RGB(Push3 * 8, 100, 100)
er4:
        Label1.BackColor = RGB(255, 0, 0)
        End If
        If Push3 <= 2 Then
        Form1.BackColor = RGB(100, 100, 100)
        Picture1.BorderStyle = 0
        End If



    Else
    red1 = 100
    green1 = 100
    blue1 = 100
    Picture1.BorderStyle = 0
    'redL = 0
    redCheck = 0
    If Push3 >= 4 Then
    shootG = 0
    If StopPlay = 0 Then
    ric
    shootG = 0
    If shoot > 1 Then
    shoot = shoot - 1
    End If
    timer_1 = Abs(Rnd * 1000)
    End If
    End If
End If
Line (Picture1.Left + Picture1.Width / 2 + d1, 0)-(Picture1.Left + Picture1.Width / 2 + d1, Form1.Height), RGB(red1, green1, blue1)
Line (0, Picture1.Top + Picture1.Height / 2 + Y1)-(Form1.Width, Picture1.Top + Y1 + Picture1.Height / 2), RGB(red1, green1, blue1)

'червь
For a = 1 To 100
op1T = Sin(timer_1 / 60 + a / 400) * timer_1 * 2 + Sin(timer_1 / 10 + a / 10) * 5000 + Form1.Height / 2 - 2000
op1L = Cos(-timer_1 / 50 + a / 10) * timer_1 * 2 + Cos(timer_1 / 50 + a / 30) * 3000 + Form1.Width / 2
Circle (op1L, op1T), Abs(Sin(a / 33)) * 150, RGB(a * 2, 0, 255 - a)
Next
'пространство

'For a2 = 1 To 100
'op1T2 = Tan(Cos(timer_1 / 100) + a2 / 2) * 5000 + Form1.Height / 2 + 2000
'op1L2 = Tan(-Cos(-timer_1 / 100) + a2 / 2) * 5000 + Form1.Height / 2 + 2000
'Line (op1L2, op1T2)-(op2L2, op2T2), RGB(0, a2 * 2, 0)
'op2T2 = (Height + 3000) - op1T2
'op2L2 = Width - op1L2
'Next



End Sub

Private Sub Timer2_Timer()
stRic = stRic + 1
If stRic = 0 Then
End If
If stRic = 40 Then
StopPlay = 0
stRic = 0
End If
End Sub

Private Sub Timer3_Timer()
bonus = bonus + 1
If bonus = 200 Then
'shoot = 100
If shoot > 0 Then
shoot = shoot - 1
On Error GoTo er11:
Call mciExecute("play " + GameDir + "sounds\Others\snap.mp3")
er11:
Label1.BackColor = RGB(0, 255, 255)
End If
bonus = 0
End If
Label2.Caption = bonus
End Sub

Private Sub Timer4_Timer()
Label3.Caption = Time
End Sub

Private Sub Timer5_Timer()
Timer_2 = Timer_2 + 1
For a1 = 1 To 50
op1T1 = Cos(timer_1 / 5 + a1 / 9) * Sin(timer_1 / 20) * 4000 + Picture1.Left - Picture1.Width / 2
op1L1 = Sin(timer_1 / 5 + a1 / 9) * Sin(timer_1 / 20) * 4000 + Picture1.Left - Picture1.Width / 2
Circle (op1L1, op1T1), Abs(Sin(a1 / 33)) * 150, RGB(a1 * 5, a1 * 5, 255 - a1)
Next
If Timer_2 = 100 Then
Timer5.Enabled = False
Timer_2 = 0
End If
End Sub


