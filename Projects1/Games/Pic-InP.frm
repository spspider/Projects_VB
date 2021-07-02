VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   Caption         =   "ARMGAN by Sp. мой e-m@il: SpSpider@rambler.ru"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Pic-InP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   -120
      Picture         =   "Pic-InP.frx":5A4A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   6120
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   5040
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   5160
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3120
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   840
      Top             =   2880
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   600
      Top             =   3960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1455
      Left            =   5280
      ScaleHeight     =   97.98
      ScaleMode       =   0  'User
      ScaleWidth      =   92.381
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   1440
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Уровень:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Excellent!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2520
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COOL!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   3240
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Всего:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "WAU!!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   3480
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   9720
      TabIndex        =   5
      Top             =   6840
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1605
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Очки:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu main 
      Caption         =   "Меню"
      Begin VB.Menu option 
         Caption         =   "Опции"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim a As Long
Dim a1 As Long
Dim viz As Long
Dim viz2 As Long
Dim viz12 As Long
Dim X As Long

Dim level As Long

Dim Dam1 As Long
Dim Dam2 As Long
Dim Dam3 As Long
Dim Dam4 As Long

Dim x1 As Long
Dim y1 As Long

Dim Xm As Long
Dim Ym As Long

Dim sc As Long
Dim scNV As Long
Dim sc2 As Long
Dim TimeOut As Long

Dim op1T As Long
Dim op1L As Long
Dim op2T As Long
Dim op2L As Long

Dim op1T1 As Long
Dim op1L1 As Long
Dim op2T1 As Long
Dim op2L1 As Long

Dim q As Boolean
Dim q1 As Long

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub File1_Click()
Picture1.Picture = LoadPicture("Picture\" + File1.FileName)
Timer3.Enabled = True
End Sub
Private Sub Form_Click()
If Timer9.Enabled = False Then
sc = sc - 1
'Call mciExecute("play data\music\out.wav")
End If
If Timer9.Enabled = True Then
sc = sc + 1
scNV = scNV + 1
'Call mciExecute("play data\music\in.wav")
End If


If (Round(sc / 2) - sc / 2) = 0 Then
q1 = True
Else
q1 = False
End If

End Sub

Private Sub Form_Load()
Form1.Icon = Picture2.Picture
On Error GoTo er:
'Call mciExecute("play data\music\LP(inMP).mp3")
File1.Left = Width
File1.FileName = "Picture"
File1.ListIndex = 0
Picture1.Picture = LoadPicture("Picture\" + File1.FileName)
er:
TimeOut = 5000

v2 = 100

Dam1 = 80
Dam2 = 70
Dam3 = 80
Dam4 = 60

'Dam1 = 110
'Dam2 = 140
'Dam3 = 160
'Dam4 = 140
Label10.Caption = 0

End Sub

Private Sub Form_Resize()
Timer3.Enabled = True
Label4.Top = Height - 1500
Label4.Left = Width - 1500
Label5.Top = Height - (Label5.Height + 1000)
Label5.Left = Width / 2 - 2000
Label9.Top = Height - (Label5.Height + 1000)
Label9.Left = Width / 2 - 2000
Label8.Top = Height - (Label5.Height + 1000)
Label8.Left = Width / 2 - 2000
End Sub

Private Sub option_Click()
If q = False Then
Timer2.Enabled = True
End If
If q = True Then
Timer3.Enabled = True
End If
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo er1:
'Call mciExecute("stop data\music\in.wav")
sc = sc + 1
sc2 = sc2 + 1
scNV = scNV + 1
'Call mciExecute("play data\music\in.wav")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (Round(sc / 2) - sc / 2) = 0 Then
q1 = True
Else
q1 = False
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
er1:
If Timer9.Enabled = True Then
scNV = scNV + 10
sc = sc + 5
End If

End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1
End Sub

Private Sub Timer1_Timer()

Form1.Cls
v1 = v1 + 1
For a = 1 To 100
op1T = Sin(v1 / 20 + a / 4) * 100 + Sin(v1 / 10 + a / 10) * 500 + Height - 2000
op1L = Cos(-v1 / 10 + a / 10) * 100 + Cos(v1 / 50 + a / 30) * 3000 + Width - 4000
Circle (op1L, op1T), Abs(Sin(a / 33)) * 150, RGB(a * 2, Dam1, 255 - a)
Next
If v1 > 20000 Then
v1 = 0
End If

If v2 <> 100 Then

For a1 = 1 To 50
op1T1 = Cos(v3 / 5 + a1 / 9) * Sin(v3 / 20) * 5000 + Ym
op1L1 = Sin(v3 / 5 + a1 / 9) * Sin(v3 / 20) * 5000 + Xm
Circle (op1L1, op1T1), Abs(Sin(a1 / 33)) * 150, RGB(a1 * 5, a1 * 5, 255 - a1)
Next

End If
End Sub



Private Sub Timer2_Timer()
X = X + 1
File1.Left = Width - Sin(X / 30) * 2600
If File1.Left < Width - (File1.Width + 100) Then
Timer2.Enabled = False
q = True
X = 0
End If
End Sub
Private Sub Timer3_Timer()
X = X + 1
File1.Left = Width - Cos(X / 30) * 2600
If File1.Left > Width + 500 Then
Timer3.Enabled = False
q = False
X = 0
End If
End Sub
Private Sub Timer4_Timer()
Label2.Caption = sc
x1 = x1 + 1
If q1 = False Then
Picture1.Left = Tan(x1 / (Dam1 - sc + 1)) * 1000 + Width / 2
Picture1.Top = Cos(x1 / (Dam2 - sc + 1)) * 2000 + Height / 2
End If
If q1 = True Then
Picture1.Left = Cos(x1 / (Dam3 - sc)) * 1000 + Width / 2
Picture1.Top = Tan(x1 / (Dam4 - sc)) * 2000 + Height / 2
End If
End Sub

Private Sub Timer5_Timer()
On Error GoTo er2:
Label3.Caption = sc2
Label7.Caption = scNV
If Picture1.Left > 0 Then
    If Picture1.Left < Width Then
        If sc2 > 1 Then
        'Call mciExecute("stop data\music\mystC.wav")
        'Call mciExecute("play data\music\mystC.wav")
        Timer6.Enabled = True
        TimeOut = TimeOut + 2
        viz = 0
        viz2 = 0
         End If
         If sc2 > 2 Then
        'Call mciExecute("play data\music\shokH.wav")
        Timer7.Enabled = True
        TimeOut = TimeOut + 3
        viz = 0
        viz2 = 0
        End If
    If sc2 > 3 Then
        'Call mciExecute("play data\music\levelUP.wav")
        Timer8.Enabled = True
        Timer9.Enabled = True
        TimeOut = TimeOut + 4
        viz = 0
        viz2 = 0
        End If
                Else
    sc2 = 0
End If
                Else
    sc2 = 0
    End If
    TimeOut = TimeOut - 1
    Label4 = TimeOut
If TimeOut <= 0 Then
    'Call mciExecute("stop data\music\LP(inMP).mp3")
   'Call mciExecute("play data\music\Alteration_Cast.wav")
    MsgBox "Время вышло, ваш рекорд:  " & scNV & "  Ваши очки: " & sc & " ", 64, "Info"
    Unload Me
End If
er2:
If sc >= 40 Then
    'Call mciExecute("pause data\music\LP(inMP).mp3")
    'Call mciExecute("play data\music\Mysticism_Cast.wav")
    'MsgBox "Уровень пройден! Мои поздравления!  Ваши очки: " & scNV & " Скорость увеличена на:" & (Dam1 - sc + 1), 64, "Info"
    
    'Call mciExecute("play data\music\LP(inMP).mp3")
sc = 0
level = level + 1
Dam1 = Dam1 - 10
Dam2 = Dam2 - 5
Dam3 = Dam3 - 10
Dam4 = Dam4 - 3

If Dam1 < 80 Then
    sc = 0
    Dam1 = 80
    Dam2 = 60
    Dam3 = 80
    Dam4 = 60
    End If
End If
Label10.Caption = level
End Sub

Private Sub Timer6_Timer()
viz = viz + 1
If viz > 10 Then
Label5.Visible = True
End If
If viz > 20 Then
Label5.Visible = False
viz = 0
viz2 = viz2 + 1
End If
If viz2 > 2 Then
Timer6.Enabled = False
viz2 = 0
viz = 0
End If
End Sub

Private Sub Timer7_Timer()

Timer6.Enabled = False
Label5.Visible = False
viz = viz + 1
If viz > 10 Then
Label8.Visible = True
End If
If viz > 20 Then
Label8.Visible = False
viz = 0
viz2 = viz2 + 1
End If
If viz2 > 2 Then
Timer7.Enabled = False
viz2 = 0
viz = 0
End If
End Sub
Private Sub Timer8_Timer()

Timer6.Enabled = False
Timer7.Enabled = False
Label5.Visible = False
Label8.Visible = False

viz = viz + 1
If viz > 10 Then
Label9.Visible = True
End If
If viz > 20 Then
Label9.Visible = False
viz = 0
viz2 = viz2 + 1
End If
If viz2 > 2 Then
Timer8.Enabled = False
viz2 = 0
viz = 0
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
Picture1.BorderStyle = 0
End Sub

Private Sub Timer9_Timer()
v2 = v2 - 1
v3 = v3 + 1

If v2 < 0 Then
Timer9.Enabled = False
v2 = 100
'Call mciExecute("play data\music\mystFAIL.wav")

'v3 = 0
End If
End Sub

