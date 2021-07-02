VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1680
   ClientLeft      =   5580
   ClientTop       =   5325
   ClientWidth     =   3060
   Icon            =   "BUD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   3060
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6120
      Top             =   1800
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   5520
      Top             =   1800
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Left            =   5520
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "проговорить"
      Height          =   375
      Left            =   -120
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   0
   End
   Begin VB.Label Label5 
      Caption         =   "23:59:59"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "секунды"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "минуты"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "часы"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim X As Long
Dim nid As NOTIFYICONDATA
Dim Y As Long ' Объявляем переменную для хранения чисел



Dim Ho, Mi, Mi2, Se As Integer
Dim PS As String
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private Sub Command1_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Click()
Cancel = True
Form1.Visible = False
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст:
nid.szTip = "часы" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub



Private Sub Form_Load()
PS = "play sounds\"
Call mciExecute("close all")



X = 100
Y = 10
' Добавить иконку формы в Traybar
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст: "И не забудь зайти на VBStreets.Narod.RU":
nid.szTip = "часы" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid



'Form1.Visible = False
' Добавить иконку формы в Traybar
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст: "И не забудь зайти на VBStreets.Narod.RU":
nid.szTip = "Мой CD-RW" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid


End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Объявляем переменные:
Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONUP
'Сюда ты можешь вставить код, который захчешь:
Form1.Visible = True
'Timer3.Enabled = True
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Form1.Visible = False
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст:
nid.szTip = "часы" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Time
If Len(Time) = 7 Then
'Text2.Text = Left(Time, 1)
'Text3.Text = Mid(Time, 3, 2)
'Text4.Text = Mid(Time, 6, 2)
Ho = Left(Time, 1)
Mi = Mid(Time, 3, 2)
Se = Mid(Time, 6, 2)
Else
'Text2.Text = Left(Time, 2)
'Text3.Text = Mid(Time, 4, 2)
'Text4.Text = Mid(Time, 7, 2)
Ho = Left(Time, 2)
Mi = Mid(Time, 4, 2)
Se = Mid(Time, 7, 2)
End If
'Ho = Text2.Text
'Mi = Text3.Text
'Se = Text4.Text
End Sub

Private Sub Timer2_Timer()
If Ho = "1" Then
Call mciExecute(PS + "1.mp3")
End If

If Ho = "2" Then
Call mciExecute(PS + "2.mp3")
End If

If Ho = "3" Then
Call mciExecute(PS + "3.mp3")
End If
'''''''''''''''''''''''''''''''''
If Ho = "4" Then
Call mciExecute(PS + "4.mp3")
End If

If Ho = "5" Then
Call mciExecute(PS + "5.mp3")
End If

If Ho = "6" Then
Call mciExecute(PS + "6.mp3")
End If

If Ho = "7" Then
Call mciExecute(PS + "7.mp3")
End If

If Ho = "8" Then
Call mciExecute(PS + "8.mp3")
End If

If Ho = "9" Then
Call mciExecute(PS + "9.mp3")
End If

If Ho = "10" Then
Call mciExecute(PS + "10.mp3")
End If

If Ho = "11" Then
Call mciExecute(PS + "11.mp3")
End If

If Ho = "12" Then
Call mciExecute(PS + "12.mp3")
End If

If Ho = "13" Then
Call mciExecute(PS + "13.mp3")
End If

If Ho = "14" Then
Call mciExecute(PS + "14.mp3")
End If

If Ho = "15" Then
Call mciExecute(PS + "15.mp3")
End If

If Ho = "16" Then
Call mciExecute(PS + "16.mp3")
End If

If Ho = "17" Then
Call mciExecute(PS + "17.mp3")
End If

If Ho = "18" Then
Call mciExecute(PS + "18.mp3")
End If

If Ho = "19" Then
Call mciExecute(PS + "19.mp3")
End If

If Ho = "20" Then
Call mciExecute(PS + "20.mp3")
End If

If Ho = "21" Then
Call mciExecute(PS + "21.mp3")
End If

If Ho = "22" Then
Call mciExecute(PS + "22.mp3")
End If

If Ho = "23" Then
Call mciExecute(PS + "23.mp3")
End If
''''''''''''''' задержка проговаривания вренмени
If Ho = 1 Or Ho = 2 Or Ho = 3 Or Ho = 4 Or HHo = 5 Or Ho = 6 Or Ho = 7 Or Ho = 8 Or Ho = 9 Then
Timer5.Interval = 500
End If

If Ho = 10 Or Ho = 11 Or Ho = 12 Or Ho = 13 Or Ho = 14 Or Ho = 15 Or Ho = 16 Or Ho = 17 Or Ho = 18 Or Ho = 19 Or Ho = 20 Or Ho = 21 Or Ho = 22 Or Ho = 23 Then
Timer5.Interval = 800
End If
''''''''''''''''''''''''''''''''''''''''''''''''


Timer5.Enabled = True
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer() '''''''''''''''''''''''''''''' минуты
If Mi = "01" Then
Call mciExecute(PS + "1a.mp3")
End If

If Mi = "02" Then
Call mciExecute(PS + "2e.mp3")
End If

If Mi = "03" Then
Call mciExecute(PS + "3.mp3")
End If
'''''''''''''''''''''''''''''''''
If Mi = "04" Then
Call mciExecute(PS + "4.mp3")
End If

If Mi = "05" Then
Call mciExecute(PS + "5.mp3")
End If

If Mi = "06" Then
Call mciExecute(PS + "6.mp3")
End If

If Mi = "07" Then
Call mciExecute(PS + "7.mp3")
End If

If Mi = "08" Then
Call mciExecute(PS + "8.mp3")
End If

If Mi = "09" Then
Call mciExecute(PS + "9.mp3")
End If

If Mi = "10" Then
Call mciExecute(PS + "10.mp3")
End If

If Mi = "11" Then
Call mciExecute(PS + "11.mp3")
End If

If Mi = "12" Then
Call mciExecute(PS + "12.mp3")
End If

If Mi = "13" Then
Call mciExecute(PS + "13.mp3")
End If

If Mi = "14" Then
Call mciExecute(PS + "14.mp3")
End If

If Mi = "15" Then
Call mciExecute(PS + "15.mp3")
End If

If Mi = "16" Then
Call mciExecute(PS + "16.mp3")
End If

If Mi = "17" Then
Call mciExecute(PS + "17.mp3")
End If

If Mi = "18" Then
Call mciExecute(PS + "18.mp3")
End If

If Mi = "19" Then
Call mciExecute(PS + "19.mp3")
End If

If Mi = "20" Then
Call mciExecute(PS + "20.mp3")
End If

If Mi = "30" Then
Call mciExecute(PS + "30.mp3")
End If

If Mi = "40" Then
Call mciExecute(PS + "40.mp3")
End If

If Mi = "50" Then
Call mciExecute(PS + "50.mp3")
End If

If Mi = 21 Or Mi = 22 Or Mi = 23 Or Mi = 24 Or Mi = 25 Or Mi = 26 Or Mi = 27 Or Mi = 28 Or Mi = 29 Then
Call mciExecute(PS + "20.mp3")
Timer7.Enabled = True
End If

If Mi = 31 Or Mi = 32 Or Mi = 33 Or Mi = 34 Or Mi = 35 Or Mi = 36 Or Mi = 37 Or Mi = 38 Or Mi = 39 Then
Call mciExecute(PS + "30.mp3")
Timer7.Enabled = True
End If

If Mi = 41 Or Mi = 42 Or Mi = 43 Or Mi = 44 Or Mi = 45 Or Mi = 46 Or Mi = 47 Or Mi = 48 Or Mi = 49 Then
Call mciExecute(PS + "40.mp3")
Timer7.Enabled = True
End If

If Mi = 51 Or Mi = 52 Or Mi = 53 Or Mi = 54 Or Mi = 55 Or Mi = 56 Or Mi = 57 Or Mi = 58 Or Mi = 59 Then
Call mciExecute(PS + "50.mp3")
Timer7.Enabled = True
End If

Timer3.Enabled = False

Timer6.Interval = Timer7.Interval + 500
Timer6.Enabled = True
End Sub
Private Sub Timer7_Timer()
Mi2 = Mid(Mi, 2, 1)

If Mi2 = 1 Then
Call mciExecute(PS + "1a.mp3")
End If

If Mi2 = 2 Then
Call mciExecute(PS + "2e.mp3")
End If

If Mi2 = 3 Then
Call mciExecute(PS + "3.mp3")
End If

If Mi2 = 4 Then
Call mciExecute(PS + "4.mp3")
End If

If Mi2 = 5 Then
Call mciExecute(PS + "5.mp3")
End If

If Mi2 = 6 Then
Call mciExecute(PS + "6.mp3")
End If

If Mi2 = 7 Then
Call mciExecute(PS + "7.mp3")
End If

If Mi2 = 8 Then
Call mciExecute(PS + "8.mp3")
End If

If Mi2 = 9 Then
Call mciExecute(PS + "9.mp3")
End If


Timer7.Enabled = False
End Sub

Private Sub Timer5_Timer()

If Ho = 1 Or Ho = 21 Then
Call mciExecute(PS + "chas.mp3")
End If

If Ho = 2 Or Ho = 3 Or Ho = 4 Or Ho = 22 Or Ho = 23 Then
Call mciExecute(PS + "chasa.mp3")
End If

If Ho = 5 Or Ho = 6 Or Ho = 7 Or Ho = 8 Or Ho = 9 Or Ho = 10 Or Ho = 11 Or Ho = 12 Or Ho = 13 Or Ho = 14 Or Ho = 15 Or Ho = 16 Or Ho = 17 Or Ho = 18 Or Ho = 19 Or Ho = 20 Then
Call mciExecute(PS + "chasov.mp3")
End If

Timer5.Enabled = False
Timer3.Interval = 600
Timer3.Enabled = True
End Sub

Private Sub Timer6_Timer()
If Mi = 1 Or Mi = 21 Or Mi = 31 Or Mi = 41 Or Mi = 51 Then
Call mciExecute(PS + "minuta.mp3")
End If

If Mi = 2 Or Mi = 3 Or Mi = 4 Or Mi = 22 Or Mi = 23 Or Mi = 24 Or Mi = 32 Or Mi = 33 Or Mi = 34 Or Mi = 42 Or Mi = 43 Or Mi = 44 Or Mi = 52 Or Mi = 53 Or Mi = 54 Then
Call mciExecute(PS + "minuti.mp3")
End If

If Mi = 5 Or Mi = 6 Or Mi = 7 Or Mi = 8 Or Mi = 9 Or Mi = 10 Or Mi = 11 Or Mi = 12 Or Mi = 13 Or Mi = 14 Or Mi = 15 Or Mi = 16 Or Mi = 17 Or Mi = 18 Or Mi = 19 Or Mi = 20 Or Mi = 25 Or Mi = 26 Or Mi = 27 Or Mi = 28 Or Mi = 29 Or Mi = 30 Or Mi = 35 Or Mi = 36 Or Mi = 37 Or Mi = 38 Or Mi = 39 Or Mi = 40 Or Mi = 45 Or Mi = 46 Or Mi = 47 Or Mi = 48 Or Mi = 49 Or Mi = 50 Or Mi = 55 Or Mi = 56 Or Mi = 57 Or Mi = 58 Or Mi = 59 Then
Call mciExecute(PS + "minut.mp3")
End If
Timer6.Enabled = False
End Sub


