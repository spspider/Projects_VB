VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD-ROM"
   ClientHeight    =   2160
   ClientLeft      =   9915
   ClientTop       =   8145
   ClientWidth     =   4410
   Icon            =   "Cd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4410
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton CommandC 
      Caption         =   "Закрыть"
      Height          =   600
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton CommandO 
      BackColor       =   &H80000007&
      Caption         =   "Открыть"
      Height          =   600
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu MnuExit 
         Caption         =   "выход"
      End
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
Dim Status As Integer
Private Sub Command4_Click()
End
End Sub
'Использование:
Private Sub CommandO_Click()
Status = mciSendString("Set CDAudio Door Open Wait", 0&, 0, 0)
End Sub
Private Sub CommandC_Click()
Timer1.Interval = 1000
Status = mciSendString("Set CDAudio Door Closed Wait", 0&, 0, 0)
End Sub
Private Sub Form_Load()
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
nid.szTip = "Мой CD-RW" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid



Form1.Visible = False
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

Private Sub MnuExit_Click()
End
End Sub
Private Sub Timer1_Timer()
On Error GoTo Error: ' Если происходит ошибка то пропускаем это место
ProgressBar1.Value = ProgressBar1.Value + X ' Если отжата то выключаем
Timer2.Interval = 1000
Error:
If Err.Number = 380 Then
Timer1.Interval = 0
End If
End Sub
Private Sub Timer2_Timer()
On Error GoTo Error1: ' Если происходит ошибка то пропускаем это место
ProgressBar1.Value = ProgressBar1.Value - Y ' Если отжата то выключаем
Error1:
If Err.Number = 380 Then
Timer1.Interval = 0
ProgressBar1.Value = 0
End If
End Sub
'На форме в разделе General объявляем переменную определенную как тип пользователя:
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Объявляем переменные:
Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONUP
'Сюда ты можешь вставить код, который захчешь:
Form1.Visible = True
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
''''''''''''''''''''''''''''''''''''''''''
Cancel = True
Form1.Visible = False
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст:
nid.szTip = "Мой CD-RW" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
''''''''''''''''''''''''''''''''''''''''''

End Sub

