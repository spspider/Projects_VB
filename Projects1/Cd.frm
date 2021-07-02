VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CD-ROM"
   ClientHeight    =   2520
   ClientLeft      =   11490
   ClientTop       =   8595
   ClientWidth     =   3870
   Icon            =   "Cd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   4
      Left            =   0
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton CommandC 
      BackColor       =   &H00FF8080&
      Caption         =   "Close"
      Height          =   2520
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   645
   End
   Begin VB.CommandButton CommandO 
      BackColor       =   &H00FF00FF&
      Caption         =   "Open"
      Height          =   2520
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   645
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
Dim v13 As Long
Dim a As Long
Dim v As Long
Dim cr1 As Long
Dim cg1 As Long
Dim cb1 As Long

Dim cr2 As Long
Dim cg2 As Long
Dim cb2 As Long
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

Private Sub Form_Click()
Timer3.Enabled = False
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
End Sub

Private Sub Form_DblClick()
End
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
Timer2.Interval = 1000
If Err.Number = 380 Then
Timer1.Interval = 0
End If
End Sub
Private Sub Timer2_Timer()
If Err.Number = 380 Then
Timer1.Interval = 0
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
Timer3.Enabled = True
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)

Timer3.Enabled = False
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

Private Sub Timer3_Timer()

cr1 = Abs((Sin(v13 / 500) * 250))
cg1 = Abs((Sin(v13 / 350) * 250))
cb1 = Abs((Sin(v13 / 500) * 250))

'cr2 = Abs((Sin(v13 / 410) * 250))
'cg2 = Abs((Sin(v13 / 260) * 250))
'cb2 = Abs((Sin(v13 / 410) * 250))

cr2 = 250 - cr1
cg2 = 250 - cg1
cb2 = 250 - cb1

CommandO.BackColor = RGB(cr1, cg1, cb1)
CommandC.BackColor = RGB(cr2, cg2, cb2)

Form1.Cls
v13 = v13 + 2
For a3 = 1 To 150
op1T3 = Sin(((v13 / 5) + a3 / 20) / 2 + 1 / 3000) * 1000 + Form1.Height / 2 '- 400
op1L3 = Cos(((v13 / 5) + a3 / 20) / 2 + 1 / 3000) * 1000 + Form1.Width / 2
Line (op1L3, op1T3)-(op2L3, op2T3), RGB(a3 * 8 / 6 + v, a3 * 8 / 6 + v, v * 2 + a3 * 8 / 6) ', B
op2T3 = Form1.Height / 2 ' - 400 '- op1T3
op2L3 = Form1.Width / 2 '- op1L3
Next
End Sub
