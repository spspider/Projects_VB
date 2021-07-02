VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Удалить"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Изменить"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'На форме в разделе General объявляем переменную определенную как тип пользователя
Dim nid As NOTIFYICONDATA



Private Sub Command1_Click()
' Добавить иконку формы в Traybar
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст: "И не забудь зайти на VBStreets.Narod.RU"
nid.szTip = "И не забудь зайти на VBStreets.Narod.RU" & vbNullChar

Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Command2_Click()
nid.hIcon = Form2.Icon
nid.szTip = "New Icon" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub Command3_Click()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Объявляем переменные
Dim msg As Long
Dim sFilter As String

msg = X / Screen.TwipsPerPixelX
Select Case msg


       Case WM_LBUTTONDOWN
       'Сюда ты можешь вставить код, который захчешь
       MsgBox "Нажата левая кнопка мыши(Нажата)"
       
       
       Case WM_LBUTTONUP
            'Сюда ты можешь вставить код, который захчешь
         MsgBox "Нажата левая кнопка мыши(Отжата)"
         
         
       Case WM_LBUTTONDBLCLK
     MsgBox "Ты кликнул 2 раза по ИКОНКЕ(Левой кнопкой)"
         'Сюда ты можешь вставить код, который захчешь
       Case WM_RBUTTONDOWN
           'Сюда ты можешь вставить код, который захчешь
             'Обычно это PopupMenu
             
            MsgBox "Нажата правая кнопка мыши(Нажата)"

      
       Case WM_RBUTTONUP
            'Сюда ты можешь вставить код, который захчешь
             MsgBox "Нажата левая кнопка мыши(Отжата)"
             
       Case WM_RBUTTONDBLCLK
       
           'Сюда ты можешь вставить код, который захчешь
              MsgBox "Ты кликнул 2 раза по ИКОНКЕ(Правой кнопкой)"
End Select

End Sub
