VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Хранитель Экрана "
   ClientHeight    =   3600
   ClientLeft      =   900
   ClientTop       =   1335
   ClientWidth     =   4725
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "main.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2220
      Picture         =   "main.frx":000C
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   1
      Top             =   0
      Width           =   2385
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   0
      Picture         =   "main.frx":20A7
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   0
      Top             =   0
      Width           =   2205
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3120
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SanKo Code [sankocom@yandex.ru]

Option Explicit
Private MI() As Byte 'Массив цветов
Private MM() As Integer 'Массив ярких букв
Private MF() As Integer 'Массив Z
Private Xt As Integer, Yt As Integer  'Ширина и высота символа
Private Xc As Integer, Yc As Integer 'Количестро символов по вертикали и горизонтали
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            ShowCursor 1
            End 'Escape - выход
End Sub

Private Sub Form_Load()
Dim hwnd As Long
On Error Resume Next
    Select Case Left(Command, 2)
        Case ""
             Start
        Case "/p"
            hwnd = Right(Command, Len(Command) - 3) 'Это hwnd маленького окошка просмотра
            'При выборе
            End
        Case "/s"
            'При нажатии на кнопку [Просмотр]
            Start
        Case "/c"
            'При нажатии на кнопку [Настройка]
            frmAbout.Show
    End Select
End Sub

Private Sub Form_Resize()
    Picture2.Left = Me.ScaleWidth - Picture2.Width
    Picture1.Top = Me.ScaleHeight - Picture1.Height
    Picture2.Top = Picture1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor 1
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Timer1_Timer() 'Обновление символов
Dim i As Integer
Dim k As Integer
Dim a As Long
Dim xm As Integer
Dim ym As Integer
                On Error Resume Next
Me.Cls
    For i = 0 To Xc
    For k = 0 To Yc
        Me.CurrentX = i * Xt
        Me.CurrentY = k * Yt
        Me.ForeColor = RGB(0, MI(i, k) / 1.5, 0) '1.5 - Яркость следа
        Me.FontSize = MF(i)
        If MI(i, k) = 255 Then Me.ForeColor = RGB(100, 255, 100) 'Цвет яркого символа
        If MI(i, k) >= 50 Then Me.Print Chr(Rnd * 255)
    Next k
    Next i
End Sub

Private Sub Timer2_Timer() 'Падение символов (скорость можно поставить меньше)
Dim r As Integer
Dim n As Integer
Dim i As Integer
Dim k As Integer
                On Error Resume Next
r = Rnd * Xc
MF(r) = 20
MM(r) = 0
        For i = 0 To Xc
            MM(i) = MM(i) + 1
            MF(i) = MF(i) - 1
            If MF(i) <= 8 Then MF(i) = 20
            If MM(i) >= Yc + 50 Then '50 - Можно изменить на 0 получится пол
                MM(i) = 0
                MF(r) = 20
            End If
                For k = 0 To MM(i)
                        MI(i, MM(i) - k) = 255 - k * 8 '10-Длинна следа (чем меньше тем длинее)
                Next k
        Next i
End Sub

Private Sub Start() 'Старт, По умолчанию
Dim i As Integer
ShowCursor 0
Me.Show
    Timer1.Interval = 1
    Timer2.Interval = 1
    Xt = Me.TextWidth("X")
    Yt = Me.TextHeight("X")
    Xc = Int(Me.ScaleWidth / Xt) - 1
    Yc = Int(Me.ScaleHeight / Yt) - 1
        ReDim MI(Xc, Yc) As Byte
        ReDim MM(Xc) As Integer
        ReDim MF(Xc) As Integer
            For i = 1 To Xc
            MM(i) = Rnd * Yc
            MF(i) = Rnd * 20
        Next i
        Randomize
End Sub
