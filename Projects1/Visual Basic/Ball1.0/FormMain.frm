VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   0  'None
   Caption         =   "ВолБол"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictureBall 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1030
      Left            =   0
      Picture         =   "FormMain.frx":1472
      ScaleHeight     =   1035
      ScaleMode       =   0  'User
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Width           =   1030
   End
   Begin VB.Timer TimerMoveBall 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9300
      Top             =   1620
   End
   Begin VB.Timer TimerRotateBall 
      Interval        =   100
      Left            =   8760
      Top             =   1620
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   7
      Left            =   7620
      Picture         =   "FormMain.frx":2C64
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   6
      Left            =   6540
      Picture         =   "FormMain.frx":4456
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   5
      Left            =   5460
      Picture         =   "FormMain.frx":5C48
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   4
      Left            =   4380
      Picture         =   "FormMain.frx":743A
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   3
      Left            =   3300
      Picture         =   "FormMain.frx":8C2C
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   2
      Left            =   2220
      Picture         =   "FormMain.frx":A41E
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   1
      Left            =   1140
      Picture         =   "FormMain.frx":BC10
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image ImageBall 
      Height          =   1005
      Index           =   0
      Left            =   60
      Picture         =   "FormMain.frx":D402
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------
Dim FlagGame As Boolean
Dim FlagRight As Boolean 'с какой стороны подача
'-----------------------------------------------------
Const sEmpty = ""
Const MinMeHeight = 6000 'мин высота
Const MinMeWidth = 1.4 * MinMeHeight 'мин ширина
'-----------------------------------------------------
Dim FlagImageBall As Integer ' картинка мяча
Dim SingleRotateBall As Integer ' вращение мяча
'-----------------------------------------------------
Dim SingleMoveBallX As Integer
Dim SpeedMoveBallX As Single
Dim SingleMoveBallY As Integer
Dim SpeedMoveBallY As Single
Dim SpeedMoveBallYg As Single
'-----------------------------------------------------
Const g = 10 / 30 'ускорение от силы тяжести
Const KoefMirror = 0.7 'коэфф отражения
'Const KoefMirror = 1 'тест
'-----------------------------------------------------
Dim MeScaleLeft As Single
Dim MeScaleTop As Single
'-----------------------------------------------------
Dim MeScaleWidthOld As Single
Dim MeScaleHeightOld As Single
'-----------------------------------------------------



'-----------------------------------------------------
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'-----------------------------------------------------



'двигать мячик2
Private Sub MoveBall2(PictureVollMen As Object, FlagMove As Boolean)

FlagGame = True

'начать движение
If TimerMoveBall.Enabled = False Then TimerMoveBall.Enabled = True

'поднять мяч с пола
If PictureBall.Top = Screen.Height - PictureBall.Height Then PictureBall.Top = PictureBall.Top - 1

'скорость удара
If FlagMove = True Then
    
    SpeedMoveBallX = Abs(PictureBall.Left - PictureVollMen.Left) / 10
    SpeedMoveBallY = Abs(PictureBall.Top - PictureVollMen.Top) / 10

    SingleRotateBall = SingleMoveBallX
    
    'вращение
    If NRnd(1, 2) = 1 Then
    TimerRotateBall.Interval = NRnd(70, 130)
    TimerRotateBall.Enabled = True
    Else
    TimerRotateBall.Enabled = False
    End If

Else
    
    SpeedMoveBallX = KoefMirror * SpeedMoveBallX
    If SpeedMoveBallX < 0.1 Then SpeedMoveBallX = 0
    SpeedMoveBallY = KoefMirror * (SpeedMoveBallYg + SpeedMoveBallY * SingleMoveBallY)
    'SpeedMoveBallY = KoefMirror * SpeedMoveBallY
    If SpeedMoveBallY < 0.1 Then SpeedMoveBallY = 0

End If

SingleMoveBallX = Sgn(PictureBall.Left - PictureVollMen.Left)
SingleMoveBallY = -1 'Sgn(PictureBall.Top - PictureVollMen.Top)

SpeedMoveBallYg = 0

End Sub





Private Sub Form_Load()
Dim mRGN As Long
'-----------------------------------------------------
Me.Width = PictureBall.Width
Me.Height = PictureBall.Height
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
'-----------------------------------------------------
mRGN = CreateEllipticRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY)  'Create an elliptical region
SetWindowRgn Me.hwnd, mRGN, True 'Set the window region
'-----------------------------------------------------
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'-----------------------------------------------------
'AddIconTrey 'добавить иконку в трей
'FlagShow = True 'форма видна
'-----------------------------------------------------
NewGame
'-----------------------------------------------------
End Sub



Private Sub NewGame()
'-----------------------------------------------------
TimerMoveBall.Enabled = False
FlagGame = False
If FlagRight = True Then FlagRight = False Else FlagRight = True  'с какой стороны подача
'-----------------------------------------------------
FlagImageBall = 0 'начальная картинка мяча
'-----------------------------------------------------
'тест
'PictureVollMenL.Left = -10000
'PictureVollMenR.Left = -10000
'-----------------------------------------------------
End Sub







'мышкой по мячу
Private Sub PictureBall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FlagGame = True

'начать движение
If TimerMoveBall.Enabled = False Then TimerMoveBall.Enabled = True

'поднять мяч с пола
If PictureBall.Top = Screen.Height - PictureBall.Height Then PictureBall.Top = PictureBall.Top - 1

'скорость удара
SpeedMoveBallX = Abs(PictureBall.Width / 2 - X) / 5
SingleMoveBallX = Sgn(PictureBall.Width / 2 - X)
SpeedMoveBallY = Abs(PictureBall.Height / 2 - Y) / 5
SingleMoveBallY = Sgn(PictureBall.Height / 2 - Y)
SingleRotateBall = SingleMoveBallX
SpeedMoveBallYg = 0

'вращение
If NRnd(1, 2) = 1 Then
TimerRotateBall.Interval = NRnd(70, 130)
TimerRotateBall.Enabled = True
Else
TimerRotateBall.Enabled = False
End If

If Shift = vbShiftMask Then 'нажать Shift
Unload Me
End If

End Sub




'-----------------------------------------------------
'генерация случайного числа
'-----------------------------------------------------
Private Function NRnd(Nmin As Integer, Nmax As Integer) As Integer
Randomize
NRnd = Int((Nmax - Nmin + 1) * Rnd) + Nmin 'произвольный номер
End Function
'-----------------------------------------------------
'генерация случайного числа
'-----------------------------------------------------







Private Sub PictureMain_DblClick()
'тест
NewGame
End Sub



'движение мяча
Private Sub TimerMoveBall_Timer()
MoveBall 'движение мяча

'If SpeedMoveBallX = 0 And SpeedMoveBallY = 0 Then TimerMoveBall.Enabled = False

'KnockBall False

'тест
'PictureMain.PaintPicture PictureBall, PictureBall.Left + 1000, PictureBall.Top + 1000, , , , , , , vbSrcCopy
End Sub


'движение мяча
Private Sub MoveBall()


'-----------------------------------------------------
'-----------------------------------------------------
'вертикальное направление
'-----------------------------------------------------
FormMain.Top = FormMain.Top + (SpeedMoveBallYg + SpeedMoveBallY * SingleMoveBallY) * TimerMoveBall.Interval
'-----------------------------------------------------
'границы поля
If FormMain.Top + FormMain.Height > Screen.Height Then
    FormMain.Top = Screen.Height - FormMain.Height

    SpeedMoveBallY = KoefMirror * (SpeedMoveBallYg + SpeedMoveBallY * SingleMoveBallY)
    'SpeedMoveBallY = KoefMirror * SpeedMoveBallY
    If SpeedMoveBallY < 0.1 Then SpeedMoveBallY = 0
    SpeedMoveBallYg = 0
    SingleMoveBallY = -1
    SingleRotateBall = SingleMoveBallX
    Exit Sub
End If
If FormMain.Top < 0 Then
    FormMain.Top = 0
    
    SpeedMoveBallY = KoefMirror * SpeedMoveBallY
    If SpeedMoveBallY < 0.1 Then SpeedMoveBallY = 0
    SpeedMoveBallYg = 0
    SingleMoveBallY = 1
    SingleRotateBall = -SingleMoveBallX
    Exit Sub
End If
'-----------------------------------------------------
    SpeedMoveBallYg = SpeedMoveBallYg + g * TimerMoveBall.Interval
'-----------------------------------------------------
'вертикальное направление
'-----------------------------------------------------
'-----------------------------------------------------



'-----------------------------------------------------
'-----------------------------------------------------
'горизонтальное направление
'-----------------------------------------------------
FormMain.Left = FormMain.Left + SpeedMoveBallX * SingleMoveBallX * TimerMoveBall.Interval
'-----------------------------------------------------
'границы поля
If FormMain.Left + FormMain.Width > Screen.Width Then
    FormMain.Left = Screen.Width - FormMain.Width
        
    SpeedMoveBallX = KoefMirror * SpeedMoveBallX
    If SpeedMoveBallX < 0.1 Then SpeedMoveBallX = 0
    SingleMoveBallX = -1
    SingleRotateBall = -SingleMoveBallY
    Exit Sub
End If
If FormMain.Left < 0 Then
    FormMain.Left = 0
        
    SpeedMoveBallX = KoefMirror * SpeedMoveBallX
    If SpeedMoveBallX < 0.1 Then SpeedMoveBallX = 0
    SingleMoveBallX = 1
    SingleRotateBall = SingleMoveBallY
    Exit Sub
End If
'-----------------------------------------------------
'горизонтальное направление
'-----------------------------------------------------
'-----------------------------------------------------

'тест
'Me.Caption = Format(SpeedMoveBallX, "0.0") & "  " & Format(SpeedMoveBallY, "0.0") & "  " & Format(SpeedMoveBallYg, "0.0") & "  " & Format((SpeedMoveBallYg + SpeedMoveBallY * SingleMoveBallY), "0.0") & "  " & SingleMoveBallY

End Sub









'мячик крутится
Private Sub TimerRotateBall_Timer()
PictureBall.Picture = ImageBall(FlagImageBall).Picture
If SingleRotateBall = 1 Then
    If FlagImageBall > 0 Then FlagImageBall = FlagImageBall - 1 Else FlagImageBall = 7
Else
    If FlagImageBall < 7 Then FlagImageBall = FlagImageBall + 1 Else FlagImageBall = 0
End If

End Sub








