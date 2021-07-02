VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   1890
   ClientTop       =   1980
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   497
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Нет
      DrawWidth       =   2
      Height          =   4935
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Пиксель
      ScaleWidth      =   481
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Написано Vir <ou953@ipn.ru>

Dim zap

Private Sub Form_Load()
    z0 = -20
    Dim a As Variant
    Dim Width As Integer
    Dim Height As Integer
    If Form1.Picture1.ScaleHeight < Form1.Picture1.ScaleWidth Then
        zs = Form1.Picture1.ScaleHeight * 5
    Else
        zs = Form1.Picture1.ScaleWidth * 5
    End If
    xmax = Form1.Picture1.ScaleWidth
    ymax = Form1.Picture1.ScaleHeight

'    Call LoadGas(a, Width, Height)
    Call AddMass
End Sub

'Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 1 Then
'        ax = ax + (y - dy) / 5
'        ay = ay + (x - dx) / 5
'        MoveG = True
'
'    End If
'    dx = x
'    dy = y
'End Sub

Private Sub Timer1_Timer()
    ax = Sin(Timer) * Timer / 1000
    ay = Cos(Timer) * Timer / 1000
    az = Sin(Timer) * Timer / 1000
    Call Resize
End Sub

