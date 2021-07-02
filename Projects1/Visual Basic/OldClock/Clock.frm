VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Часы"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   FillStyle       =   0  'Заливка
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   338
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameAbout 
      Caption         =   "О программе..."
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   -3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdWWW 
         Caption         =   "http://&www.rashid4ever.narod.ru"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Назад"
         Default         =   -1  'True
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox AboutText 
         Height          =   2535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "Clock.frx":030A
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Выход"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&О программе"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame frSetting 
      Height          =   1455
      Left            =   3480
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox chkDight 
         Caption         =   "&Цифровые"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkHour 
         Caption         =   "&Час"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkMin 
         Caption         =   "&Минута"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkSec 
         Caption         =   "&Секунда"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   20
         X2              =   1420
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   1440
         Y1              =   1000
         Y2              =   1000
      End
   End
   Begin VB.Timer Timing 
      Interval        =   120
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame frClock 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3255
      Begin VB.Shape Shape2 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Заливка
         Height          =   135
         Left            =   1560
         Shape           =   3  'Круг
         Top             =   1680
         Width           =   135
      End
      Begin VB.Line SecP 
         BorderColor     =   &H000000FF&
         Visible         =   0   'False
         X1              =   1620
         X2              =   1620
         Y1              =   1740
         Y2              =   360
      End
      Begin VB.Shape Shape1 
         Height          =   3000
         Left            =   120
         Shape           =   3  'Круг
         Top             =   240
         Width           =   3000
      End
      Begin VB.Line HourP 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   1620
         X2              =   1620
         Y1              =   1740
         Y2              =   1080
      End
      Begin VB.Line MinP 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   1620
         X2              =   840
         Y1              =   1740
         Y2              =   600
      End
      Begin VB.Label Digital 
         Alignment       =   2  'Центровка
         BackStyle       =   0  'Прозрачно
         Caption         =   "00.00.00"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Frame frNfo 
      Height          =   975
      Left            =   3480
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
      Begin VB.TextBox NfoTxT 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Нет
         Height          =   810
         Left            =   40
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Clock.frx":0427
         Top             =   120
         Width           =   1370
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pHour, pMin, pSec, ShowAbout
Private Sub chkDight_Click()
If chkDight.Value = 1 Then Digital.Visible = True Else Digital.Visible = False
End Sub
Private Sub chkHour_Click()
If chkHour.Value = 1 Then HourP.Visible = True Else HourP.Visible = False
End Sub
Private Sub chkMin_Click()
If chkMin.Value = 1 Then MinP.Visible = True Else MinP.Visible = False
End Sub
Private Sub chkSec_Click()
If chkSec.Value = 1 Then SecP.Visible = True Else SecP.Visible = False
End Sub
Private Sub CmdAbout_Click()
frameAbout.Visible = True
ShowAbout = True
End Sub
Private Sub cmdBack_Click()
ShowAbout = False
frameAbout.Top = -225
frameAbout.Visible = False
End Sub
Private Sub cmdExit_Click()
End
End Sub
Private Sub cmdWWW_Click()
Shell "explorer http://www.rashid4ever.narod.ru"
End Sub
Private Sub Form_Load()
frameAbout.Top = -225
chkDight.Value = 1
Digital.Visible = True
chkHour.Value = 1
HourP.Visible = True
chkMin.Value = 1
MinP.Visible = True
chkSec.Value = 1
SecP.Visible = True
NfoTxT.Text = Left$(NfoTxT.Text, Len(NfoTxT.Text) - 2)
AboutText.Text = Left$(AboutText.Text, Len(AboutText.Text) - 2)
End Sub
Private Sub Timing_Timer()
Me.Caption = "Часы: " & DateTime.Time
If ShowAbout = True Then
If frameAbout.Top = 0 Then GoTo 5
frameAbout.Top = Val(frameAbout.Top) + 25
5 End If
20 Rem Clock Pointers moving code starting from here
Digital.Caption = DateTime.Time
pHour = Val(Left$(Digital.Caption, 2))
pMin = Val(Mid$(Digital.Caption, 3, 2))
pSec = Val(Right$(Digital.Caption, 2))
If pHour = 1 Or pHour = 13 Then HourP.X2 = 1920: HourP.Y2 = 1080
If pHour = 2 Or pHour = 14 Then HourP.X2 = 2280: HourP.Y2 = 1320
If pHour = 3 Or pHour = 15 Then HourP.X2 = 2500: HourP.Y2 = 1740
If pHour = 4 Or pHour = 16 Then HourP.X2 = 2280: HourP.Y2 = 2040
If pHour = 5 Or pHour = 17 Then HourP.X2 = 1920: HourP.Y2 = 2520
If pHour = 6 Or pHour = 18 Then HourP.X2 = 1800: HourP.Y2 = 2620
If pHour = 7 Or pHour = 19 Then HourP.X2 = 1200: HourP.Y2 = 2400
If pHour = 8 Or pHour = 20 Then HourP.X2 = 960: HourP.Y2 = 2040
If pHour = 9 Or pHour = 21 Then HourP.X2 = 880: HourP.Y2 = 1740
If pHour = 10 Or pHour = 22 Then HourP.X2 = 960: HourP.Y2 = 1440
If pHour = 11 Or pHour = 23 Then HourP.X2 = 1320: HourP.Y2 = 1200
If pHour = 12 Or pHour = 0 Then HourP.X2 = 1620: HourP.Y2 = 860
If pMin = 5 Then MinP.X2 = 2400: MinP.Y2 = 600
If pMin = 10 Then MinP.X2 = 2760: MinP.Y2 = 1200
If pMin = 15 Then MinP.X2 = 3000: MinP.Y2 = 1740
If pMin = 20 Then MinP.X2 = 2760: MinP.Y2 = 2400
If pMin = 25 Then MinP.X2 = 2160: MinP.Y2 = 2880
If pMin = 30 Then MinP.X2 = 1620: MinP.Y2 = 3120
If pMin = 35 Then MinP.X2 = 960: MinP.Y2 = 2760
If pMin = 40 Then MinP.X2 = 600: MinP.Y2 = 2280
If pMin = 45 Then MinP.X2 = 240: MinP.Y2 = 1740
If pMin = 50 Then MinP.X2 = 360: MinP.Y2 = 1200
If pMin = 55 Then MinP.X2 = 840: MinP.Y2 = 600
If pMin = 0 Then MinP.X2 = 1712: MinP.Y2 = 360
If pSec = 0 Then SecP.X2 = 1620: SecP.Y2 = 360
If pSec = 1 Then SecP.X2 = 1712: SecP.Y2 = 452
If pSec = 2 Then SecP.X2 = 1804: SecP.Y2 = 544
If pSec = 3 Then SecP.X2 = 1896: SecP.Y2 = 636
If pSec = 4 Then SecP.X2 = 1988: SecP.Y2 = 728
If pSec = 5 Then SecP.X2 = 2080: SecP.Y2 = 820
If pSec = 6 Then SecP.X2 = 2172: SecP.Y2 = 912
If pSec = 7 Then SecP.X2 = 2264: SecP.Y2 = 1004
If pSec = 8 Then SecP.X2 = 2356: SecP.Y2 = 1096
If pSec = 9 Then SecP.X2 = 2448: SecP.Y2 = 1188
If pSec = 10 Then SecP.X2 = 2540: SecP.Y2 = 1280
If pSec = 11 Then SecP.X2 = 2632: SecP.Y2 = 1372
If pSec = 12 Then SecP.X2 = 2724: SecP.Y2 = 1464
If pSec = 13 Then SecP.X2 = 2816: SecP.Y2 = 1556
If pSec = 14 Then SecP.X2 = 2908: SecP.Y2 = 1648
If pSec = 15 Then SecP.X2 = 3000: SecP.Y2 = 1740
If pSec = 30 Then SecP.X2 = 1620: SecP.Y2 = 3120
If pSec = 45 Then SecP.X2 = 240: SecP.Y2 = 1740
End Sub
