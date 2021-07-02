VERSION 5.00
Object = "{EF10642A-E482-4577-A6E7-6E0D1A1FEFE9}#2.0#0"; "NCTAudioDesign2.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795.641
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   Begin VB.Timer Timer5 
      Interval        =   3000
      Left            =   2880
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   120
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTTimeDisplay2 NC2Y 
      Height          =   615
      Left            =   6840
      OleObjectBlob   =   "Raxrabotka1.frx":0000
      TabIndex        =   13
      Top             =   10080
      Width           =   2655
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTStepButton2 NY 
      Height          =   1500
      Left            =   7440
      OleObjectBlob   =   "Raxrabotka1.frx":0074
      TabIndex        =   9
      Top             =   8280
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "X "
      Height          =   2775
      Left            =   6720
      TabIndex        =   14
      Top             =   8040
      Width           =   2895
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTTimeDisplay2 NC2X 
      Height          =   615
      Left            =   9840
      OleObjectBlob   =   "Raxrabotka1.frx":00C4
      TabIndex        =   12
      Top             =   10080
      Width           =   2415
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTEqualizer2 B1 
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "Raxrabotka1.frx":0138
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8880
      Width           =   4500
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTStepButton2 NX 
      Height          =   1500
      Left            =   10320
      OleObjectBlob   =   "Raxrabotka1.frx":01F0
      TabIndex        =   10
      Top             =   8280
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Выстрел"
      Height          =   615
      Left            =   12480
      MaskColor       =   &H8000000B&
      TabIndex        =   8
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CheckBox Ch1 
      BackColor       =   &H8000000B&
      Height          =   255
      Left            =   13560
      TabIndex        =   6
      Top             =   9720
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   13800
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2520
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   13800
      TabIndex        =   4
      Top             =   1200
      Width           =   1230
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTTimeDisplay2 Ti1 
      Height          =   495
      Left            =   12480
      OleObjectBlob   =   "Raxrabotka1.frx":0240
      TabIndex        =   3
      Top             =   9120
      Width           =   1455
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTVolumeButtonA2 NK1 
      Height          =   1500
      Left            =   12480
      OleObjectBlob   =   "Raxrabotka1.frx":02B4
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2160
      Top             =   120
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTCMU2 NC1 
      CausesValidation=   0   'False
      Height          =   735
      Left            =   14300
      OleObjectBlob   =   "Raxrabotka1.frx":031C
      TabIndex        =   1
      Top             =   4560
      Width           =   735
   End
   Begin NCTAUDIODESIGN2LibCtl.NCTAmpBar2 N1 
      CausesValidation=   0   'False
      Height          =   5295
      Left            =   14560
      OleObjectBlob   =   "Raxrabotka1.frx":036E
      TabIndex        =   0
      Top             =   5400
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Y"
      Height          =   2775
      Left            =   9720
      TabIndex        =   15
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Главная панель"
      Height          =   3735
      Left            =   6240
      TabIndex        =   16
      Top             =   7080
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "Купить"
         Height          =   495
         Left            =   4320
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "готовность:"
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      Caption         =   "Мощность"
      Height          =   6735
      Left            =   14160
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000B&
      Caption         =   "Боезапас"
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   8640
      Width           =   4815
   End
   Begin VB.Shape S1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      FillColor       =   &H0080C0FF&
      Height          =   495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape p1 
      Height          =   375
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   10027.25
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   15
      X2              =   15240
      Y1              =   4359.673
      Y2              =   4359.673
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "готовность:"
      Height          =   255
      Left            =   12600
      TabIndex        =   7
      Top             =   9720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v As Long

Private Sub Command1_Click()
v = NK1.Value / 12
B1.Band1 = B1.Band1 + v

If v > (82 - B1.Band1) Then
B1.Band2 = B1.Band2 + (v - (82 - B1.Band1)) ' Первый
B1.Band1 = 81
End If

If v > (82 - B1.Band2) Then
B1.Band3 = B1.Band3 + (v - (82 - B1.Band2)) ' Второй
B1.Band2 = 81
End If

If v > (82 - B1.Band3) Then
B1.Band4 = B1.Band4 + (v - (82 - B1.Band3))  'Третий
B1.Band3 = 81
End If

If v > (82 - B1.Band4) Then
B1.Band5 = B1.Band5 + (v - (82 - B1.Band4))  ' Четвёртый
B1.Band4 = 81
End If

If v > (82 - B1.Band5) Then
B1.Band6 = B1.Band6 + (v - (82 - B1.Band5))  ' пятый
B1.Band5 = 81
End If

If v > (82 - B1.Band6) Then
B1.Band7 = B1.Band7 + (v - (82 - B1.Band6))  ' Шестой
B1.Band6 = 81
End If

If v > (82 - B1.Band7) Then
B1.Band8 = B1.Band8 + (v - (82 - B1.Band7))  ' Седьмой
B1.Band7 = 81
End If

If v > (82 - B1.Band8) Then
B1.Band9 = B1.Band9 + (v - (82 - B1.Band8))  ' Восьмой
B1.Band8 = 81
End If

If v > (82 - B1.Band9) Then
B1.Band10 = B1.Band10 + (v - (82 - B1.Band9))  ' Девятый
B1.Band9 = 81
End If

If v > (82 - B1.Band10) Then
B1.Band11 = B1.Band11 + (v - (82 - B1.Band10))  ' Десятый
B1.Band10 = 81
End If

If v > (82 - B1.Band11) Then
B1.Band12 = B1.Band12 + (v - (82 - B1.Band11))  ' Одинадцатый
B1.Band11 = 81
End If

If v > (82 - B1.Band12) Then
B1.Band12 = 82   ' Двенадцатый, последний
End If

End Sub

Private Sub Command2_Click()
Form2.Show wbModal
End Sub

Private Sub Timer1_Timer()
N1.Value = NK1.Value / 10
NC1.Value = N1.Value
Ti1.Minutes = NK1.Value * 144 / 100
Text2.Text = N1.Value
Line1.Y1 = NX.Value * 10
Line1.Y2 = NX.Value * 10
Line2.X1 = NY.Value * 10
Line2.X2 = NY.Value * 10
NC2X.Seconds = NX.Value * 216 / 5
NC2Y.Seconds = NY.Value * 27
p1.Left = Line2.X1 - 160
p1.Top = Line1.Y1 - 160
If S1.Left > 11000 Then
Timer4.Enabled = True
End If
If S1.Left < 0 Then
Timer3.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
Text3.Text = NC1.Value
If Text2.Text = NC1.Value Then
Ch1.Value = 1
Else
Ch1.Value = 0
End If
End Sub
Private Sub Timer3_Timer()
S1.Left = S1.Left + 20
Timer4.Enabled = False
End Sub
Private Sub Timer4_Timer()
S1.Left = S1.Left - 20
Timer3.Enabled = False
End Sub
Private Sub Timer5_Timer()
S1.Visible = True
Timer3.Enabled = True
End Sub
