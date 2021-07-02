VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Опции SP"
   ClientHeight    =   8505
   ClientLeft      =   17505
   ClientTop       =   1740
   ClientWidth     =   6345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   Begin VB.Frame Frame5 
      Caption         =   "Радиус расхождения"
      Height          =   975
      Left            =   360
      TabIndex        =   17
      Top             =   6600
      Width           =   5655
      Begin VB.HScrollBar HSRR 
         Height          =   375
         Left            =   240
         Max             =   1500
         TabIndex        =   18
         Top             =   360
         Value           =   1000
         Width           =   4695
      End
      Begin VB.Label Label7 
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Дополнительное увеличение"
      Height          =   1095
      Left            =   360
      TabIndex        =   14
      Top             =   5400
      Width           =   5655
      Begin VB.HScrollBar HSR 
         Height          =   375
         Left            =   240
         Max             =   5000
         Min             =   -5000
         TabIndex        =   15
         Top             =   480
         Value           =   1000
         Width           =   4695
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Скорость превращения"
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   4200
      Width           =   5655
      Begin VB.HScrollBar HSP 
         Height          =   375
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   12
         Top             =   480
         Value           =   30
         Width           =   4695
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "По умолчанию"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   840
   End
   Begin VB.Frame Frame4 
      Caption         =   "Инициализация функций Cos и Sin"
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   5655
      Begin VB.HScrollBar HSS 
         Height          =   375
         Left            =   120
         Max             =   200
         Min             =   -200
         TabIndex        =   5
         Top             =   840
         Value           =   8
         Width           =   4815
      End
      Begin VB.HScrollBar HSC 
         Height          =   375
         Left            =   120
         Max             =   200
         Min             =   -200
         TabIndex        =   4
         Top             =   360
         Value           =   8
         Width           =   4815
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Инициализация функции Tan"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      Begin VB.HScrollBar HST1 
         Height          =   375
         Left            =   120
         Max             =   200
         Min             =   -200
         TabIndex        =   3
         Top             =   240
         Value           =   50
         Width           =   4815
      End
      Begin VB.HScrollBar HST2 
         Height          =   375
         Left            =   120
         Max             =   200
         Min             =   -200
         TabIndex        =   2
         Top             =   720
         Value           =   -40
         Width           =   4815
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
HST1.Value = 50
HST2.Value = -40
HSC.Value = 8
HSS.Value = 8
HSP.Value = 30
HSR.Value = 1000
HSRR.Value = 1000
End Sub

Private Sub Form_Load()
On Error GoTo er:
HST1.Value = GetSetting("SP\SPSoft\SP\трапеция", "HST1", "HST1")
HST2.Value = GetSetting("SP\SPSoft\SP\трапеция", "HST2", "HST2")
HSC.Value = GetSetting("SP\SPSoft\SP\трапеция", "HSC", "HSC")
HSS.Value = GetSetting("SP\SPSoft\SP\трапеция", "HSS", "HSS")
HSP.Value = GetSetting("SP\SPSoft\SP\трапеция", "HSP", "HSP")
HSR.Value = GetSetting("SP\SPSoft\SP\трапеция", "HSR", "HSR")
HSRR.Value = GetSetting("SP\SPSoft\SP\трапеция", "HSRR", "HSRR")
er:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.TimerOPE.Enabled = True


SaveSetting "SP\SPSoft\SP\трапеция", "HST1", "HST1", HST1.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HST2", "HST2", HST2.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HSC", "HSC", HSC.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HSS", "HSS", HSS.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HSP", "HSP", HSP.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HSR", "HSR", HSR.Value
SaveSetting "SP\SPSoft\SP\трапеция", "HSRR", "HSRR", HSRR.Value
End Sub

Private Sub Timer1_Timer()
On Error GoTo er:
Label1.Caption = HST1.Value
Label2.Caption = HST2.Value
Label3.Caption = HSC.Value
Label4.Caption = HSS.Value
Label5.Caption = HSP.Value
Label6.Caption = HSR.Value
Label7.Caption = HSRR.Value
er:
End Sub
