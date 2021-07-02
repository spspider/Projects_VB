VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False ' кнопка не работает
Command1.Caption = "Готово!"
Timer1.Enabled = False 'отключаем таймер
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Timer1_Timer()
i = Int(9 * Rnd)
Command1.Visible = True
Command1.Top = 200 + 600 * (i \ 3)
Command1.Left = 200 + 900 * (i Mod 3)
End Sub
Private Sub Form_Load()
Timer1.Interval = 300
End Sub
