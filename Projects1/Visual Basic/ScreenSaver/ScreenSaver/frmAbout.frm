VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000002&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О Программе"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      Begin VB.Label Label2 
         BackColor       =   &H00D1D1D1&
         Caption         =   "   Попытка создать ""Скрин Савер"", помоему начало получилось не плохо."
         Height          =   855
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " e-mail: sankocom@yandex.ru"
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   4
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Белгородская обл.   г. Алексеевка"
         Height          =   435
         Index           =   0
         Left            =   900
         TabIndex        =   3
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D1D1D1&
         Caption         =   "Кононенко Александр [SanKo]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Top             =   60
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   60
         Picture         =   "frmAbout.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Удачи !!!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ Matrix Screen Saver ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

