VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   Icon            =   "avtozagruz.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "копировать..."
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1050
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "в:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "из:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FileCopy Text1.Text, Text2.Text
End Sub

Private Sub Form_Load()
Text1.Text = App.Path + "\" + App.EXEName + ".exe"
Text2.Text = "C:\Documents and Settings\All Users\Главное меню\Программы\Автозагрузка" + "\" + App.EXEName + ".exe"
End Sub
