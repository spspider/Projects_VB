VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Копирование"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "ОК"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CC1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Копирование"
      FontSize        =   10
      MaxFileSize     =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "путь к копируемому файлу:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "путь к исходному файлу:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo er:
CC1.ShowSave
Text1.Text = CC1.FileName
er:
End Sub

Private Sub Command2_Click()
On Error GoTo er1:
CC1.ShowOpen
Text2.Text = CC1.FileName
er1:
End Sub

Private Sub Command3_Click()
On Error GoTo er2:
FileCopy Text2.Text, Text1.Text
er2:
End Sub


