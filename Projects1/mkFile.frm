VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "сотворить файл"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "файл лежит в папке Data"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "текст в нём:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "имя файла:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d = FreeFile
Open "Data\" & Text1.Text & ".txt" For Append As d
Print #d, Text2.Text
Close #d
End Sub

Private Sub Form_Load()
On Error GoTo er:
MkDir "Data"
er:
End Sub
