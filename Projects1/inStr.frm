VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'При нажатии на кнопку:
If InStr(1, Text1.Text, Text2.Text) <> 0 Then 'Сверяется 2 текстовых поля
MsgBox "Совпадение" 'При совпадении юзер узнает об этом
Else
MsgBox "Не Совпадение" 'При не совпадении юзер все равно узнает об этом
End If
End Sub
