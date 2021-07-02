VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":000D
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "”далить запись"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ƒобавить запись"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'ѕри нажатии на кнопку:
List1.RemoveItem (List1.ListIndex) '” объекта ListBox есть свойство RemoveItem - оно нужно дл€ удалени€ записи, но только по индексу (у самой первой записи индекс = 0, у второй 1 - итак до упора)
End Sub

Private Sub Command2_Click()
'—юда можно втыкнуть On Error Resume Next, т.к. если запись не выбрана, то произойдет ошибка
List1.AddItem Text2.Text 'ƒобавл€ем запись, здесь есть свойство AppItem оно нужно дл€ добавлени€ записи, потом идет пробел, а за ним любой текст (в этом примере будет добавл€тьс€ запись с текстом введенным в текстовое поле")
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text '“екстовое поле 1 равно выбранной записи
End Sub
