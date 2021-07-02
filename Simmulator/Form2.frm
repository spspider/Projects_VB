VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Sector developer"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4905
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tS As String
Private Sub Command1_Click()
List1.AddItem (Text1.Text)
List2.AddItem (Text2.Text)
End Sub

Private Sub Command2_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Command3_Click()
Dim MyFile 'Объявляем переменную для свободного файла
Dim X As Integer
Dim S As String 'Переменная для хранения считанных данных

MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("Sector.txt") For Output As #MyFile 'Открываем файл TEST.TXT для чтения
Print #MyFile, Val(List1.ListCount)
For X = 1 To List1.ListCount
List1.ListIndex = X - 1
List2.ListIndex = X - 1
Print #MyFile, List1.Text
Print #MyFile, List2.Text
Next X
Close #MyFile 'Закрываем файл
End Sub

Private Sub Command4_Click()
Dim MyFile 'Объявляем переменную для свободного файла
Dim i As Integer 'Переменная для цикла
Dim S As String 'Переменная для хранения окончательных данных
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("Sector.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
Line Input #MyFile, tS
List1.Clear
List2.Clear
For i = 1 To Val(tS)
Line Input #MyFile, S
List1.AddItem (S)
Line Input #MyFile, S
List2.AddItem (S)
'Line Input #MyFile, tS 'Читаем файл TEST.TXT построчно
'If i >= 5 Then S = tS 'Если пятая строка, то запоминаем ее в переменную S
Next i
Close #MyFile 'Закрываем файл
End Sub
