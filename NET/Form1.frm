VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   2295
   ClientTop       =   1455
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14205
   Begin VB.CommandButton Command12 
      Caption         =   "CNL"
      Height          =   495
      Left            =   10320
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "D"
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "U"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "R"
      Height          =   255
      Left            =   9840
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "L"
      Height          =   255
      Left            =   9360
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "RSlovo"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save sl"
      Height          =   495
      Left            =   11400
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tab"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Record"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Go"
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7575
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   13815
      ExtentX         =   24368
      ExtentY         =   13361
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "http://win.mail.ru/cgi-bin/signup"
      Top             =   960
      Width           =   13815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Navigate"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 'событие отпускания клавиши
Dim Sl As String
Dim Ts As String
Dim LastS As String
Dim AllS As String
Dim Sl1 As Integer
Dim i As Integer
Private Sub Command1_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command2_Click()
Command2.BackColor = RGB(255, 100, 100)
'For i = 1 To 8
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Call keybd_event(vbKeyTab, 0, KEYEVENTF_KEYUP, 0)
'Next i
'Rslovo
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Rslovo
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Sl = "11"
'slovo
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Call keybd_event(vbKeyDown, 0, 0, 0)
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Sl = "1998"
'slovo
'Call keybd_event(vbKeyTab, 0, 0, 0)
'Rslovo
'savesl
End Sub

Private Sub Command3_Click()
Command3.BackColor = RGB(255, 100, 100)


End Sub

Private Sub Command4_Click()
Call keybd_event(vbKeyTab, 0, 0, 0)
Call keybd_event(vbKeyTab, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub WebBrowser1_DownloadBegin()
Command2.BackColor = RGB(100, 100, 100)
End Sub

Function slovo() ' Sl - вводимое слово
For i = 1 To Len(Sl)
Sl1 = Asc(Mid(Sl, i, 1))
    If Sl1 >= 97 Then ' upper register
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))) - 32, 1, 1, 1)
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))) - 32, 0, KEYEVENTF_KEYUP, 0)
    End If
    If Sl1 <= 90 Then ' lower register
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))), 1, 1, 1)
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))), 0, KEYEVENTF_KEYUP, 0)
    End If
    If Sl1 = 92 Then ' \
    Call keybd_event(156, 0, 0, 0)
    End If
    If Sl1 = 58 Then ' :
    Call keybd_event(122, 0, 0, 0)
    End If
    If Sl1 = 46 Then ' .
    Call keybd_event(110, 0, 0, 0)
    End If
Next i
End Function
Function Rslovo() ' Sl - генерируемое слово
Dim rSl As String
Sl = ""
For i = 1 To 10
rSl = Fix(Sin(Rnd / 2) * 70 + 97)
If rSl >= 122 Then rSl = 122
If rSl <= 97 Then rSl = 97
Sl = Sl & Chr(rSl)
Next i
For i = 1 To Len(Sl)
Sl1 = Asc(Mid(Sl, i, 1))
    If Sl1 >= 97 Then ' upper register
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))) - 32, 1, 1, 1)
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))) - 32, 0, KEYEVENTF_KEYUP, 0)
    End If
    If Sl1 <= 90 Then ' lower register
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))), 1, 1, 1)
    Call keybd_event((Val(Asc(Mid(Sl, i, 1)))), 0, KEYEVENTF_KEYUP, 0)
    End If
Next i
End Function

Function savesl()
Dim MyFile 'Объявляем переменную для свободного файла
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("Data\Bin.txt") For Input As #MyFile
For i = 1 To 2
Line Input #MyFile, Ts 'Читаем файл TEST.TXT построчно
If i = 1 Then
LastS = Ts
List1.AddItem (LastS) 'Если пятая строка, то запоминаем ее в переменную S
End If
If i = 2 Then
AllS = Ts
List1.AddItem (AllS)
End If
Next i
'for i
Close #MyFile 'Закрываем файл
Open ("Data\Bin.txt") For Output As #MyFile
List1.ListIndex = 0
Print #MyFile, List1.Text
List1.ListIndex = 1
Print #MyFile, (List1.ListCount - 2)
Print #MyFile, Sl
Close #MyFile 'Закрываем файл
End Function

