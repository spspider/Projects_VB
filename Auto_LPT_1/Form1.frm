VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LPT"
   ClientHeight    =   4560
   ClientLeft      =   2760
   ClientTop       =   2490
   ClientWidth     =   12030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   9360
      TabIndex        =   45
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Порты"
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   11895
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   960
         TabIndex        =   50
         Text            =   "00:00"
         Top             =   720
         Width           =   735
      End
      Begin VB.ListBox List12 
         Height          =   2205
         Left            =   960
         TabIndex        =   49
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   9960
         TabIndex        =   48
         Text            =   "Text9"
         Top             =   600
         Width           =   975
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   10920
         TabIndex        =   47
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox List11 
         Height          =   2400
         Left            =   9960
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DELETE"
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DELETE"
         Height          =   255
         Left            =   6600
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3240
         Width           =   3375
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LOAD"
         Height          =   255
         Left            =   8280
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SAVE"
         Height          =   255
         Left            =   6600
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OFF"
         Height          =   255
         Left            =   5760
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Text            =   "00:00"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADD"
         Height          =   255
         Left            =   240
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Text            =   "Text7"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5760
         TabIndex        =   36
         Text            =   "00:00"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox List10 
         Height          =   2400
         ItemData        =   "Form1.frx":8653
         Left            =   5760
         List            =   "Form1.frx":8655
         TabIndex        =   35
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox List9 
         Height          =   2400
         ItemData        =   "Form1.frx":8657
         Left            =   8880
         List            =   "Form1.frx":8659
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox List8 
         Height          =   2400
         ItemData        =   "Form1.frx":865B
         Left            =   7680
         List            =   "Form1.frx":865D
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.ListBox List7 
         Height          =   2400
         ItemData        =   "Form1.frx":865F
         Left            =   6600
         List            =   "Form1.frx":8661
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Height          =   2400
         ItemData        =   "Form1.frx":8663
         Left            =   4920
         List            =   "Form1.frx":8665
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox List5 
         Height          =   2400
         ItemData        =   "Form1.frx":8667
         Left            =   3720
         List            =   "Form1.frx":8669
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADD"
         Height          =   255
         Left            =   3720
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3240
         Width           =   2895
      End
      Begin VB.ListBox List4 
         Height          =   2205
         ItemData        =   "Form1.frx":866B
         Left            =   240
         List            =   "Form1.frx":866D
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.ListBox List3 
         Height          =   2205
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   1680
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "0000"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "00000000"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   28
         Text            =   "00000001"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Form1.frx":866F
         Left            =   7680
         List            =   "Form1.frx":8671
         TabIndex        =   27
         Text            =   " 888"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   24
         Text            =   "00000000"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         MaxLength       =   5
         TabIndex        =   23
         Text            =   "00:00"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Form1.frx":8673
         Left            =   3720
         List            =   "Form1.frx":8675
         TabIndex        =   21
         Text            =   "888"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control 890"
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data 888"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THEN"
         Height          =   255
         Left            =   7800
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AS"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IN"
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IF"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "00000000"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "00000000"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "00000"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0000"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "off"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "off"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "off"
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "on"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "on"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "on"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n As Integer
Dim QQQ As Integer
Dim i As Integer
Dim FormVis As Integer

Dim BudKon As Currency
Dim BudNach As Currency
Dim BudVrem As Currency

Dim BudVrem1 As Currency
Dim BudNach1 As Currency

Dim CodInp1 As String
Dim CodInp2 As String
Dim CodSrav1 As String
Dim CodSrav2 As String
Dim CodItog1 As String
Dim CodItog2 As String

Private Declare Function Inp Lib "inpout32\inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32\inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
' иконка в трее

' сделать так, что бы была проверка 00001 и если этой еденицы нет, то он её дописывал.
Dim nid As NOTIFYICONDATA
Private Sub Command10_Click() ' Print
Dim MyFile 'Объявляем переменную для свободного файла
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("Data\Bin.txt") For Output As #MyFile
Print #MyFile, List4.ListCount
Print #MyFile, List6.ListCount
For i = 0 To List1.ListCount - 1
List4.ListIndex = i
Print #MyFile, List4.Text
Next i
For i = 0 To List1.ListCount - 1
List1.ListIndex = i
Print #MyFile, List1.Text
Next i
For i = 0 To List1.ListCount - 1
List3.ListIndex = i
Print #MyFile, List3.Text
Next i

For i = 0 To List1.ListCount - 1
List12.ListIndex = i
Print #MyFile, List12.Text
Next i

For i = 0 To List6.ListCount - 1
List5.ListIndex = i
Print #MyFile, List5.Text
Next i
For i = 0 To List6.ListCount - 1
List6.ListIndex = i
Print #MyFile, List6.Text
Next i
For i = 0 To List6.ListCount - 1
List7.ListIndex = i
Print #MyFile, List7.Text
Next i
For i = 0 To List6.ListCount - 1
List8.ListIndex = i
Print #MyFile, List8.Text
Next i
For i = 0 To List6.ListCount - 1
List9.ListIndex = i
Print #MyFile, List9.Text
Next i
For i = 0 To List6.ListCount - 1
List10.ListIndex = i
Print #MyFile, List10.Text
Next i
Close #MyFile 'Закрываем файл
End Sub

Function LoadBin()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
List12.Clear
Dim Ns As Integer
Dim MyFile 'Объявляем переменную для свободного файла
Dim i As Integer 'Переменная для цикла
Dim i1 As Integer 'Переменная для цикла
Dim tS As String 'Переменная для считывания строк
Dim tS1 As String 'Переменная для считывания строк
Dim S As String 'Переменная для хранения окончательных данных
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("Data\Bin.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
For i = 1 To 2
Line Input #MyFile, tS 'Читаем файл TEST.TXT построчно
List2.AddItem (tS)
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
List2.ListIndex = 0
For i = 0 To (Val(List2.Text) * 1 + 1)
Line Input #MyFile, tS1
If i >= (Val(List2.Text) * 0 + 2) Then
List4.AddItem (tS1)
End If
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
List2.ListIndex = 0
For i = 0 To Val(List2.Text) * 2 + 1
Line Input #MyFile, tS1
If i >= (Val(List2.Text) * 1 + 2) Then
List1.AddItem (tS1)
End If
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
For i = 0 To Val(List2.Text) * 3 + 1
Line Input #MyFile, tS1
If i >= (Val(List2.Text) * 2 + 2) Then
List3.AddItem (tS1)
End If
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
For i = 0 To Val(List2.Text) * 4 + 1
Line Input #MyFile, tS1
If i >= (Val(List2.Text) * 3 + 2) Then
List12.AddItem (tS1)
End If
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + List2.Text)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 Then
List5.AddItem (tS1)
End If
Next i
Close #MyFile

Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + Val(List2.Text) * 2)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 + Val(List2.Text) Then
List6.AddItem (tS1)
End If
Next i
Close #MyFile
Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + Val(List2.Text) * 3)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 + Val(List2.Text) * 2 Then
List7.AddItem (tS1)
End If
Next i
Close #MyFile
Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + Val(List2.Text) * 4)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 + Val(List2.Text) * 3 Then
List8.AddItem (tS1)
End If
Next i
Close #MyFile
Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + Val(List2.Text) * 5)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 + Val(List2.Text) * 4 Then
List9.AddItem (tS1)
End If
Next i
Close #MyFile
Open ("Data\Bin.txt") For Input As #MyFile
List2.ListIndex = 0
Ns = Val(List2.Text)
List2.ListIndex = 1
For i = 0 To (Ns * 4 + 1 + Val(List2.Text) * 6)
Line Input #MyFile, tS1
If i >= Ns * 4 + 2 + Val(List2.Text) * 5 Then ' "4" - число столбцов первой программы
List10.AddItem (tS1)
End If
Next i
Close #MyFile
End Function
Private Sub Command11_Click() 'Read
LoadBin
End Sub

Private Sub Command13_Click()
On Error GoTo er:
List5.RemoveItem (List5.ListIndex)
List6.RemoveItem (List6.ListIndex)
List7.RemoveItem (List7.ListIndex)
List8.RemoveItem (List8.ListIndex)
List9.RemoveItem (List9.ListIndex)
List10.RemoveItem (List10.ListIndex)
er:
End Sub

Private Sub Command14_Click()
On Error GoTo er2:
List4.RemoveItem (List4.ListIndex)
List1.RemoveItem (List1.ListIndex)
List3.RemoveItem (List3.ListIndex)
List12.RemoveItem (List12.ListIndex)
er2:
End Sub

Private Sub Command2_Click()
nid.hIcon = Form1.Icon
nid.szTip = "New Icon" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub Command3_Click()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Command1_Click()
List1.AddItem (Text8.Text)
List3.AddItem (Text10.Text)
List4.AddItem (Text15.Text)
List12.AddItem (Text16.Text)
End Sub

Private Sub Command5_Click()
Out 888, GetDecimalNumber(Text4.Text)
End Sub

Private Sub Command6_Click()
Out 889, GetDecimalNumber(Text5.Text)
End Sub

Private Sub Command7_Click()
Out 890, GetDecimalNumber(Text6.Text)
End Sub

Private Sub Command8_Click()
List5.AddItem (Combo1.Text)
List6.AddItem (Text11.Text)
List7.AddItem (Text12.Text)
List8.AddItem (Combo2.Text)
List9.AddItem (Text13.Text)
List10.AddItem (Text14.Text)
End Sub

Private Sub Command9_Click()
Text11.Text = "off"
End Sub

Private Sub Form_Load()
LoadBin
File1.Path = "Sound"

Combo1.AddItem ("888")
Combo1.AddItem ("889")
Combo1.AddItem ("890")
Combo2.AddItem ("888")
Combo2.AddItem ("890")
' Добавить иконку формы в Traybar
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'При наведении курсора на Иконку, выдвинется текст: "И не забудь зайти на VBStreets.Narod.RU":
nid.szTip = "Control" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long

Dim sFilter As String

msg = X / Screen.TwipsPerPixelX

Select Case msg
Case WM_LBUTTONUP
Form1.Show
Case WM_RBUTTONUP
Shell_NotifyIcon NIM_DELETE, nid
Shell ("taskkill.exe /f /IM " + App.EXEName + ".exe")
End Select
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConvertToBin(ByVal bByte As Byte) As String
 Dim i As Integer
 For i = 0 To QQQ
  If bByte And 2 ^ i Then
   ConvertToBin = 1 & ConvertToBin
  Else
   ConvertToBin = 0 & ConvertToBin
  End If
 Next i
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDecimalNumber(ByVal strBinaruNumber As String) As Long
 'On Error Resume Next
 Dim strRead As String
 Dim lngResult As Long
 Dim lngCount As Long
 Dim i As Long
 
 lngCount = Len(strBinaruNumber): i = 1
 strRead = Mid(strBinaruNumber, i, 1)
 Do While Not strRead = vbNullString
  lngResult = lngResult + (CLng(strRead) * (2 ^ (lngCount - i)))
  i = i + 1
  strRead = Mid(strBinaruNumber, i, 1)
 Loop
 
 GetDecimalNumber = lngResult
End Function
''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
Form1.Hide
'Shell_NotifyIcon NIM_DELETE, nid
Cancel = True
End Sub

Private Sub Text15_Change()
Text16.Text = Text15.Text
End Sub

Private Sub Timer1_Timer()
QQQ = 7
Text1.Text = ConvertToBin(Inp(888))
QQQ = 4
Text2.Text = ConvertToBin(Inp(889))
QQQ = 3
Text3.Text = ConvertToBin(Inp(890))

If Len(Time) = 8 Then
Text7.Text = Mid(Time, Len(Time) - 7, 5)
Else
Text7.Text = "0" + Mid(Time, Len(Time) - 6, 4)
End If

End Sub

Private Sub Timer2_Timer()

For i = 0 To List5.ListCount
On Error GoTo er:
List5.ListIndex = i
List6.ListIndex = i
List7.ListIndex = i
List8.ListIndex = i
List9.ListIndex = i
List10.ListIndex = i
er:

If List6.Text = "off" Then
'If List6.Text = Text7.Text Then
        If List5.Text = "888" Then
        If List7.Text = Text1.Text Then
        Out List8.Text, List9.Text
        End If
    End If
    If List5.Text = "889" Then
        If List7.Text = Text2.Text Then
        Out List8.Text, List9.Text
        End If
    End If
    If List5.Text = "890" Then
        If List7.Text = Text3.Text Then
        Out List8.Text, List9.Text
        End If
    End If
'MsgBox "Work"
'End If
Else
BudVrem = Val(Mid(Text7.Text, 1, 2)) + Val(Mid(Text7.Text, 4, 2)) / 60
BudNach = Val(Mid(List6.Text, 1, 2)) + Val(Mid(List6.Text, 4, 2)) / 60
BudKon = Val(Mid(List10.Text, 1, 2)) + Val(Mid(List10.Text, 4, 2)) / 60

If BudVrem <= BudKon Then ' конечное время будильника больше чем текущее время
    If BudNach <= BudVrem Then ' начало включения будиильника меньше чем текущее время.
            If List5.Text = "888" Then
            If List7.Text = Text1.Text Then
            Out List8.Text, List9.Text
            End If
            End If
            If List5.Text = "890" Then
            If List7.Text = Text3.Text Then
            Out List8.Text, List9.Text
            End If
            End If
    End If
Else
            If List5.Text = "888" Then
            Out List8.Text, 0
            End If
            If List5.Text = "890" Then
            Out List8.Text, 0
            End If



End If
End If
Next i

'''''''''''''''''''''''''''''''''''''''''''''''
For i = 0 To List4.ListCount - 1
List1.ListIndex = i
List3.ListIndex = i
List4.ListIndex = i
List12.ListIndex = i
If (Val(Mid(List12.Text, 1, 2)) + Val(Mid(List12.Text, 4, 2)) / 60) >= (Val(Mid(Text7.Text, 1, 2)) + Val(Mid(Text7.Text, 4, 2)) / 60) Then
If (Val(Mid(List4.Text, 1, 2)) + Val(Mid(List4.Text, 4, 2)) / 60) <= (Val(Mid(Text7.Text, 1, 2)) + Val(Mid(Text7.Text, 4, 2)) / 60) Then
DEC1
'Call mciExecute("Sound\ISignIn.mp3")
End If
End If
Next i

End Sub
Function DEC1() ' выполнение заданий, добавляет значения к текущему, если "00000000" - выключает

Dim i2 As Integer
Dim i1 As Integer


CodInp1 = Text1.Text
CodSrav1 = List1.Text



If CodSrav1 = "00000000" Then
CodItog1 = CodSrav1
Else
For i2 = 1 To 8
If Mid(CodInp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "0" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "0" + Mid(CodItog1, 10 - i2, 8)
End If
End If
Next i2
End If


CodSrav1 = CodItog1
Out 888, GetDecimalNumber(CodItog1)

CodInp2 = Text3.Text
CodSrav2 = List3.Text

If CodSrav2 = "0000" Then
CodItog2 = CodSrav2
Else
For i1 = 1 To 4
If Mid(CodInp2, 5 - i1, 1) = 1 Then
If Mid(CodSrav2, 5 - i1, 1) = 1 Then
CodItog2 = Mid(CodInp2, 1, 4 - i1) + "1" + Mid(CodItog2, 6 - i1, 4)
End If
End If
If Mid(CodInp2, 5 - i1, 1) = 0 Then
If Mid(CodSrav2, 5 - i1, 1) = 1 Then
CodItog2 = Mid(CodInp2, 1, 4 - i1) + "1" + Mid(CodItog2, 6 - i1, 4)
End If
End If
If Mid(CodInp2, 5 - i1, 1) = 0 Then
If Mid(CodSrav2, 5 - i1, 1) = 0 Then
CodItog2 = Mid(CodInp2, 1, 4 - i1) + "0" + Mid(CodItog2, 6 - i1, 4)
End If
End If
If Mid(CodInp2, 5 - i1, 1) = 1 Then
If Mid(CodSrav2, 5 - i1, 1) = 0 Then
CodItog2 = Mid(CodInp2, 1, 4 - i1) + "0" + Mid(CodItog2, 6 - i1, 4)
End If
End If
Next i1
End If

Out 890, GetDecimalNumber(CodItog2)

End Function


