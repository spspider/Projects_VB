VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   2745
   ClientTop       =   3930
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   11520
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Text            =   "     d0=12#d1=12#d2=1"
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5640
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Str, Fstr1, NStr, CodeS, CodeS_1, CodeS_2, Str1 As String
Dim CodeMass(3), Fstr(3) As String
Dim InStrC, x, x1 As Integer
Private Sub Combo1_Change()
Text2.Text = Combo1.Index
End Sub

Private Sub Command1_Click()

Dim i As Integer
i = 0
'MSComm1.DTREnable = True
'MSComm1.RTSEnable = True
For i = 0 To 500
MSComm1.Output = "1"
Next i
End Sub

Private Sub Command2_Click()
Str = ""
Str1 = ""
Text1.Text = ""
List1.Clear
End Sub


Private Sub Command3_Click()
' Fire Rx Event Every Two Bytes
MSComm1.RThreshold = 1
Text1.MaxLength = 900
' When Inputting Data, Input 2 Bytes at a time
MSComm1.InputLen = 3
MSComm1.Settings = "9600,N,8,1"
' Disable DTR
MSComm1.DTREnable = False
MSComm1.RTSEnable = True

' Open COM1
MSComm1.CommPort = (Combo1.ListIndex + 1)
'MSComm1.CommPort = 5
MSComm1.PortOpen = True
End Sub
Function CodeS1()
Fstr(0) = "#"
If (InStr(Str, Fstr(0)) <> 0) Then 'поиск n в строке
List1.AddItem (Str)
Str = ""
Else

End If
End Function

Function CodeS2()
Fstr(0) = "d1"
Fstr(1) = "d2"
Fstr(2) = "d0"
Fstr(3) = "n"

CodeMass(0) = Text2.Text
CodeMass(1) = Text3.Text
CodeMass(2) = Text4.Text

NStr = "#"
Str = Text1.Text
x = 0
For x = 0 To 3
'For x1 = 0 To 3 'обнуление всего диапазона массива
'CodeMass(x1) = 0
'Next x1
If (InStr(Str, Fstr(x)) <> 0) And (InStr(Str, NStr) <> 0) Then
Str = Mid(Str, (InStr(Str, Fstr(x))), 255)
CodeMass(x) = Val(Mid(Str, InStr(Str, Fstr(x)) + 3, InStr(Str, NStr) - InStr(Str, Fstr(x)) + 255)) 'считает длину символов с первой строки
Str = Mid(Str, (InStr(Str, NStr)) + 1, 255)
Else
Str1 = Str
End If
Next x
If (InStr(Str, Fstr(3)) <> 0) Then 'поиск n в строке
Str = Mid(Str, (InStr(Str, Fstr(3))) + 1, 255)
Command3.BackColor = RGB(200, 10, 10)
Else
Command3.BackColor = &H8000000F
End If

Text2.Text = CodeMass(0)
Text3.Text = CodeMass(1)
Text4.Text = CodeMass(2)

Text1.Text = Str
Text6.Text = Str1
End Function

Private Sub Command4_Click()
Fstr(0) = "d0"
Fstr(1) = "d1"
Fstr(2) = "d2"
CodeMass(0) = Text2.Text
CodeMass(1) = Text3.Text
CodeMass(2) = Text4.Text

NStr = "#"
Str = Text5.Text
x = 0
For x = 0 To 3
CodeMass(x) = 0
If (InStr(Str, Fstr(x)) <> 0) And (InStr(Str, NStr) <> 0) Then
Str1 = Mid(Str, (InStr(Str, Fstr(x))), 255)
CodeMass(x) = Val(Mid(Str1, InStr(Str1, Fstr(x)) + 3, InStr(Str1, NStr) - InStr(Str1, Fstr(x)) + 255)) 'считает длину символов с первой строки
Str1 = Mid(Str1, (InStr(Str1, NStr)) + 1, 255)
Str = Mid(Str, (InStr(Str, NStr)) + 1, 255)
Else
Str1 = Str
End If


Next x
Text2.Text = CodeMass(0)
Text3.Text = CodeMass(1)
Text4.Text = CodeMass(2)

Text5.Text = Str
Text6.Text = Str1
End Sub



Private Sub Form_Load()
Combo1.AddItem ("com1")
Combo1.AddItem ("com2")
Combo1.AddItem ("com3")
Combo1.AddItem ("com4")
Combo1.AddItem ("com5")
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False

End Sub

Private Sub MSComm1_OnComm()
Dim sData As String ' Holds our incoming data
If MSComm1.CommEvent = comEvReceive Then
sData = MSComm1.Input ' Get data (2 bytes)
lWord = (sData)
Str = Str & CStr(lWord)
'Text1.Text = Str
Text1.Text = Str
CodeS1

End If
End Sub

Private Sub Text4_Change()
'Str = Mid(Str, (InStr(Str, NStr)) + 1, 255)
End Sub

Private Sub Timer1_Timer()

End Sub

