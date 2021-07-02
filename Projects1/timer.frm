VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALT+F4"
   ClientHeight    =   4530
   ClientLeft      =   5835
   ClientTop       =   4215
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   4200
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   4200
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   645
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   3600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   3840
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   3360
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Начать"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Скрыть"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   4080
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   3
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   1
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Расширение"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "средний фак."
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "факториал X"
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "факториал Y"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "шаг B (Y)"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "шаг A (X)"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "ось Y"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "ось X"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ALT+F4"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slu As Integer
Dim sav As Integer
Dim w As String
Dim h As String
Dim linpm As Integer
Dim lin As Integer
Dim sluy As Integer
Dim sred As Integer
Dim plu As Integer
Dim min As Integer
Dim x As Integer
Dim y As Integer
Dim f As Integer
Dim b As Integer
Dim a As Integer
Private Declare Function SetCursorPos Lib "user32" (ByVal r As Long, ByVal r1 As Long) As Long
Private Sub Command1_Click()
Form1.Visible = False
End Sub

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
If Text1.Text = Str(strok1) Then
MsgBox "123", 64, "123"

End If
If Len(Text2.Text) >= 2 Then
Text3.SetFocus
End If
If Len(Text3.Text) >= 2 Then
Text4.SetFocus
End If
If Text2.Text >= "24" Then
Text2.Text = "24"
End If
If Text3.Text >= "59" Then
Text3.Text = "59"
End If
If Text4.Text >= "59" Then
Text4.Text = "59"
End If
End Sub
Private Sub Command2_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
End Sub
Private Sub Timer2_Timer()
Text6.Text = x
Text8.Text = y
a = a + plu
b = b + min
x = 1024 / 2 + a
y = 768 / 2 + b
c = SetCursorPos(x, y)
If x > (Screen.Width / 15) Then
plu = -1
Text6.BackColor = &HFF&
Else
Text6.BackColor = &H80000004
End If
If y > (Screen.Height / 15) Then
min = -1
Text8.BackColor = &HFF&
Else
Text8.BackColor = &H80000004
End If
If x < 0 Then
plu = 1
Text6.BackColor = 65280
Else
Text6.BackColor = &H80000004
End If
If y < 0 Then
min = 1
Text8.BackColor = 65280
Else
Text8.BackColor = &H80000004
End If
w = Screen.Width / 15
h = Screen.Height / 15
Text9.Text = lin
Text10.Text = linpm
Text11.Text = plu
Text12.Text = min
Text13.Text = sred
Text14.Text = w & " x " & h
End Sub

Private Sub Timer3_Timer()
slu = Fix(Rnd * 100)
Text5.Text = slu
If slu > sred Then
plu = 2
Else
plu = -2
End If
End Sub
Private Sub Timer4_Timer()
sluy = Fix(Rnd * 100)
Text7.Text = sluy
If sluy > sred Then
min = -2
Else
min = 2
End If
End Sub
Private Sub Timer5_Timer()
sred = (slu + sluy) / 2
End Sub

Private Sub Timer6_Timer()
lin = Fix(Rnd * 100)
If lin > sred Then
linpm = Fix(Rnd * 100)
If linpm > 50 Then
plu = 0
Else
min = 0
End If
End If
End Sub
