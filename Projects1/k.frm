VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   5085
   ClientLeft      =   5040
   ClientTop       =   3825
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5145
   Begin MSMask.MaskEdBox TiK 
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.TextBox Text16 
      Height          =   405
      Left            =   2160
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin MSMask.MaskEdBox Ti1 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3120
      Top             =   240
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2520
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1920
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Конец"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Начало"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1560
      Width           =   615
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
Private Sub Timer1_Timer()
Text1.Text = Time
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
Private Sub Timer8_Timer()
Text16.Text = Time
If Text16.Text = Ti1 Then
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Form1.Visible = False
End If
If Text16.Text = TiK Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Unload Me
End If
End Sub
