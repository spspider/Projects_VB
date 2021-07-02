VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2970
   ClientLeft      =   5040
   ClientTop       =   4620
   ClientWidth     =   5145
   Icon            =   "Диспетчер задач.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Начать"
      Height          =   1095
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   735
   End
   Begin VB.Timer Timer6 
      Interval        =   3000
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer Timer5 
      Interval        =   3000
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Опции"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3720
         Top             =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Повтор"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Продолжительность (сек.)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Задержка (мин.)"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Информация"
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   4935
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
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
Text1.Text = Text2.Text
Text3.Text = Text16.Text * 60
Timer9.Enabled = True
Text15.Text = Text4.Text
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

Private Sub Timer7_Timer()
Unload Me
End Sub
Private Sub Timer9_Timer()
Text3.Text = Text3.Text - 1

    If Text3.Text <= 0 Then
    Timer2.Enabled = True
    Text1.Text = Text1.Text - 1
        If Text1.Text <= 0 Then
        Timer2.Enabled = False
        
        Text15.Text = Text15.Text - 1
        Text1.Text = Text2.Text
        Text3.Text = Text16.Text * 60
            
            If Text15.Text = 0 Then
            Unload Me
            End If
        End If
    End If
End Sub
