VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mathematic Blast"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   Icon            =   "TanCosSin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fr1 
      Caption         =   "Window"
      Height          =   3615
      Left            =   960
      TabIndex        =   14
      Top             =   1560
      Width           =   6135
      Begin VB.OptionButton Op1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Repeat"
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   6600
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Text            =   "30"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Text            =   "30"
      Top             =   5640
      Width           =   855
   End
   Begin VB.ComboBox C2 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Text            =   "-none-"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox C1 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Text            =   "-none-"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Text            =   "1000"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "1000"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Text            =   "300"
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Text            =   "300"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label text10 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Text11 
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   5160
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   7200
      Y1              =   5280
      Y2              =   5280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim s As Integer
Dim s1 As Integer
Dim d As Integer
Dim d1 As Integer
Dim l As Integer
Dim l1 As Integer
Dim q As Integer
Dim q1 As Integer

Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
End Sub
Private Sub Command3_Click()
Timer1.Enabled = False
Text7.Text = ""
a = 0
Op1.Left = Fr1.Left + 10
Op1.Top = Fr1.Height * (27 / 2)
Command1.Enabled = True
End Sub
Private Sub Form_Load()
ScaleMode = vbPixels
AutoRedraw = False

C1.AddItem "Tan"
C1.AddItem "Sin"
C1.AddItem "Cos"
C1.AddItem "-none-"
'''''''''''''''''''''
C2.AddItem "Tan"
C2.AddItem "Sin"
C2.AddItem "Cos"
C2.AddItem "-none-"
End Sub
Private Sub Form_resize()
Fr1.Top = 100
Fr1.Left = 70
Fr1.Height = ScaleHeight - 260
Fr1.Width = ScaleWidth - 130
Line1.Y1 = Fr1.Height + 110
Line1.Y2 = Fr1.Height + 110
Line1.X2 = Fr1.Width + 100
Command1.Left = Fr1.Width + 5
Command1.Top = Fr1.Height + 120
Command2.Left = Command1.Left
Command2.Top = Command1.Top + 40
Command3.Left = Command1.Left
Command3.Top = Command2.Top + 40
Op1.Left = Fr1.Left + 10
Op1.Top = Fr1.Height * (27 / 2)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
text10.Caption = q
Text11.Caption = q1

Text7.Text = a
On Error GoTo er:
Text8.Text = Text5.Text
Text9.Text = Text6.Text
l = Text8.Text
l1 = Text9.Text
er:
If a > 32766 Then
a = 0
End If
s = Text1.Text
s1 = Text2.Text
d = Text3.Text
d1 = Text4.Text
''''''''''
a = a + 1

'''''''''''''''''''''''''''''''''''''
If C1.Text = "Tan" Then
q = Math.Tan(a / l) * s + d
End If
If C1.Text = "Sin" Then
q = Math.Sin(a / l) * s + d
End If
If C1.Text = "Cos" Then
q = Math.Cos(a / l) * s + d
End If
If C1.Text = "-none-" Then
q = (a / l) * s + d
End If
''''''''''''''''''''''''''''''''''''
If C2.Text = "Tan" Then
q1 = Math.Tan(a / l1) * s1 + d1
End If
If C2.Text = "Sin" Then
q1 = Math.Sin(a / l1) * s1 + d1
End If
If C2.Text = "Cos" Then
q1 = Math.Cos(a / l1) * s1 + d1
End If
If C2.Text = "-none-" Then
q1 = (a / l1) * s1 + d1
End If
''''''''''''''''''''''''''''''''''''
Op1.Left = q
Op1.Top = 3120 - q1
''''''''''''''''''''''''''''''''''''
If Op1.Left > 6240 Or Op1.Top > 3240 Then
Op1.Visible = False
Else
Op1.Visible = True
End If
If Op1.Left < 0 Or Op1.Top < 0 Then
Op1.Visible = False
Else
Op1.Visible = True
End If

End Sub

