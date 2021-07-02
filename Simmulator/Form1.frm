VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   "Simmulator - Sp(c)"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "453456"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808000&
      Caption         =   "Center Screen"
      Height          =   500
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   975
      Left            =   11520
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00404000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "1"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "Draw"
      Height          =   500
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Sector development"
      Height          =   500
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DLine1(20), DLine2(20) As String
Dim x1, y1, x2, y2, Xm, Ym As Integer
Dim q1 As Boolean
Dim CentX, CentY As Currency

Private Sub Combo1_Validate(Cancel As Boolean)
'draw
End Sub

Private Sub Command1_Click()
Form2.Show
Form1.AutoRedraw = True
End Sub
Private Sub Command2_Click()
'On Error GoTo er:
Form2.List1.ListIndex = 0
Form2.List2.ListIndex = 0
For i = 0 To Form2.List1.ListCount - 2
DLine1(i) = (Val(Mid(Form2.List1.Text, 1, 2)) + Val(Mid(Form2.List1.Text, 3, 2)) / 60 + Val(Mid(Form2.List1.Text, 4, 2)) / 3600)
DLine2(i) = (Val(Mid(Form2.List2.Text, 1, 3)) + Val(Mid(Form2.List2.Text, 4, 2)) / 60 + Val(Mid(Form2.List2.Text, 5, 2)) / 3600)
Form2.List1.ListIndex = Form2.List1.ListIndex + 1
Form2.List2.ListIndex = Form2.List2.ListIndex + 1
Next i
'Form2.List1.ListIndex = 0
'Form2.List2.ListIndex = 0
'er:
draw
End Sub
Function draw()
Form1.Cls
Dim i As Long
For i = 0 To Form2.List1.ListCount - 2
x1 = (DLine1(i) - CentX) * Val(Combo1.Text) + Form1.Width / 2 '/ 40 '/ HScroll1.Value 444600
y1 = (DLine2(i) - CentY) * Val(Combo1.Text) + Form1.Height / 2 '/ 40 ' / HScroll1.Value 0300900
If i = Form2.List1.ListCount - 2 Then
x2 = (DLine1(0) - CentX) * Val(Combo1.Text) + Form1.Width / 2 ' / 40 ' / HScroll1.Value
y2 = (DLine2(0) - CentY) * Val(Combo1.Text) + Form1.Height / 2 ' / HScroll1.Value
Else
x2 = (DLine1(i + 1) - CentX) * Val(Combo1.Text) + Form1.Width / 2  ' / 40 ' / HScroll1.Value
y2 = (DLine2(i + 1) - CentY) * Val(Combo1.Text) + Form1.Height / 2 ' / HScroll1.Value
End If
Line (x1, y1)-(x2, y2), RGB(0, 100, 255)
Next i

For i = 1 To 10
Line (i * 2000, 0)-(i * 2000, Form1.Height), RGB(0, 100, 100)
Line (0, i * 2000)-(Form1.Width, i * 2000), RGB(0, 100, 100)
Next i

End Function
Private Sub Command3_Click()
Form1.MousePointer = 2
Label2.Visible = True
Label3.Visible = True
q1 = True
'Label3.Visible = True
End Sub

Private Sub Command4_Click()
Text2.Text = Val(Mid(Text1.Text, 1, 2)) + Val(Mid(Text1.Text, 3, 2)) / 60 + Val(Mid(Text1.Text, 5, 2)) / 3660
End Sub

Private Sub Form_Click()
If q1 = True Then
Form1.MousePointer = 0
Label2.Visible = False
Label3.Visible = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub
Private Sub Form_Load()
CentX = 45
CentY = 30
For i = 1 To 20
Combo1.AddItem (i * 100)
Next i
End Sub

Private Sub Form_Resize()
Command1.Left = Form1.Width - Command1.Width - 100
Command2.Left = Form1.Width - Command2.Width - 100
Command3.Left = Form1.Width - Command3.Width - 100
Frame1.Left = Form1.Width - Frame1.Width - 500
End Sub

Private Sub Timer1_Timer()
'draw
Label1.Caption = Time
Label2.Caption = "N" & Xm ' / 200
Label3.Caption = "E" & Ym ' / 200
Label2.Left = Xm + 100
Label2.Top = Ym
Label3.Left = Xm + 100
Label3.Top = Ym + Label3.Height
End Sub

