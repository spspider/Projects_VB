VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   Icon            =   "Linesled3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   1575
      Left            =   3480
      Max             =   100
      Min             =   1
      TabIndex        =   24
      Top             =   360
      Value           =   2
      Width           =   375
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   1575
      Left            =   2760
      Max             =   100
      Min             =   1
      TabIndex        =   23
      Top             =   360
      Value           =   2
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Text            =   "2"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2760
      TabIndex        =   21
      Text            =   "2"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "связать"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1575
      Left            =   1200
      Max             =   3000
      Min             =   -3000
      TabIndex        =   17
      Top             =   360
      Value           =   200
      Width           =   375
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1575
      Left            =   1920
      Max             =   3000
      Min             =   -3000
      TabIndex        =   16
      Top             =   360
      Value           =   200
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Text            =   "200"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "200"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "1500"
      Top             =   2640
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      Left            =   360
      Max             =   2050
      TabIndex        =   9
      Top             =   240
      Value           =   1500
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Text            =   "21000"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Text            =   "17000"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "^"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.OptionButton Op2 
      Caption         =   "Option2"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   4440
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   3360
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   2640
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   1800
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Linesled3.frx":5A4A
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer

Dim op1L As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Function ref()
Form1.Cls
For a = 0 To Text3.Text
op1T = Sin(a / Text6.Text) * Tan(a / 4) * Text4.Text + Screen.Height + Text1.Text
op1L = Cos(a / Text7.Text) * Tan(a / 4) * Text5.Text + Screen.Width + Text2.Text

Line (op1L + x, op1T + x)-(op2L + x, op2T + x), RGB(0, 255, 100)

op2T = 2 * Screen.Height - op1T
op2L = 2 * Screen.Width - op1L
Next
End Function
Private Sub Command1_Click()
ref
End Sub
Private Sub Command4_Click()
Text2.Text = Text2.Text + 1000
End Sub
Private Sub Command5_Click()
Text2.Text = Text2.Text - 1000
End Sub
Private Sub Command2_Click()
Text1.Text = Text1.Text + 1000
End Sub
Private Sub Command3_Click()
Text1.Text = Text1.Text - 1000
End Sub

Private Sub Command6_Click()
Form1.Cls
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
a = a + 1
Op1.Top = Sin(a / Text6.Text) * Tan(a / 4) * Text4.Text + Screen.Height + Text1.Text
Op1.Left = Cos(a / Text7.Text) * Tan(a / 4) * Text5.Text + Screen.Width + Text2.Text
    'x = x - 1
    'If x < 1 Then
    'Form1.Cls
    'x = 255
    'End If
Line (Op1.Left + x, Op1.Top + x)-(Op2.Left + x, Op2.Top + x), RGB(0, 255, 100)
Op2.Top = 2 * Screen.Height - Op1.Top
Op2.Left = 2 * Screen.Width - Op1.Left
If a > Text3.Text Then
Timer1.Enabled = False
a = 0
End If
End Sub

Private Sub VScroll1_Change()
Text3.Text = VScroll1.Value
End Sub
Private Sub VScroll2_Change()
Text4.Text = VScroll2.Value
If Check1.Value = 1 Then
Text5.Text = VScroll2.Value
VScroll3.Value = VScroll2.Value
End If
End Sub
Private Sub VScroll3_Change()
Text5.Text = VScroll3.Value
If Check1.Value = 1 Then
Text4.Text = VScroll3.Value
VScroll2.Value = VScroll3.Value
End If
End Sub
Private Sub VScroll4_Change()
Text6.Text = VScroll4.Value
End Sub
Private Sub VScroll5_Change()
Text7.Text = VScroll5.Value
End Sub

