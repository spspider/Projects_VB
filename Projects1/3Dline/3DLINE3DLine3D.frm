VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "10"
      Top             =   3960
      Width           =   735
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3855
      Left            =   120
      Max             =   300
      Min             =   1
      TabIndex        =   4
      Top             =   120
      Value           =   10
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   100
      Left            =   120
      Max             =   30000
      Min             =   -30000
      TabIndex        =   3
      Top             =   6960
      Width           =   8535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6735
      LargeChange     =   100
      Left            =   8760
      Max             =   30000
      Min             =   -30000
      TabIndex        =   2
      Top             =   120
      Value           =   3000
      Width           =   375
   End
   Begin VB.OptionButton Op2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Op1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v1 As Integer
Dim v2 As Integer

Dim op1L As Long
Dim op1T As Long
Dim op2L As Long
Dim op2T As Long

Dim x As Currency
Dim step As Integer
Dim a As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Sub Form_Load()
v1 = 800
End Sub
Private Sub Timer1_Timer()
x = x + 1 ^ -5
VScroll1.Left = Form1.Width - (VScroll1.Width + 1000)
HScroll1.Top = Form1.Height - (HScroll1.Height + 1000)
HScroll1.Width = Form1.Width - 1000
VScroll1.Height = Form1.Height - 1000
Form1.Cls

v1 = HScroll1.Value
v2 = VScroll1.Value

'v1 = v1 - 100
'If v1 < -9000 Then
'v1 = 200
'v1 = 0
'End If

For a = 0 To Text1.Text
op1T = Tan(x / 100) * Cos(a / 2) * v2 + 5000
op1L = Tan(a / 2) * Cos(a / 2) * v1 + 20000
Line (op1L + x, op1T + x)-(op2L + x, op2T + x), RGB(0, 255, 100)
op2T = Screen.Height - op1T
op2L = Screen.Width - op1L
Next
End Sub
Private Sub VScroll2_Change()
Text1.Text = VScroll2.Value
End Sub
