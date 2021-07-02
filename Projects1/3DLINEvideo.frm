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
   Begin VB.VScrollBar VScroll1 
      Height          =   6975
      Left            =   8760
      Max             =   5000
      Min             =   -5000
      TabIndex        =   3
      Top             =   120
      Value           =   3000
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   120
      Max             =   5000
      Min             =   -5000
      TabIndex        =   2
      Top             =   7080
      Width           =   8655
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
      Left            =   120
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

Dim a As Integer
Dim step As Integer
Dim x As Integer
Dim r As Integer
Dim l As Integer
Dim u As Integer
Dim d As Integer
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long


Private Sub Form_Load()
v1 = 200
End Sub

Private Sub Timer1_Timer()
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

For a = 0 To 10
Op1.Top = Sin(a / 2) * Tan(a / 4) * v2 + Screen.Height + 6000
Op1.Left = Cos(a / 2) * Tan(a / 4) * v1 + Screen.Width + 8000
Line (Op1.Left + x, Op1.Top + x)-(Op2.Left + x, Op2.Top + x), RGB(0, 255, 100)
Op2.Top = 2 * Screen.Height - Op1.Top
Op2.Left = 2 * Screen.Width - Op1.Left
Next
End Sub
