VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   2850
   ClientTop       =   3675
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5175
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "aA:\."
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Dim Sl As String
Dim Sl1 As Integer
Dim i As Integer
Function slovo()
Call keybd_event(vbKeyTab, 0, 0, 0)
Sl = Text2.Text
For i = 1 To Len(Sl)
Sl1 = Asc(Mid(Sl, i, 1))
'Text3.Text = Text3.Text & Sl1
If Sl1 >= 97 Then ' upper register
Call keybd_event((Val(Asc(Mid(Sl, i, 1)))) - 32, 1, 1, 1)
End If
If Sl1 <= 90 Then ' lower register
Call keybd_event((Val(Asc(Mid(Sl, i, 1)))), 1, 1, 1)
End If
If Sl1 = 92 Then ' \
Call keybd_event(156, 0, 0, 0)
End If
If Sl1 = 58 Then ' :
Call keybd_event(122, 0, 0, 0)
End If
If Sl1 = 46 Then ' .
Call keybd_event(110, 0, 0, 0)
End If
Next i
End Function

Function Rslovo()
Dim rSl As String
For i = 1 To 10
rSl = Fix(Sin(Rnd / 2) * 70 + 97)
If rSl >= 122 Then rSl = 122
If rSl <= 97 Then rSl = 97
Sl = Sl & Chr(rSl)
Next i
Text4.Text = Sl
Sl = ""
End Function
Private Sub Command1_Click()
slovo
End Sub

Private Sub Command2_Click()
Rslovo
End Sub

