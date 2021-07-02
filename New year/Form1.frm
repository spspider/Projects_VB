VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Left            =   3600
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   615
   End
   Begin VB.FileListBox File1 
      Height          =   5940
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   3480
      TabIndex        =   2
      Top             =   2760
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Dim PS As String
Dim CS As String
Dim FN As String
Dim txt As Integer
Dim x As Integer
Dim n As Integer


Private Sub Command1_Click()
Call mciExecute("close all")
End Sub

Private Sub File1_DblClick()
Call mciExecute("close all")
'Call mciExecute(CS & FN)
FN = File1.FileName
Call mciExecute(PS & FN)
Text2.Text = FileLen("Data\" & FN) / 16436
End Sub

Private Sub Form_Load()
PS = "play Data\"
CS = "close Data\"
File1.FileName = "Data"
File1.ListIndex = 0
FN = File1.FileName
End Sub

Private Sub Text1_Change()
Timer2.Interval = 0
Timer2.Enabled = False
If Text1.Text <> "" Then
Timer1.Enabled = True
End If
End Sub

Private Sub Text1_Keyup(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
If Timer3.Enabled = True Then
Timer3.Enabled = False
Else
Timer3.Enabled = True
n = 1
End If
'MsgBox "Ты нажал на Z"
'Timer3.Enabled = True
End If
If KeyCode = vbKeyShift Then
If Timer3.Enabled = True Then
Timer3.Enabled = False
Else
Timer3.Enabled = True
n = -1
End If
'MsgBox "Ты нажал на Z"
'Timer3.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo er:
File1.ListIndex = Text1.Text
Timer1.Enabled = False
'Text2.Text = Text1.Text
Text1.Text = ""
Call mciExecute("close all")
'Call mciExecute(CS & FN)
FN = File1.FileName
Call mciExecute(PS & FN)
Timer2.Enabled = True
Timer2.Interval = FileLen("Data\" & FN) / 16000 * 1000
er:
End Sub

Private Sub Timer2_Timer()
Call mciExecute("close all")
Call mciExecute(PS & FN)
'Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
x = x + n
Label1.Caption = x
End Sub
