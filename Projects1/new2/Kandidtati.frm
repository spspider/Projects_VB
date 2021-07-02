VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   3585
   ClientTop       =   3225
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   2775
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2055
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "1234"
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim o As Integer
Dim n1 As String

Private Sub Command1_Click()
For n = 0 To 4321
n1 = ""
Text1.Text = n

'Text2.Text = Mid(n, 1, 1)

If InStr(1, n, "1") <> 0 Then n1 = n1 + "A"
If InStr(1, n, "2") <> 0 Then n1 = n1 + "B"
If InStr(1, n, "3") <> 0 Then n1 = n1 + "C"
If InStr(1, n, "4") <> 0 Then n1 = n1 + "D"
If InStr(1, n, "5") <> 0 Then n1 = n1 + "E"
If InStr(1, n, "6") <> 0 Then n1 = n1 + "F"
If InStr(1, n, "7") <> 0 Then n1 = n1 + "G"
If InStr(1, n, "8") <> 0 Then n1 = n1 + "H"

Text2.Text = n1
If Len(n1) = 4 Then

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If InStr(1, n1, "A") <> 0 Then
If InStr(1, n1, "B") <> 0 Then
  If InStr(1, n1, "C") <> 0 Then

If InStr(1, n1, "D") <> 0 Then
  If InStr(1, n1, "H") <> 0 Then
If InStr(1, n1, "E") <> 0 Then
Else
  List1.AddItem (n1)
  End If
  End If
End If
  
  End If
End If



If InStr(1, n1, "F") <> 0 Then
  If InStr(1, n1, "G") <> 0 Then
  If InStr(1, n1, "D") <> 0 Then
  List1.AddItem (n1)
End If
End If
End If
  
  

'If InStr(1, n1, "D") <> 0 Then
'  If InStr(1, n1, "H") Then
  
'If InStr(1, n1, "E") <> 0 Then

'If InStr(1, n1, "F") <> 0 Then
'   If InStr(1, n1, "G") Then '
'   If InStr(1, n1, "D") Then
    
'If InStr(1, n1, "G") <> 0 Then
'If InStr(1, n1, "H") <> 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''

End If
Next
End Sub



Private Sub Form_Load()
n = 1234
o = 4
End Sub

Private Sub Timer1_Timer()
n = n + 1
n1 = ""
Text1.Text = n

'Text2.Text = Mid(n, 1, 1)

If InStr(1, n, "1") <> 0 Then n1 = n1 + "A"
If InStr(1, n, "2") <> 0 Then n1 = n1 + "B"
If InStr(1, n, "3") <> 0 Then n1 = n1 + "C"
If InStr(1, n, "4") <> 0 Then n1 = n1 + "D"
If InStr(1, n, "5") <> 0 Then n1 = n1 + "E"
If InStr(1, n, "6") <> 0 Then n1 = n1 + "F"
If InStr(1, n, "7") <> 0 Then n1 = n1 + "G"
If InStr(1, n, "8") <> 0 Then n1 = n1 + "H"

Text2.Text = n1
If Len(n1) = 4 Then

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If InStr(1, n1, "A") <> 0 Then
If InStr(1, n1, "B") <> 0 Then
  If InStr(1, n1, "C") <> 0 Then

If InStr(1, n1, "D") <> 0 Then
  If InStr(1, n1, "H") <> 0 Then
If InStr(1, n1, "E") <> 0 Then
Else
  List1.AddItem (n1)
  End If
  End If
End If
  
  End If
End If



If InStr(1, n1, "F") <> 0 Then
  If InStr(1, n1, "G") <> 0 Then
  If InStr(1, n1, "D") <> 0 Then
  List1.AddItem (n1)
End If
End If
End If
  
  

'If InStr(1, n1, "D") <> 0 Then
'  If InStr(1, n1, "H") Then
  
'If InStr(1, n1, "E") <> 0 Then

'If InStr(1, n1, "F") <> 0 Then
'   If InStr(1, n1, "G") Then '
'   If InStr(1, n1, "D") Then
    
'If InStr(1, n1, "G") <> 0 Then
'If InStr(1, n1, "H") <> 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''

End If


If n = 4321 Then Timer1.Enabled = False
End Sub
