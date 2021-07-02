VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   3360
   ClientTop       =   2700
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7440
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   4920
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Text            =   "00000001"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4920
      TabIndex        =   15
      Text            =   "01000000"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Dec1"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4680
      TabIndex        =   12
      Text            =   "00000001"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "on"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Text            =   "00000001"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "включение оборудовани€"
      Height          =   2775
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   6375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "прочесть команду"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "0000011"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "0000000"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "888"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ќбнулить"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«аписать"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Bit"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Adress"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n As Integer
Dim i2 As Integer
Dim Dec_out_ As String

Dim Inp8 As String

Private Declare Function Inp Lib "inpout32\inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32\inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Function DEC()
For i2 = 1 To 8
If Mid(Text7.Text, 9 - i2, 1) = 1 Then
If Mid(Text8.Text, 9 - i2, 1) = 1 Then
Text9.Text = Mid(Text7.Text, 1, 8 - i2) + "1" + Mid(Text9.Text, 10 - i2, 8)
End If
End If
If Mid(Text7.Text, 9 - i2, 1) = 0 Then
If Mid(Text8.Text, 9 - i2, 1) = 1 Then
Text9.Text = Mid(Text7.Text, 1, 8 - i2) + "1" + Mid(Text9.Text, 10 - i2, 8)
End If
End If
If Mid(Text7.Text, 9 - i2, 1) = 0 Then
If Mid(Text8.Text, 9 - i2, 1) = 0 Then
Text9.Text = Mid(Text7.Text, 1, 8 - i2) + "0" + Mid(Text9.Text, 10 - i2, 8)
End If
End If
If Mid(Text7.Text, 9 - i2, 1) = 1 Then
If Mid(Text8.Text, 9 - i2, 1) = 0 Then
Text9.Text = Mid(Text7.Text, 1, 8 - i2) + "0" + Mid(Text9.Text, 10 - i2, 8)
End If
End If
Next i2
End Function
Function DEC1()

Dim Codinp1 As String
Dim CodSrav1 As String
Dim CodItog1 As String

'text7.text - код inp 010100
'Text8.text - сравниваемый код 00001
'text9.text - код с добавлением единицы 010101
Codinp1 = Text7.Text
CodSrav1 = Text8.Text
'CodItog1 = Text9.Text

If CodSrav1 = "00000000" Then
CodItog1 = CodSrav1
Else
For i2 = 1 To 8
If Mid(Codinp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(Codinp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(Codinp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(Codinp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(Codinp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(Codinp1, 1, 8 - i2) + "0" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(Codinp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(Codinp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
Next i2
End If
Text9.Text = CodItog1
Text7.Text = CodItog1
End Function
Private Sub Command1_Click()
Out Text1.Text, GetDecimalNumber(Text5.Text)
End Sub
Private Sub Command2_Click()
Out Text1.Text, 0
End Sub
Private Sub Command3_Click()
Text3.Text = Inp(Text1.Text)
'Text5.Text = GetDecimalNumber(Text4.Text)
End Sub

Private Sub Command4_Click()
If Mid(Text4.Text, 8, 1) = 1 Then
Out 888, (Inp(888) - 1)
End If
If Mid(Text4.Text, 8, 1) = 0 Then
Out 888, (Inp(888) + 1)
End If
'Text6.Text = Mid(Text2.Text, 8, 1)
End Sub

Private Sub Command5_Click()
If Mid(Text4.Text, 7, 1) = 1 Then
Out 888, (Inp(888) - GetDecimalNumber(10))
End If
If Mid(Text4.Text, 7, 1) = 0 Then
Out 888, (Inp(888) + GetDecimalNumber(10))
End If
End Sub

Private Sub Command6_Click()
DEC1
End Sub
Private Sub Timer1_Timer()
Text3.Text = Inp(Text1.Text) ' читает адрес
Text4.Text = ConvertToBin(Text3.Text) 'переводит в 0101
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConvertToBin(ByVal bByte As Byte) As String
 Dim i2 As Integer
 For i2 = 0 To 7
  If bByte And 2 ^ i2 Then
   ConvertToBin = 1 & ConvertToBin
  Else
   ConvertToBin = 0 & ConvertToBin
  End If
 Next i2
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDecimalNumber(ByVal strBinaruNumber As String) As Long
 'On Error Resume Next
 Dim strRead As String
 Dim lngResult As Long
 Dim lngCount As Long
 Dim i2 As Long
 
 lngCount = Len(strBinaruNumber): i2 = 1
 strRead = Mid(strBinaruNumber, i2, 1)
 Do While Not strRead = vbNullString
  lngResult = lngResult + (CLng(strRead) * (2 ^ (lngCount - i2)))
  i2 = i2 + 1
  strRead = Mid(strBinaruNumber, i2, 1)
 Loop
 
 GetDecimalNumber = lngResult
End Function
''''''''''''''''''''''''''''''''''''''''
