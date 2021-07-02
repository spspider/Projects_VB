VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LPT"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3480
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "h"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   840
      Top             =   2040
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "000000000"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "23:00"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "000000001"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "07:00"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Inp Lib "inpout32\inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32\inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)
Dim BudVrem As Currency
Dim BudNach As Currency
Dim BudKon As Currency

Private Sub Command1_Click()
Form1.Hide
End Sub

Private Sub Form_Load()



'Text7.Text = BudVrem
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell ("taskkill.exe /f /IM " + App.EXEName + ".exe")
End Sub

Private Sub Timer1_Timer()
If Len(Time) = 8 Then
Text1.Text = Mid(Time, Len(Time) - 7, 5)
Else
Text1.Text = "0" + Mid(Time, Len(Time) - 6, 4)
End If
Text2.Text = ConvertToBin(Inp(888))

If Text1.Text = Text3.Text Then
DEC1
End If

End Sub
Public Function ConvertToBin(ByVal bByte As Byte) As String
 Dim i As Integer
 For i = 0 To 8
  If bByte And 2 ^ i Then
   ConvertToBin = 1 & ConvertToBin
  Else
   ConvertToBin = 0 & ConvertToBin
  End If
 Next i
End Function
Public Function GetDecimalNumber(ByVal strBinaruNumber As String) As Long
 'On Error Resume Next
 Dim strRead As String
 Dim lngResult As Long
 Dim lngCount As Long
 Dim i As Long
 
 lngCount = Len(strBinaruNumber): i = 1
 strRead = Mid(strBinaruNumber, i, 1)
 Do While Not strRead = vbNullString
  lngResult = lngResult + (CLng(strRead) * (2 ^ (lngCount - i)))
  i = i + 1
  strRead = Mid(strBinaruNumber, i, 1)
 Loop
 
 GetDecimalNumber = lngResult
End Function

Function DEC1() ' выполнение заданий, добавляет значения к текущему, если "00000000" - выключает

Dim i2 As Integer
Dim i1 As Integer
Dim CodInp1 As Integer
Dim CodSrav1 As Integer
Dim CodItog1 As Integer

CodInp1 = Val(Text2.Text)
CodSrav1 = Val(Text4.Text)

For i2 = 1 To 8
If Mid(CodInp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 1 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "1" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 0 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "0" + Mid(CodItog1, 10 - i2, 8)
End If
End If
If Mid(CodInp1, 9 - i2, 1) = 1 Then
If Mid(CodSrav1, 9 - i2, 1) = 0 Then
CodItog1 = Mid(CodInp1, 1, 8 - i2) + "0" + Mid(CodItog1, 10 - i2, 8)
End If
End If
Next i2

Out 888, GetDecimalNumber(CodItog1)

End Function

Private Sub Timer2_Timer()
If Timer1.Interval = 1000 Then
Timer1.Enabled = False
BudVrem = Val(Mid(Text1.Text, 1, 2)) + Val(Mid(Text1.Text, 4, 2)) / 60
BudNach = Val(Mid(Text3.Text, 1, 2)) + Val(Mid(Text3.Text, 4, 2)) / 60
BudKon = Val(Mid(Text5.Text, 1, 2)) + Val(Mid(Text5.Text, 4, 2)) / 60
If BudVrem <= BudKon Then
If BudNach <= BudVrem Then
DEC1
End If
End If

End If
End Sub
