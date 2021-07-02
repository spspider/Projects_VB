VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "D:\XXX\YYY\AAA\BBB\"
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub MakeDir(dirname As String)

Dim i As Long
Dim path As String

Do

i = InStr(i + 1, dirname & "\", "\")

path = Left$(dirname, i - 1)
If Right$(path, 1) <> ":" And Dir$(path, vbDirectory) = "" Then
MkDir path
End If
Loop Until i >= Len(dirname)
End Sub

Private Sub Command1_Click()
Call MakeDir(Text1.Text)
End Sub

