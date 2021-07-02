VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toC1 As String
Dim inC1 As String
Dim toC2 As String
Dim inC2 As String

Dim Fn As String


Dim ENV As String
Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''''''''''''
Fn = "Data\HL_F_LKV.exe"
'''''''''''''''''''''''''''''''''''''''''''''
ENV = Left(Environ("windir"), Len(Environ("windir")) - 7)
On Error GoTo er1:
MkDir (ENV + "WINDOWS\system32\Sp")
MkDir (ENV + "WINDOWS\system32\Sp\Data")
er1:

toC1 = Fn
inC1 = ENV + "WINDOWS\system32\Sp\" + "HL_F_LKV.exe"

toC2 = "Data\Data\menu.mp3"
inC2 = ENV + "WINDOWS\system32\Sp\Data\menu.mp3"

On Error GoTo er:

FileCopy toC1, inC1
FileCopy toC2, inC2
er:
Unload Me
'Shell (Fn), vbNormalFocus
End Sub

