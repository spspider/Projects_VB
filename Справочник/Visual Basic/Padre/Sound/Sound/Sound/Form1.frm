VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
'Чтобы воспроизвести файл:
Private Sub Command2_Click()
Call mciExecute("play 2.wav") 'Проигрываем файл 2.wav, он должен находиться в той же папке, где и сама прога(НИКОГДА НЕ ПИШИ ПОЛНЫЙ ПУТЬ - например c:\2.wav, - на вражеском компе, звуков не будет, т.к. комп будет искать на другом компе, в диске c:\ звуковой файл, а если ты напишешь 2.wav, то он будет искать в той же папке, где и сама прога)
End Sub

'Чтобы закрыть файл:
Private Sub Command1_Click()
Call mciExecute("close 2.wav")
End Sub
