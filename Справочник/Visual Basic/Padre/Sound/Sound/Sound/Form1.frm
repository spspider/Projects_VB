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
'����� ������������� ����:
Private Sub Command2_Click()
Call mciExecute("play 2.wav") '����������� ���� 2.wav, �� ������ ���������� � ��� �� �����, ��� � ���� �����(������� �� ���� ������ ���� - �������� c:\2.wav, - �� ��������� �����, ������ �� �����, �.�. ���� ����� ������ �� ������ �����, � ����� c:\ �������� ����, � ���� �� �������� 2.wav, �� �� ����� ������ � ��� �� �����, ��� � ���� �����)
End Sub

'����� ������� ����:
Private Sub Command1_Click()
Call mciExecute("close 2.wav")
End Sub
