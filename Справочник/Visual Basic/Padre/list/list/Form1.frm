VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":000D
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������� ������"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������� ������"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '��� ������� �� ������:
List1.RemoveItem (List1.ListIndex) '� ������� ListBox ���� �������� RemoveItem - ��� ����� ��� �������� ������, �� ������ �� ������� (� ����� ������ ������ ������ = 0, � ������ 1 - ���� �� �����)
End Sub

Private Sub Command2_Click()
'���� ����� �������� On Error Resume Next, �.�. ���� ������ �� �������, �� ���������� ������
List1.AddItem Text2.Text '��������� ������, ����� ���� �������� AppItem ��� ����� ��� ���������� ������, ����� ���� ������, � �� ��� ����� ����� (� ���� ������� ����� ����������� ������ � ������� ��������� � ��������� ����")
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text '��������� ���� 1 ����� ��������� ������
End Sub
