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
   Begin VB.CommandButton Command3 
      Caption         =   "�������"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�� ����� � ������� General ��������� ���������� ������������ ��� ��� ������������
Dim nid As NOTIFYICONDATA



Private Sub Command1_Click()
' �������� ������ ����� � Traybar
nid.cbSize = Len(nid)
nid.hWnd = Form1.hWnd
nid.uID = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallbackMessage = WM_MOUSEMOVE
nid.hIcon = Form1.Icon
'��� ��������� ������� �� ������, ���������� �����: "� �� ������ ����� �� VBStreets.Narod.RU"
nid.szTip = "� �� ������ ����� �� VBStreets.Narod.RU" & vbNullChar

Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Command2_Click()
nid.hIcon = Form2.Icon
nid.szTip = "New Icon" & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub Command3_Click()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��������� ����������
Dim msg As Long
Dim sFilter As String

msg = X / Screen.TwipsPerPixelX
Select Case msg


       Case WM_LBUTTONDOWN
       '���� �� ������ �������� ���, ������� �������
       MsgBox "������ ����� ������ ����(������)"
       
       
       Case WM_LBUTTONUP
            '���� �� ������ �������� ���, ������� �������
         MsgBox "������ ����� ������ ����(������)"
         
         
       Case WM_LBUTTONDBLCLK
     MsgBox "�� ������� 2 ���� �� ������(����� �������)"
         '���� �� ������ �������� ���, ������� �������
       Case WM_RBUTTONDOWN
           '���� �� ������ �������� ���, ������� �������
             '������ ��� PopupMenu
             
            MsgBox "������ ������ ������ ����(������)"

      
       Case WM_RBUTTONUP
            '���� �� ������ �������� ���, ������� �������
             MsgBox "������ ����� ������ ����(������)"
             
       Case WM_RBUTTONDBLCLK
       
           '���� �� ������ �������� ���, ������� �������
              MsgBox "�� ������� 2 ���� �� ������(������ �������)"
End Select

End Sub
