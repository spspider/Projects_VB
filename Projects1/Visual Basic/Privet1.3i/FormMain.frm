VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FormMain 
   Caption         =   "��������"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1260
      Top             =   540
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------
'MS Agent
'------------------------------------------
Dim MyChar As Object '������ Characters
Dim LoadRequest As Object
'���� �������
Private Const LanguageID_Ru = &H419
'-----------------------------------------------
'����� �����
Private Const TTSModeID_Boris = "{06377F81-D48E-11D1-B17B-0020AFED142E}"
'����� ��������
Private Const TTSModeID_Svetlana = "{06377F80-D48E-11d1-B17B-0020AFED142E}"
'-----------------------------------------------
Dim MyAgentLeft As Long
Dim MyAgentTop As Long
'------------------------------------------
Dim MyCharTTSModeID As String
'-----------------------------------------------
'MS Agent
'------------------------------------------






'--------------------------------------------------
'User Name
'--------------------------------------------------
Dim FlagMale As Boolean
Dim UserNames(1 To 5) As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'--------------------------------------------------
'User Name
'--------------------------------------------------


Dim ITimer As Integer '����� ������� 10 ���
Dim FlagStop As Boolean '���



Private Sub Form_Load()
'--------------------------------
'--------------------------------
'--------------------------------
On Error Resume Next
'--------------------------------
' ��������� �������� �� ���������
MyAgent.Characters.Load "MyChar", "merlin.acs"
Set MyChar = MyAgent.Characters("MyChar")
'--------------------------------
If Err.Number Then
MsgBox "MS Agent �� ������! ��� ���� ��������� �� ���� ���������� �������� �� ������."
Unload Me
Exit Sub
End If
'--------------------------------
On Error GoTo 0
'--------------------------------
'--------------------------------
'--------------------------------
MyChar.LanguageID = LanguageID_Ru ' ��������� ������� ����
'--------------------------------
'--------------------------------
'--------------------------------
On Error Resume Next
'--------------------------------
MyChar.TTSModeID = TTSModeID_Boris ' ��������� ������ - ����� ������ (������� ������)
'--------------------------------
'If Err.Number Then '���� ������
'MsgBox "����� MS Agent �� ��������!"
'End If
'--------------------------------
On Error GoTo 0
'--------------------------------
'--------------------------------
'--------------------------------
MyChar.Commands.Add "Congratulation", "&��������"
MyChar.Commands.Add "StopMe", "&�����"
MyChar.Commands.Add "TimeCommand", "&������� ���?"
MyChar.Commands.Add "Exit", "�&����"
'--------------------------------
Get_UserName 'User Name
'--------------------------------
Congratulation0
'--------------------------------
ITimer = 0
FlagStop = False
Timer1.Interval = 1500
Timer1.Enabled = True
'--------------------------------
End Sub




Private Sub Congratulation0()
'--------------------------------
MyChar.Stop
MyChar.Play "RestPose"
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
MyChar.Show ' ������� �� �����
'--------------------------------
MyChar.Play "Wave"
MyChar.Speak "��������, " & UserNames(NRnd(1, 5)) & "!"
MyChar.Speak "����� ����������!|����� ���������!|����� ���������!|����� �����������!"
MyChar.Play "RestPose"
'--------------------------------
End Sub




Private Sub Congratulation()
Dim IVar As Integer
'--------------------------------
IVar = NRnd(1, 2)
'--------------------------------
FlagStop = False
'--------------------------------
MyChar.Stop
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 '�������
MyChar.Speak UserNames(NRnd(1, 5)) & ", ��� �����������?|" & UserNames(NRnd(1, 5)) & ", ��� �������?|" & UserNames(NRnd(1, 5)) & ", ��� ����?|" & UserNames(NRnd(1, 5)) & ", ��� ����������?"
MyChar.Play "Startlistening"
MyChar.Speak "��, �������� ��-����?|�� ����� ���� - �� �����!|����� �������?|��� ������?|��� �������?|�������� �� ����!"
Case 2 '��������
MyChar.Speak UserNames(NRnd(1, 5)) & ", �������!|" & UserNames(NRnd(1, 5)) & ", �������� �����?|" & UserNames(NRnd(1, 5)) & ", ����� �������!|" & UserNames(NRnd(1, 5)) & ", �� �� ����!|" & UserNames(NRnd(1, 5)) & ", ������!|" & UserNames(NRnd(1, 5)) & ", �� ���������!"
MyChar.Play "Uncertain"
MyChar.Speak "�� ��� ��� ��� �����!|����� �������!|������� ������ �� ����!|��� �� ���� ������ �� �����!|�� �����-�� � �������� ��������!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 '�������
    MyChar.Play "Greet"
    If FlagMale = True Then '�����
    MyChar.Speak UserNames(NRnd(1, 5)) & ", ������� ���������!|" & UserNames(NRnd(1, 5)) & ", �� ������ ��� ����!|" & UserNames(NRnd(1, 5)) & ", �� ��������� �����!|" & UserNames(NRnd(1, 5)) & ", � ������� �����!"
    MyChar.Play "Acknowledge"
    MyChar.Speak "� �����, �� �������!|��� �������!|��� �� ����!|��� ������!"
    Else
    MyChar.Speak UserNames(NRnd(1, 5)) & ", ������� ���������!|" & UserNames(NRnd(1, 5)) & ", �� ����� ���������!|" & UserNames(NRnd(1, 5)) & ", �� ����� ������!|" & UserNames(NRnd(1, 5)) & ", �� ��� ���������!|" & UserNames(NRnd(1, 5)) & ", � ����� ����� �����, ��� ��!"
    MyChar.Play "GestureUp"
    MyChar.Speak "���������� ���!|������ ����!|������ ���!|������� ���!|�������� ���!|��������� ��!|���������!|�������� ��!|����� ���!|������� ���!|� ���� ������� ������, ��� �� ��������������� �����!|� ���� �� ������ ���, ��� ����, �� ���� ��-������!|�, ����� ���� ����!|�, ��������� ����� ������!|�, ��������� ���� ������!|�, ������������!"
    End If
Case 2 '��������
    MyChar.Play "Decline"
    If FlagMale = True Then '�����
    MyChar.Speak UserNames(NRnd(1, 5)) & ", ���������� �� ���-��?|" & UserNames(NRnd(1, 5)) & ", ������ �� ����!|" & UserNames(NRnd(1, 5)) & ", ���-�� �� ������� �������!|" & UserNames(NRnd(1, 5)) & ", �� � ����� � ����!|" & UserNames(NRnd(1, 5)) & ", �� ��� �� ��������?"
    Else
    MyChar.Speak UserNames(NRnd(1, 5)) & ", ����� � ������� ����������?|" & UserNames(NRnd(1, 5)) & ", ������ � ����� ���-����!|" & UserNames(NRnd(1, 5)) & ", ���-�� �� ������� �������!|" & UserNames(NRnd(1, 5)) & ", �� � ����� � ����!|" & UserNames(NRnd(1, 5)) & ", �������� �������� ���-��?|" & UserNames(NRnd(1, 5)) & ", �� ��� �� ���������?"
    End If
    MyChar.Play "Sad"
    MyChar.Speak "������, ���� ������!|�� � ����!|��� ������� �� ����!|��� �����!|� �� �������!|��� �� ���������!|��� ������!|��� �������������!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 '�������
MyChar.Play "GetAttentionContinued"
MyChar.Play "GetAttentionContinued"
MyChar.Play "GetAttentionContinued"
MyChar.Speak UserNames(NRnd(1, 5)) & ", ������ ��������!|" & UserNames(NRnd(1, 5)) & ", �� ����� �������!|" & UserNames(NRnd(1, 5)) & ", ���������� ���� �������!|" & UserNames(NRnd(1, 5)) & ", ���� ������� �������������!"
MyChar.Play "GestureDown"
MyChar.Speak "������, ����� ������, ������� ��� ����� ����?|��, ����� ������ ������!|�� ���� ������?|��������, �� �� ����?|����� ������ ���-��!"
Case 2 '��������
MyChar.Play "Idle3_1"
MyChar.Speak UserNames(NRnd(1, 5)) & ", � ����� �� ���������!|" & UserNames(NRnd(1, 5)) & ", ��� ������ � ����!|" & UserNames(NRnd(1, 5)) & ", � ����� ���� �����!|" & UserNames(NRnd(1, 5)) & ", � ��� � ��� � ����� ������?"
MyChar.Play "Explain"
MyChar.Speak "���� � �� ����!|� ������ ����� �����!|���� ��� ��������� ������!|������ ������� �� ����!|����� �� ���� ���� ������!|��� ��� ��� �?|� �� ����� ������ �� ���������!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
MyChar.Play "ReadContinued"
'--------------------------------
'��������
'--------------------------------
Select Case NRnd(1, 2) 'IVar

Case 1
MyChar.Speak "���� ���� ������ �������, " & UserNames(NRnd(1, 5)) & ", �� ����� �������.|" _
& "���� � ���� ��� ����, " & UserNames(NRnd(1, 5)) & ", �� �� ����� � ����������.|" _
& "������ ����� ������ ��, ��� ���� �� �����, " & UserNames(NRnd(1, 5)) & ", ��� ��, ��� ���� �����.|" _
& "��������� ������, " & UserNames(NRnd(1, 5)) & ", ����������� ������ ����� ����.|" _
& "��� ����� ���� ���� ������, " & UserNames(NRnd(1, 5)) & "? - ��������� ������ � �� ��������� ��.|" _
& "����� ������� ������ �������, " & UserNames(NRnd(1, 5)) & ", ����� ����� ���������� ����������.|" _
& "����� ��, ��� �� ���� �������, " & UserNames(NRnd(1, 5)) & ", � �� ���������, �����, ��� ����� �� ���� �������!|" _
& "�������� ���, ��� �������� �� ������ �������, " & UserNames(NRnd(1, 5)) & ", � �� ������� ������ ���������.|" _
& "������, " & UserNames(NRnd(1, 5)) & ", ���� ��� ������� ������ �������.|" _
& "����� ������ �������, " & UserNames(NRnd(1, 5)) & ", �� �� ����� �������� �����.|" _
& "������� ����, " & UserNames(NRnd(1, 5)) & ", ������� �� �����.|" _
& "����� ����� �������� ������� �����, " & UserNames(NRnd(1, 5)) & ", �� ��� ��������������.|" _
& "� ������ ������ �������, " & UserNames(NRnd(1, 5)) & ", ����� ��, ��� �������� ����� ������.|" _
& "���� ���� ����������� �� ����������, " & UserNames(NRnd(1, 5)) & ", ������, ���� ���� ���������.|" _
& "�� �����, ��� ���� �� ����� ����, " & UserNames(NRnd(1, 5)) & ", �� �����, ��� �� �� ������ �����.|" _
& "������ ������� �������� �� �����, " & UserNames(NRnd(1, 5)) & ", � ������ ������� ������� ���� �����.|" _
& "������ ������� �����, ����� � ���� �� ���� �������, " & UserNames(NRnd(1, 5)) & ", � ������ ������� ������ ������ ���� ��������.|" _
& "����� ���� ������� ��������, " & UserNames(NRnd(1, 5)) & ", �� ��������� �����.|" _
& "���� �������, " & UserNames(NRnd(1, 5)) & ", �� ������� ������ ������� ������ �������� ���� ���.|" _
& "���� �� �� �������� ���� ������, " & UserNames(NRnd(1, 5)) & ", �� ��� �� � �������?|" _
& "��������� ������� �� ������� �������� �����, " & UserNames(NRnd(1, 5)) & ", ��� ������� ��� ��������� ������� ��������������.|" _
& "������ ������� ������ ����� ����� ���� ����, " & UserNames(NRnd(1, 5)) & ", � ������ ������� ������������ ������ - ���� ������ �����.|" _
& "������� �� ����� �� ��������, " & UserNames(NRnd(1, 5)) & " - �� ���� �����������, ���� �����������.|" _
& "�� ������ ���� ������ - ����� ������, " & UserNames(NRnd(1, 5)) & ", �� ������ ���� ������ - ����� ������!|" _
& "����� ������, " & UserNames(NRnd(1, 5)) & ", ����� ����: ''�� ����� ������!''"

Case 2
MyChar.Speak "������� ��������� ��������� ��������� ���� �������, " & UserNames(NRnd(1, 5)) & ", �� ����� ��� �����������, ��� ��� ����������.|" _
& "������� � ������� �� �������, " & UserNames(NRnd(1, 5)) & ", �� ����� ��������� ����� ��������� ��� ������������.|" _
& "������� ���������� ������� ����� �������� �����, " & UserNames(NRnd(1, 5)) & ", � ������� - �������, �������, �������!|" _
& "��������� �������, ��� � ��� ������ ������, ���� �� ����� ����, " & UserNames(NRnd(1, 5)) & ", � ��� ���� ������ ����.|" _
& "������� ������� ��� ����, ����� � ���������, " & UserNames(NRnd(1, 5)) & ", � �� ��� ����, ����� �� ���� �����.|" _
& "�� ������ �����������, " & UserNames(NRnd(1, 5)) & ", ������ ��� �� ����, ��� ��� ����������.|" _
& "������� �������� �� ��, " & UserNames(NRnd(1, 5)) & ", � ������� - �� �� ���������.|" _
& "�������, " & UserNames(NRnd(1, 5)) & ", ������ ���� ������, ��� ��� ������ ������� � ����� ������ ��.|" _
& "��� ������� ����� �������, " & UserNames(NRnd(1, 5)) & ", ������ �� ����� ������� ���� ���?|" _
& "���� ������� ��� ������� ��, ��� ��� ������, " & UserNames(NRnd(1, 5)) & ", ������, ��� ������ ������� ����!|" _
& "��� �� ����� ����������� ��������� ������, " & UserNames(NRnd(1, 5)) & ", ��������� ������ ����������� ��� ����.|" _
& "��� ������ ����� ������, " & UserNames(NRnd(1, 5)) & ", � ��� ������� - ����������!|" _
& "����� ���� ��� � ��� ������, " & UserNames(NRnd(1, 5)) & ", ��� ������ ���� ������ �����!|" _
& "�������, " & UserNames(NRnd(1, 5)) & " - ��� ����� ������� ��������� ������� �� �������.|" _
& "����������� �����, " & UserNames(NRnd(1, 5)) & ", ������� ����, ��� ������������ ����-���� �������.|" _
& "������ 20% ������ ���� 80% ����������, " & UserNames(NRnd(1, 5)) & ", � ���������� 80% ������ - ����� ���� 20% ����������.|" _
& "����, " & UserNames(NRnd(1, 5)) & ", ������ ������ ����� ������� ��� ���������, ��� ������ �� �����.|" _
& "��� �� ���� �� ��������, " & UserNames(NRnd(1, 5)) & ", �� ��������� ����� ������ ������� ������ �� ���, ������ �� ���, ������ �� ���!|" _
& "���� ������ �������, " & UserNames(NRnd(1, 5)) & ", ����� ������� ���� ����������.|" _
& "�������� ��� �������, " & UserNames(NRnd(1, 5)) & ", ������ ����, ����� ������������ ��-�� ����.|" _
& "��, " & UserNames(NRnd(1, 5)) & ", �������������� ��� ���, ���� ����� �����������.|" _
& "����� ��, ��� �� ������ ����� �������, " & UserNames(NRnd(1, 5)) & ", � �� �������� �����!|" _
& "������������ ���� ���� �� ������ ����, " & UserNames(NRnd(1, 5)) & ", � ������������ ���������� ������ �����.|" _
& "���� ���������� - ��� ��������, ���� ���������� - ��� ���� ��������, � �����, " & UserNames(NRnd(1, 5)) & ", ������� ������� �� ����!|" _
& "������ ������ ������� ������, " & UserNames(NRnd(1, 5)) & ", � ����� ������ ��, ��� �� ������ ��������."

End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
End Sub



'-----------------------------------------------------
'��������� ���������� �����
'-----------------------------------------------------
Private Function NRnd(Nmin As Integer, Nmax As Integer) As Integer
Randomize
NRnd = Int((Nmax - Nmin + 1) * Rnd) + Nmin '������������ �����
End Function
'-----------------------------------------------------
'��������� ���������� �����
'-----------------------------------------------------





Private Sub MyAgent_Command(ByVal UserInput As Object)

Select Case UserInput.Name

Case "Congratulation" '������������
Timer1.Enabled = False
ITimer = 0
Timer1.Interval = NRnd(8, 10) * 1000
Congratulation
Timer1.Enabled = True

Case "StopMe" '����������
FlagStop = True
Timer1.Enabled = False
ITimer = 0
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Play "Blink"
MyChar.Play "RestPose"
MyChar.Play "Idle3_2"

Case "TimeCommand"
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Speak UserNames(NRnd(1, 5)) & ", ������ " & Format(Time, "h:nn")
If FlagStop = False Then
    Timer1.Interval = 1000
    Timer1.Enabled = True
Else
    MyChar.Play "RestPose"
    MyChar.Play "Blink"
    MyChar.Play "RestPose"
    MyChar.Play "Idle3_2"
End If

Case "Exit" '�������
Timer1.Enabled = False
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Play "Wave"
MyChar.Speak UserNames(NRnd(1, 5)) & ", ����!|" & UserNames(NRnd(1, 5)) & ", ���������!|" & UserNames(NRnd(1, 5)) & ", �� �������!|" & UserNames(NRnd(1, 5)) & ", ��������!|" & UserNames(NRnd(1, 5)) & ", �� �������!"
Set LoadRequest = MyChar.Hide '����� ��� ������ ���� �������

End Select

End Sub


Private Sub MyAgent_RequestComplete(ByVal Request As Object)
Select Case Request
Case LoadRequest '����� ��� ������ ���� �������
    MyAgent.Characters.Unload "MyChar"
    Set MyChar = Nothing
    Unload Me
End Select
End Sub



Private Sub Timer1_Timer()
ITimer = ITimer + 1
If ITimer > 5 Then '����� 5*(8...10) ������ �������� ��������
    Timer1.Enabled = False
    ITimer = 0
    Timer1.Interval = NRnd(8, 10) * 1000
    Congratulation
    Timer1.Enabled = True
End If
End Sub


Private Sub MyAgent_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Timer1.Enabled = False
ITimer = 0 '���������� ������
Timer1.Enabled = True
End Sub








'--------------------------------------------------
'User Name
'--------------------------------------------------
Private Sub Get_UserName()
Dim strUserName As String

Select Case NRnd(1, 2) 'IVar

Case 1 '�����
FlagMale = True

Case 2 '�������
FlagMale = False

End Select


If FlagMale = True Then '�����
    UserNames(1) = "����������"
    UserNames(2) = "����������"
    UserNames(3) = "����������"
    UserNames(4) = "����������"
    UserNames(5) = "����������"
Else
    UserNames(1) = "����������"
    UserNames(2) = "����������"
    UserNames(3) = "����������"
    UserNames(4) = "����������"
    UserNames(5) = "����������"
End If

On Error GoTo ErrorHandler

strUserName = Space(255) 'Create a buffer
Call GetUserName(strUserName, 255)   'Get the username
strUserName = StripTerm(strUserName)

Select Case strUserName '��� �������� ������������
    
'����� �� ������������� ����� ���� ������������� ����� ����
'������ "Alexandr", "Alexandr1", "Alexandr2" � �.�. ��������
'�������� ����� �������������, ��������: "PoloviyAO" - ��� ������� ���

    Case "Alexandr", "Alexandr1", "Alexandr2", "PoloviyAO" '��� ������������ � ������ "���������"
    UserNames(1) = "���������"
    UserNames(2) = "����"
    UserNames(3) = "�����"
    UserNames(4) = "�����"
    UserNames(5) = "��������"
    FlagMale = True '�������

    Case "Alexandra", "Alexandra1", "Alexandra2" '��� ������������ � ������ "����������"
    UserNames(1) = "����������"
    UserNames(2) = "����"
    UserNames(3) = "�������"
    UserNames(4) = "����"
    UserNames(5) = "��������"
    FlagMale = False '�������
    
    Case "Alexey"
    UserNames(1) = "�������"
    UserNames(2) = "˸��"
    UserNames(3) = "˸���"
    UserNames(4) = "�����"
    UserNames(5) = "��������"
    FlagMale = True

    Case "Anatoliy"
    UserNames(1) = "��������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "�����"
    UserNames(5) = "�������"
    FlagMale = True

    Case "Valentina"
    UserNames(1) = "���������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "�������"
    UserNames(5) = "�������"
    FlagMale = False
    
    Case "Valeriy"
    UserNames(1) = "�������"
    UserNames(2) = "������"
    UserNames(3) = "�������"
    UserNames(4) = "��������"
    UserNames(5) = "���������"
    FlagMale = True
    
    Case "Viktor"
    UserNames(1) = "������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "����"
    UserNames(5) = "�������"
    FlagMale = True

    Case "Vladimir"
    UserNames(1) = "��������"
    UserNames(2) = "����"
    UserNames(3) = "�����"
    UserNames(4) = "������"
    UserNames(5) = "����������"
    FlagMale = True

    Case "Vyacheslav"
    UserNames(1) = "��������"
    UserNames(2) = "�����"
    UserNames(3) = "������"
    UserNames(4) = "������"
    UserNames(5) = "��������"
    FlagMale = True
    
    Case "Galina"
    UserNames(1) = "������"
    UserNames(2) = "����"
    UserNames(3) = "�������"
    UserNames(4) = "��������"
    UserNames(5) = "�������"
    FlagMale = False
    
    Case "Dariya"
    UserNames(1) = "�����"
    UserNames(2) = "����"
    UserNames(3) = "�������"
    UserNames(4) = "������"
    UserNames(5) = "��������"
    FlagMale = False

    Case "Denis"
    UserNames(1) = "�����"
    UserNames(2) = "����"
    UserNames(3) = "�������"
    UserNames(4) = "���������"
    UserNames(5) = "���������"
    FlagMale = True
    
    Case "Dmitriy"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "�����"
    UserNames(4) = "�����"
    UserNames(5) = "�������"
    FlagMale = True

    Case "Evgeniya"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "����"
    UserNames(4) = "�������"
    UserNames(5) = "�������"
    FlagMale = False

    Case "Elena"
    UserNames(1) = "�����"
    UserNames(2) = "����"
    UserNames(3) = "��������"
    UserNames(4) = "�������"
    UserNames(5) = "�������"
    FlagMale = False

    Case "Ivan"
    UserNames(1) = "����"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "������"
    UserNames(5) = "�������"
    FlagMale = True

    Case "Igor"
    UserNames(1) = "�����"
    UserNames(2) = "������"
    UserNames(3) = "�������"
    UserNames(4) = "�����"
    UserNames(5) = "����������"
    FlagMale = True
    
    Case "Irina"
    UserNames(1) = "�����"
    UserNames(2) = "���"
    UserNames(3) = "������"
    UserNames(4) = "������"
    UserNames(5) = "���������"
    FlagMale = False
    
    Case "Kseniya"
    UserNames(1) = "������"
    UserNames(2) = "�����"
    UserNames(3) = "�������"
    UserNames(4) = "��������"
    UserNames(5) = "���������"
    FlagMale = False

    Case "Lubov"
    UserNames(1) = "������"
    UserNames(2) = "������"
    UserNames(3) = "�������"
    UserNames(4) = "����"
    UserNames(5) = "����������"
    FlagMale = False

    Case "Ludmila"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "����"
    UserNames(4) = "�������"
    UserNames(5) = "�������"
    FlagMale = False
    
    Case "Michail"
    UserNames(1) = "������"
    UserNames(2) = "����"
    UserNames(3) = "�����"
    UserNames(4) = "����"
    UserNames(5) = "��������"
    FlagMale = True
    
    Case "Mariya"
    UserNames(1) = "�����"
    UserNames(2) = "����"
    UserNames(3) = "�������"
    UserNames(4) = "������"
    UserNames(5) = "��������"
    FlagMale = False
    
    Case "Nadegda"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(1) = "��������"
    UserNames(1) = "�������"
    FlagMale = False

    Case "Nataliya"
    UserNames(1) = "�������"
    UserNames(2) = "������"
    UserNames(3) = "���������"
    UserNames(4) = "�������"
    UserNames(5) = "����������"
    FlagMale = False

    Case "Nikolay"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "����"
    UserNames(5) = "�������"
    FlagMale = True

    Case "Nina"
    UserNames(1) = "����"
    UserNames(2) = "�������"
    UserNames(3) = "�������"
    UserNames(4) = "������"
    UserNames(5) = "�������"
    FlagMale = False

    Case "Oksana"
    UserNames(1) = "������"
    UserNames(1) = "�����"
    UserNames(1) = "�������"
    UserNames(1) = "����"
    UserNames(1) = "���������"
    FlagMale = False

    Case "Oleg"
    UserNames(1) = "����"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "������"
    UserNames(5) = "���������"
    FlagMale = True

    Case "Olesya"
    UserNames(1) = "�����"
    UserNames(2) = "�����"
    UserNames(3) = "��������"
    UserNames(4) = "���������"
    UserNames(5) = "��������"
    FlagMale = False

    Case "Olga"
    UserNames(1) = "�����"
    UserNames(2) = "���"
    UserNames(3) = "������"
    UserNames(4) = "�������"
    UserNames(5) = "������"
    FlagMale = False

    Case "Pavel"
    UserNames(1) = "�����"
    UserNames(2) = "����"
    UserNames(3) = "�����"
    UserNames(4) = "�����"
    UserNames(5) = "�����������"
    FlagMale = True

    Case "Svetlana"
    UserNames(1) = "��������"
    UserNames(2) = "�����"
    UserNames(3) = "�������"
    UserNames(4) = "������"
    UserNames(5) = "��������"
    FlagMale = False

    Case "Sergey"
    UserNames(1) = "������"
    UserNames(2) = "�����"
    UserNames(3) = "������"
    UserNames(4) = "�����"
    UserNames(5) = "���������"
    FlagMale = True

    Case "Tatyana"
    UserNames(1) = "�������"
    UserNames(2) = "����"
    UserNames(3) = "������"
    UserNames(4) = "�������"
    UserNames(5) = "�������"
    FlagMale = False

    Case "Yuriy"
    UserNames(1) = "����"
    UserNames(2) = "���"
    UserNames(3) = "����"
    UserNames(4) = "�����"
    UserNames(5) = "������"
    FlagMale = True
    
    Case "Yuliya"
    UserNames(1) = "����"
    UserNames(2) = "���"
    UserNames(3) = "������"
    UserNames(4) = "���"
    UserNames(5) = "������"
    FlagMale = False

    Case "Jana"
    UserNames(1) = "���"
    UserNames(2) = "���"
    UserNames(3) = "������"
    UserNames(4) = "��������"
    UserNames(5) = "������"
    FlagMale = False

'�������� ����� ����� �������������

End Select

Exit Sub

ErrorHandler:
    
End Sub


'--------------------------------------------------
'strip the rest of the buffer
'--------------------------------------------------
Private Function StripTerm(strInput As String) As String
Dim ZeroPos As Long
StripTerm = Trim(strInput)
ZeroPos = InStr(StripTerm, Chr$(0))
If ZeroPos > 0 Then StripTerm = Left(StripTerm, ZeroPos - 1)
End Function
'--------------------------------------------------
'strip the rest of the buffer
'--------------------------------------------------
'--------------------------------------------------
'User Name
'--------------------------------------------------










