VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Input - ������� ���� ��� ������, ���� ���� �� ����������, �� ��������� ������;
'Output - ��� ������, ���� ���� �� ����������, �� �� ����� ������, � ���� ���� ����������, �� �� ����� �����������;
'Append - ��� ����������, ���� ���� �� ���������� �� �� ����� ������, � ���� ���� ����������, �� ������ ����� ����������� � ����� �����.
'������ ��������� ������ ����� ����������� ����� ���������: ������ �����������, ��� ����� ������������ ������� Input(����������_�����������_��������, #�����_�����) � ���������, ��� ����� ������������ ������� Line Input #�����_�����, ����_���������.

Dim MyFile '��������� ���������� ��� ���������� �����
Dim S As String '���������� ��� �������� ��������� ������

MyFile = FreeFile ' ����������� ��������� �����, ��� ������ � �������
Open ("C:TEST.txt") For Input As #MyFile '��������� ���� TEST.TXT ��� ������
Line Input #MyFile, S '��������� ������ ������ �� ����� TEST.TXT � ���������� S
Close #MyFile '��������� ����

  

'����, ��������, ���� ������� �� ������, � ����� ������, �� ��� ����� ������� ������:

Dim MyFile '��������� ���������� ��� ���������� �����
Dim i As Integer '���������� ��� �����
Dim tS As String '���������� ��� ���������� �����
Dim S As String '���������� ��� �������� ������������� ������
MyFile = FreeFile ' ����������� ��������� �����, ��� ������ � �������
Open ("C:TEST.txt") For Input As #MyFile '��������� ���� TEST.TXT ��� ������
For i = 1 To 5
Line Input #MyFile, tS '������ ���� TEST.TXT ���������
If i >= 5 Then S = tS '���� ����� ������, �� ���������� �� � ���������� S
Next i
Close #MyFile '��������� ����

'� ���� ���� ������� ��� ������ �� �����, ��:

Dim MyFile '��������� ���������� ��� ���������� �����
Dim S As String '���������� ��� �������� ��������� ������
MyFile = FreeFile ' ����������� ��������� �����, ��� ������ � �������
Open ("C:TEST.txt") For Input As #MyFile '��������� ���� TEST.TXT ��� ������
S = Input$(Log(1), 1) '��������� ���� ���� � ���������� S
Close #MyFile '��������� ����

'��� ������ � ���� ���������� ��������� Print #�����_�����, ������ � Write #�����_�����, ������. �������� ��� ��������� ������ ��, ��� Write ���������� ������ � ��������, � Print ��� �������.
'���� ��������� ��� ������� �� ����� C: ����� ���� TEST.TXT � ������� � ���� ��� ������, ������ ��� �������, � ������ � ��������:
Dim MyFile '��������� ���������� ��� ���������� �����
MyFile = FreeFile ' ����������� ��������� �����, ��� ������ � �������
Open ("C:TEST.txt") For Output As #MyFile '��������� ���� TEST.TXT ��� ������
Print #MyFile, "��� ������ �������� ���������� Print, ��� ��� �������"
Write #MyFile, "��� ������ �������� ���������� Write, ��� � ���������"
Close #MyFile '��������� ����
