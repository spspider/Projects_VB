VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DrawAuto()

    StretchBlt wwww.hdc, auto.x, auto.y, 23, 31, zzzzz.hdc, 0, 0, 23, 31, SRCCOPY 'wwww.hdc - ���� ���������� (auto.x,auto.y) ���������� ������ ��� ����� ���������� ������;zzzzz.hdc - ������ ���������� ������ �����(0,0,23,31) �������� ���������� ������� � �����,������.

End Sub

Public Sub DrawOtherAuto(Num As Integer)

    Select Case Num ' ���������� ������ �����������

    Case 0 To 5

    StretchBlt picRoad.hdc, OtherAuto(Num).x, OtherAuto(Num).y, 23, 31, picSrc.hdc, Num * 23, 31, 23, 31, SRCCOPY

    Case 6 To 11

    StretchBlt picRoad.hdc, OtherAuto(Num).x, OtherAuto(Num).y, 23, 31, picSrc.hdc, (Num - 6) * 23, 62, 23, 31, SRCCOPY

    End Select '��� ������� � ������� �������

End Sub



