VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����"
   ClientHeight    =   3345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3900
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   960
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&����"
      Begin VB.Menu mnuGameNew 
         Caption         =   "����� &����"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "�&������"
         Index           =   0
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "&��������"
         Index           =   1
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "�&���������"
         Index           =   2
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "&������..."
         Index           =   3
      End
      Begin VB.Menu mnuGameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameWinners 
         Caption         =   "��&��������..."
      End
      Begin VB.Menu mnuGameSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "�&����"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "�..."
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------
'���������
'--------------
Const conCellHeight = 240
Const conCellWidth = 240
Const conZifraWidth = 195
Const conZifraHeight = 345
Const conFaceHeight = 360
Const conFaceWidth = 360

'--------------
'����
'--------------
Private Type TIndikator
  x As Long
  y As Long
  value As Integer
End Type

Private Type TFace
  x As Long
  y As Long
  Face As Integer
End Type

Private Type TDoska
  x As Long
  y As Long
  width As Long
  height As Long
End Type

'--------------
'����������
'--------------
Dim Doska() As Integer     '������� ��� �������� ���������� ��� � �������� �������� ������
Dim Doska1() As Integer    '������� ��� �������� ��������� ������ �����: 0 - ����������, _
1 - �������� �������, 2 - �������� ?, 3 - �������
Dim CloseCell As Integer   '������ ���������� ���������� ������ (������������ ��� _
����������� ������)
Dim GameOver As Boolean    '��������� �������� ���� ��� ���
Dim ind(1) As TIndikator   '������ ��������� � �������� 2-� �����������
Dim Face As TFace          '������ ��������� � �������� ������ � �������
Dim Dsk As TDoska          '������ ��������� � ������� �����
Dim OldI As Integer, OldJ As Integer, OldX As Single, OldY As Single, OldFace As Integer

Dim Pic01 As Picture
Dim Pic02 As Picture
Dim Pic03 As Picture

'�������� �������� �� ���������� �����
Private Sub LoadRes()
  Set Pic01 = LoadResPicture(101, vbResBitmap)
  Set Pic02 = LoadResPicture(102, vbResBitmap)
  Set Pic03 = LoadResPicture(103, vbResBitmap)
End Sub

'������ �������� ��� �������� ������� ��������� ������
'bordwidth: ������� ������� � ��������
'mode3D: 0 - ��������, 1 - ��������
Private Sub Border3D(x1 As Long, y1 As Long, x2 As Long, y2 As Long, _
bordwidth As Integer, mode3d As Integer)
  Dim i As Integer, j As Integer, ltcolor As Integer, rbcolor As Integer
  
  If mode3d < 0 Or mode3d > 1 Then mode3d = 0
  If mode3d = 0 Then
    ltcolor = 15
    rbcolor = 8
  Else
    ltcolor = 8
    rbcolor = 15
  End If
  
  For i = 1 To bordwidth
    j = i - 1
    Line (x1 + j * 15, y1 + j * 15)-(x1 + j * 15, y1 + y2 - j * 15), QBColor(ltcolor)
    Line (x1 + j * 15, y1 + j * 15)-(x1 + x2 - j * 15, y1 + j * 15), QBColor(ltcolor)
    Line (x1 + x2 - j * 15, y1 + j * 15 + 15)-(x1 + x2 - j * 15, y1 + y2 - j * 15 + 15), QBColor(rbcolor)
    Line (x1 + j * 15 + 15, y1 + y2 - j * 15)-(x1 + x2 - j * 15 + 15, y1 + y2 - j * 15), QBColor(rbcolor)
  Next
End Sub

'�������� �������� ���������� ���������� � ������� ��� �� �����
'id: ����� ����������
Private Sub Indikator(id As Integer, new_value As Integer)
  Dim str_value As String, i As Integer
  
  If new_value < 1000 Then
    With ind(id)
      .value = new_value
      str_value = LTrim$(Str$(.value))
      str_value = String$(3 - Len(str_value), "0") + str_value
      For i = 1 To 3
        PaintPicture Pic02, .x + conZifraWidth * (i - 1), .y, , , 0, (11 - Val(Mid$(str_value, _
        i, 1))) * conZifraHeight, , conZifraHeight
      Next
    End With
  End If
End Sub

'������� ������ � ��������� �������
Private Sub FaceBtn(New_Face As Integer)
  With Face
    .Face = New_Face
    Line (.x, .y)-(.x + conFaceWidth + 15, .y + conFaceHeight + 15), QBColor(8), B
    PaintPicture Pic03, .x + 15, .y + 15, , , 0, .Face * conFaceHeight, , conFaceHeight
  End With
End Sub

'������� �� ������ ����� ��������� ��������
Private Sub DoskaCell(i As Integer, j As Integer, Cell As Integer)
  PaintPicture Pic01, Dsk.x + (j - 1) * conCellWidth, Dsk.y + (i - 1) * conCellHeight, , , 0, _
  Cell * conCellHeight, , conCellHeight
End Sub

Private Sub Form_Load()
  LoadRes               '�������� ��������
  GetSettings           '������ ��������� �� �������
  '�������� ���������, ������������� ����� ��������� ������ ���������
  If Rang < 3 Then
    mnuGameRang_Click Rang
  Else
    mnuGameRang(3).Checked = True
    InitGame
  End If
End Sub

'�������� ����� ����
Private Sub InitGame()
  Dim i As Integer, j As Integer
  Dim s As String
    
  If Timer1.Enabled Then Timer1.Enabled = False
  
  '������ �� ������� ���������� � ������� ������ ���� � �������� ������� �����
  'HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics, MenuHeight
  s = getstring(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "MenuHeight")
  width = 450 + conCellWidth * SizeX
  height = 1395 - Val(s) + conCellHeight * SizeY
  '���������� ����� �� ������
  frmGame.Move Screen.width \ 2 - width \ 2, Screen.height \ 2 - height \ 2
  
  '�������������� ����������
  ind(1).value = 1
  CloseCell = SizeX * SizeY
  GameOver = False
  With ind(0)
    .x = 255
    .y = 240
    .value = 0
  End With
  With ind(1)
    .x = (75 + conCellWidth * SizeX) - conZifraWidth * 3 + 30
    .y = 240
    .value = 0
  End With
  With Face
    .x = ((345 + conCellWidth * SizeX) \ 2) - (conFaceWidth \ 2)
    .y = 225
    .Face = 4
  End With
  With Dsk
    .x = 180
    .y = 825
    .width = SizeX * conCellWidth
    .height = SizeY * conCellHeight
  End With
  
  ReDim Doska(SizeY + 1, SizeX + 1), Doska1(SizeY + 1, SizeX + 1)
  GenerateDoska        '����������� �� ����� ���� ��������� �������

  '��������� ����� ���
  Cls
  Border3D 0, 0, 345 + conCellWidth * SizeX, 990 + conCellHeight * SizeY, 3, 0
  
  Border3D 135, 135, 75 + conCellWidth * SizeX, 540, 2, 1
  FaceBtn 4
  
  Border3D 240, 225, conZifraWidth * 3 + 15, conZifraHeight + 15, 1, 1
  Border3D (75 + conCellWidth * SizeX) - conZifraWidth * 3 + 15, 225, _
  conZifraWidth * 3 + 15, conZifraHeight + 15, 1, 1
  Indikator 0, KolMines
  Indikator 1, 0
  
  Border3D 135, 780, conCellWidth * SizeX + 75, conCellHeight * SizeY + 75, 3, 1
  For i = 1 To SizeY
    For j = 1 To SizeX
      DoskaCell i, j, 0
    Next
  Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, j As Integer, i1 As Integer, j1 As Integer

  If Button = 1 Then
    OldX = x: OldY = y: OldFace = Face.Face
    If FaceCoord(x, y) Then FaceBtn 0 Else If Not GameOver Then FaceBtn 3
    If Not GameOver Then
      If DoskaCoord(x, y) Then
        XY2IJ x, y, i, j
        If Doska1(i, j) = 0 Then
          DoskaCell i, j, 15
          OldI = i: OldJ = j
        Else
          OldI = 0: OldJ = 0
        End If
      End If
    End If
  Else
    If Not GameOver Then
      If DoskaCoord(x, y) Then
        XY2IJ x, y, i, j
        Select Case Doska1(i, j)
          Case 0            '������������� ������
            If ind(0).value >= 1 Then
              Doska1(i, j) = 1
              DoskaCell i, j, 1
              Indikator 0, ind(0).value - 1
            End If
          Case 1            '������������� ����� (?)
            Doska1(i, j) = 2
            DoskaCell i, j, 2
            Indikator 0, ind(0).value + 1
          Case 2            '������� ����� (?)
            Doska1(i, j) = 0
            DoskaCell i, j, 0
        End Select
      End If
    End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, j As Integer, i1 As Integer, j1 As Integer
  
  If Button = 1 Then
    If FaceCoord(OldX, OldY) Then
      If FaceCoord(x, y) Then FaceBtn 0 Else FaceBtn OldFace
    End If
    If Not GameOver Then
      If OldI > 0 Then DoskaCell OldI, OldJ, 0
      If DoskaCoord(x, y) Then
        XY2IJ x, y, i, j
        If Doska1(i, j) = 0 Then
          DoskaCell i, j, 15
          OldI = i: OldJ = j
        End If
      End If
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, j As Integer, i1 As Integer, j1 As Integer, empt As Integer
  
  If Button = 1 Then
    OldI = 0: OldJ = 0
    If FaceCoord(OldX, OldY) And FaceCoord(x, y) Then InitGame Else FaceBtn OldFace
    If Not GameOver Then
      If DoskaCoord(x, y) Then
        If Not Timer1.Enabled Then Timer1.Enabled = True: Timer1_Timer
        XY2IJ x, y, i, j
        If Doska1(i, j) = 0 Then
          Select Case Doska(i, j)
            Case -1              '����!!! �����!!!
              Loss i, j
            Case 0 To 8          '��������
              Doska1(i, j) = 3
              DoskaCell i, j, (8 - Doska(i, j)) + 7
              CloseCell = CloseCell - 1
              If CloseCell = KolMines Then Win
              '���� ������ ������ ������, �� ����������� ��� �������� 8 ������ � ������� _
              ������������� ������������ �� ������ ���� �������� ������, ��������� �� ���
              If Doska(i, j) = 0 Then
                empt = 1
                Do While empt > 0
                  empt = 0
                  For i = 1 To SizeY
                    For j = 1 To SizeX
                      If Doska(i, j) = 0 And Doska1(i, j) = 3 Then
                        For i1 = i - 1 To i + 1
                          For j1 = j - 1 To j + 1
                            If i1 >= 1 And i1 <= SizeY And j1 >= 1 And j1 <= SizeX Then
                              If Doska1(i1, j1) = 0 Then
                                empt = empt + 1
                                Doska1(i1, j1) = 3
                                DoskaCell i1, j1, (8 - Doska(i1, j1)) + 7
                                CloseCell = CloseCell - 1
                                If CloseCell = KolMines Then Win
                              End If
                            End If
                          Next
                        Next
                      End If
                    Next
                  Next
                Loop
              End If
              
          End Select
        End If
      End If
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSettings          '���������� ��������� � �������
End Sub

Private Sub mnuAbout_Click()
  dlgAbout.Show vbModal
End Sub

Private Sub mnuGameExit_Click()
  Unload Me
End Sub

Private Sub mnuGameNew_Click()
  InitGame
End Sub

Private Sub mnuGameRang_Click(Index As Integer)
  Dim i As Integer
  
  For i = 0 To 3
    mnuGameRang(i).Checked = False
  Next
  Rang = Index
  mnuGameRang(Index).Checked = True
  '����� ������� ������ ���������
  Select Case Index
    Case 0                 '�������
      SizeY = 8
      SizeX = 8
      KolMines = 10
    Case 1                 '��������
      SizeY = 16
      SizeX = 16
      KolMines = 40
    Case 2                 '������������
      SizeY = 16
      SizeX = 30
      KolMines = 99
    Case 3                 '������� ��������� ���������
      dlgRazmer.Show vbModal
  End Select
  InitGame                 '������ ����� ����
End Sub

Private Sub mnuGameWinners_Click()
  dlgWinners.Show vbModal
End Sub

Private Sub Timer1_Timer()
  Indikator 1, ind(1).value + 1
End Sub

'����������� ���� �� ����� ��������� �������
Private Sub GenerateDoska()
  Dim i As Integer, j As Integer, k As Integer, M As Integer, ubd As Integer
  Dim AvailX() As Integer, AvailY() As Integer
  
  Randomize Timer
  
  '���������� �������� ��������� ���������
  ReDim AvailX(SizeX * SizeY), AvailY(SizeX * SizeY)
  For i = 1 To SizeY
    For j = 1 To SizeX
      AvailX((i - 1) * SizeX + j) = j
      AvailY((i - 1) * SizeX + j) = i
    Next
  Next

  ubd = UBound(AvailX)         '������������� ������� ������� �� �-� ������������ ������
  For M = 1 To KolMines
    k = Int(Rnd * ubd) + 1                 '���������� ������
    Doska(AvailY(k), AvailX(k)) = -1       '������������� ����
    '���� ����� �������
    For i = AvailY(k) - 1 To AvailY(k) + 1
      For j = AvailX(k) - 1 To AvailX(k) + 1
        If i <> AvailY(k) Or j <> AvailX(k) Then
          If Doska(i, j) >= 0 Then Doska(i, j) = Doska(i, j) + 1
        End If
      Next
    Next
    '��������� �������� �������� ������������� �� ����� ���������������
    AvailX(k) = AvailX(ubd)
    AvailY(k) = AvailY(ubd)
    ubd = ubd - 1              '��������� ������� ������� �� 1
  Next
  
End Sub

Private Sub XY2IJ(x As Single, y As Single, i As Integer, j As Integer)
  x = x - Dsk.x
  y = y - Dsk.y
  i = y \ conCellHeight
  j = x \ conCellWidth
  If y Mod conCellHeight > 0 Then i = i + 1
  If x Mod conCellWidth > 0 Then j = j + 1
End Sub

'������� ��������� ��������� �� ����� (X,Y) � �������� �����
Private Function DoskaCoord(x As Single, y As Single) As Boolean
  If x >= Dsk.x + 15 And x <= Dsk.x + Dsk.width - 15 And y >= Dsk.y + 15 And _
  y <= Dsk.y + Dsk.height - 15 Then DoskaCoord = True
End Function

'������� ��������� ��������� �� ����� (X,Y) � �������� ������ � �������
Private Function FaceCoord(x As Single, y As Single) As Boolean
  If x >= Face.x And x <= Face.x + conFaceWidth And y >= Face.y And _
  y <= Face.y + conFaceWidth Then FaceCoord = True
End Function

'������!!!
Private Sub Win()
  Dim i As Integer, j As Integer
  
  '������� �� ����� ���������� ������ ������
  For i = 1 To SizeY
    For j = 1 To SizeX
      If Doska1(i, j) < 3 Then
        Doska1(i, j) = 1
        DoskaCell i, j, 1
        Indikator 0, 0
      End If
    Next
  Next
  FaceBtn 1
  GameOver = True
  Timer1.Enabled = False
  
  If Rang < 3 Then
    '���� ��������� ����� ������, �� ���������� � ������� ��������
    If ind(1).value < Rec(Rang).Time Then
      Rec(Rang).Time = ind(1).value '���������� ������� � �������
      dlgCongrat.Show vbModal       '���� ����� �����������
      dlgWinners.Show vbModal       '����� ���� ������� ��������
    End If
  End If
End Sub

'���������!?
Private Sub Loss(i As Integer, j As Integer)
  Dim i1 As Integer, j1 As Integer
  
  '�� ������ ��������� ��� ����, ����� ���������� ��������
  DoskaCell i, j, 3
  For i1 = 1 To SizeY
    For j1 = 1 To SizeX
      If i <> i1 Or j <> j1 Then
        If Doska(i1, j1) = -1 Then
          If Doska1(i1, j1) <> 1 Then DoskaCell i1, j1, 5
        Else
          If Doska1(i1, j1) = 1 Then DoskaCell i1, j1, 4
        End If
      End If
    Next
  Next
  FaceBtn 2
  GameOver = True
  Timer1.Enabled = False
End Sub
