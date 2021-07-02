VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сапер"
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
      Caption         =   "&Игра"
      Begin VB.Menu mnuGameNew 
         Caption         =   "Новая &игра"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "Н&овичок"
         Index           =   0
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "&Любитель"
         Index           =   1
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "С&пециалист"
         Index           =   2
      End
      Begin VB.Menu mnuGameRang 
         Caption         =   "&Другой..."
         Index           =   3
      End
      Begin VB.Menu mnuGameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameWinners 
         Caption         =   "По&бедители..."
      End
      Begin VB.Menu mnuGameSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "В&ыход"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "О..."
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------
'Константы
'--------------
Const conCellHeight = 240
Const conCellWidth = 240
Const conZifraWidth = 195
Const conZifraHeight = 345
Const conFaceHeight = 360
Const conFaceWidth = 360

'--------------
'Типы
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
'Переменные
'--------------
Dim Doska() As Integer     'Матрица для хранения растановки мин и значений соседних клеток
Dim Doska1() As Integer    'Матрица для хранения состояния клеток доски: 0 - неоткрытая, _
1 - помечена флажком, 2 - помечена ?, 3 - открыта
Dim CloseCell As Integer   'Хранит количество неоткрытых клеток (используется для _
определения победы)
Dim GameOver As Boolean    'Указывает окончена игра или нет
Dim ind(1) As TIndikator   'Хранит положения и значения 2-х индикаторов
Dim Face As TFace          'Хранит положение и значение кнопки с рожицей
Dim Dsk As TDoska          'Хранит положение и размеры доски
Dim OldI As Integer, OldJ As Integer, OldX As Single, OldY As Single, OldFace As Integer

Dim Pic01 As Picture
Dim Pic02 As Picture
Dim Pic03 As Picture

'Загрузка картинок из ресурсного файла
Private Sub LoadRes()
  Set Pic01 = LoadResPicture(101, vbResBitmap)
  Set Pic02 = LoadResPicture(102, vbResBitmap)
  Set Pic03 = LoadResPicture(103, vbResBitmap)
End Sub

'Рисует выпуклую или вогнутую границу указанной ширины
'bordwidth: толщина границы в пикселях
'mode3D: 0 - выпуклая, 1 - вогнутая
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

'Изменяет значение выбранного индикатора и выводит его на экран
'id: номер индикатора
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

'Выводит кнопку с указанной рожицей
Private Sub FaceBtn(New_Face As Integer)
  With Face
    .Face = New_Face
    Line (.x, .y)-(.x + conFaceWidth + 15, .y + conFaceHeight + 15), QBColor(8), B
    PaintPicture Pic03, .x + 15, .y + 15, , , 0, .Face * conFaceHeight, , conFaceHeight
  End With
End Sub

'Выводит на клетке доски указанную картинку
Private Sub DoskaCell(i As Integer, j As Integer, Cell As Integer)
  PaintPicture Pic01, Dsk.x + (j - 1) * conCellWidth, Dsk.y + (i - 1) * conCellHeight, , , 0, _
  Cell * conCellHeight, , conCellHeight
End Sub

Private Sub Form_Load()
  LoadRes               'Загрузка картинок
  GetSettings           'Чтение установок из реестра
  'Загрузка установок, сохранившихся после последней работы программы
  If Rang < 3 Then
    mnuGameRang_Click Rang
  Else
    mnuGameRang(3).Checked = True
    InitGame
  End If
End Sub

'Начинает новую игру
Private Sub InitGame()
  Dim i As Integer, j As Integer
  Dim s As String
    
  If Timer1.Enabled Then Timer1.Enabled = False
  
  'Читаем из реестра информацию о текущей высоте меню и изменяем размеры формы
  'HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics, MenuHeight
  s = getstring(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "MenuHeight")
  width = 450 + conCellWidth * SizeX
  height = 1395 - Val(s) + conCellHeight * SizeY
  'Центрируем форму на экране
  frmGame.Move Screen.width \ 2 - width \ 2, Screen.height \ 2 - height \ 2
  
  'Инициализируем переменные
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
  GenerateDoska        'Располагаем на доске мины случайным образом

  'Формируем общий вид
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
          Case 0            'Устанавливаем флажок
            If ind(0).value >= 1 Then
              Doska1(i, j) = 1
              DoskaCell i, j, 1
              Indikator 0, ind(0).value - 1
            End If
          Case 1            'Устанавливаем метку (?)
            Doska1(i, j) = 2
            DoskaCell i, j, 2
            Indikator 0, ind(0).value + 1
          Case 2            'Снимаем метку (?)
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
            Case -1              'Мина!!! БАБАХ!!!
              Loss i, j
            Case 0 To 8          'Пронесло
              Doska1(i, j) = 3
              DoskaCell i, j, (8 - Doska(i, j)) + 7
              CloseCell = CloseCell - 1
              If CloseCell = KolMines Then Win
              'Если открыл пустую клетку, то открываются все соседние 8 клеток и процесс _
              автоматически продолжается до выбора всех соседних клеток, свободных от мин
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
  SaveSettings          'Сохранение установок в реестре
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
  'Выбор нужного уровня сложности
  Select Case Index
    Case 0                 'Новичок
      SizeY = 8
      SizeX = 8
      KolMines = 10
    Case 1                 'Любитель
      SizeY = 16
      SizeX = 16
      KolMines = 40
    Case 2                 'Профессионал
      SizeY = 16
      SizeX = 30
      KolMines = 99
    Case 3                 'Вручную выбираешь установки
      dlgRazmer.Show vbModal
  End Select
  InitGame                 'Начало новой игры
End Sub

Private Sub mnuGameWinners_Click()
  dlgWinners.Show vbModal
End Sub

Private Sub Timer1_Timer()
  Indikator 1, ind(1).value + 1
End Sub

'Располагает мины на доске случайным образом
Private Sub GenerateDoska()
  Dim i As Integer, j As Integer, k As Integer, M As Integer, ubd As Integer
  Dim AvailX() As Integer, AvailY() As Integer
  
  Randomize Timer
  
  'Заполнение массивов доступных координат
  ReDim AvailX(SizeX * SizeY), AvailY(SizeX * SizeY)
  For i = 1 To SizeY
    For j = 1 To SizeX
      AvailX((i - 1) * SizeX + j) = j
      AvailY((i - 1) * SizeX + j) = i
    Next
  Next

  ubd = UBound(AvailX)         'Устанавливаем верхнюю границу до к-й генерируется индекс
  For M = 1 To KolMines
    k = Int(Rnd * ubd) + 1                 'Генерируем индекс
    Doska(AvailY(k), AvailX(k)) = -1       'Устанавливаем мину
    'Цикл счета соседей
    For i = AvailY(k) - 1 To AvailY(k) + 1
      For j = AvailX(k) - 1 To AvailX(k) + 1
        If i <> AvailY(k) Or j <> AvailX(k) Then
          If Doska(i, j) >= 0 Then Doska(i, j) = Doska(i, j) + 1
        End If
      Next
    Next
    'Последние элементы массивов устанавливаем на место сгенерированных
    AvailX(k) = AvailX(ubd)
    AvailY(k) = AvailY(ubd)
    ubd = ubd - 1              'Уменьшаем верхнюю границу на 1
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

'Функция проверяет находится ли точка (X,Y) в пределах доски
Private Function DoskaCoord(x As Single, y As Single) As Boolean
  If x >= Dsk.x + 15 And x <= Dsk.x + Dsk.width - 15 And y >= Dsk.y + 15 And _
  y <= Dsk.y + Dsk.height - 15 Then DoskaCoord = True
End Function

'Функция проверяет находится ли точка (X,Y) в пределах кнопки с рожицей
Private Function FaceCoord(x As Single, y As Single) As Boolean
  If x >= Face.x And x <= Face.x + conFaceWidth And y >= Face.y And _
  y <= Face.y + conFaceWidth Then FaceCoord = True
End Function

'Победа!!!
Private Sub Win()
  Dim i As Integer, j As Integer
  
  'Выводим на месте неоткрытых клеток флажки
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
    'Если установил новый рекорд, то заносишься в таблицу рекордов
    If ind(1).value < Rec(Rang).Time Then
      Rec(Rang).Time = ind(1).value 'Сохранение рекорда в таблице
      dlgCongrat.Show vbModal       'Ввод имени рекордсмена
      dlgWinners.Show vbModal       'Вывод всей таблицы рекордов
    End If
  End If
End Sub

'Поражение!?
Private Sub Loss(i As Integer, j As Integer)
  Dim i1 As Integer, j1 As Integer
  
  'На экране выводятся все мины, кроме отмеченных флажками
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
