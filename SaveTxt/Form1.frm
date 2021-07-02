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
'Input - открыть файл для чтения, если файл не существует, то возникает ошибка;
'Output - для записи, если файл не существует, то он будет создан, а если файл существует, то он будет перезаписан;
'Append - для добавления, если файл не существует то он будет создан, а если файл существует, то данные будут добавляться в конец файла.
'Чтение текстовых файлов можно производить двумя способами: читать посимвольно, для этого используется функция Input(Количество_считываемых_символов, #Номер_файла) и построчно, для этого используется функция Line Input #Номер_файла, Куда_считывать.

Dim MyFile 'Объявляем переменную для свободного файла
Dim S As String 'Переменная для хранения считанных данных

MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("C:TEST.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
Line Input #MyFile, S 'Считываем первую строку из файла TEST.TXT в переменную S
Close #MyFile 'Закрываем файл

  

'Если, например, надо считать не первую, а пятую строку, то код будет немного другой:

Dim MyFile 'Объявляем переменную для свободного файла
Dim i As Integer 'Переменная для цикла
Dim tS As String 'Переменная для считывания строк
Dim S As String 'Переменная для хранения окончательных данных
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("C:TEST.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
For i = 1 To 5
Line Input #MyFile, tS 'Читаем файл TEST.TXT построчно
If i >= 5 Then S = tS 'Если пятая строка, то запоминаем ее в переменную S
Next i
Close #MyFile 'Закрываем файл

'А если надо считать все данные из файла, то:

Dim MyFile 'Объявляем переменную для свободного файла
Dim S As String 'Переменная для хранения считанных данных
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("C:TEST.txt") For Input As #MyFile 'Открываем файл TEST.TXT для чтения
S = Input$(Log(1), 1) 'Считываем весь файл в переменную S
Close #MyFile 'Закрываем файл

'Для записи в файл существуют операторы Print #Номер_файла, Данные и Write #Номер_файла, Данные. Отличает эти операторы только то, что Write записывает данные в кавычках, а Print без кавычек.
'Ниже следующий код создаст на диске C: новый файл TEST.TXT и запишет в него две строки, первую без кавычек, а вторую в кавычках:
Dim MyFile 'Объявляем переменную для свободного файла
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("C:TEST.txt") For Output As #MyFile 'Открываем файл TEST.TXT для записи
Print #MyFile, "Эта строка записана оператором Print, она без кавычек…"
Write #MyFile, "Эта строка записана оператором Write, она в кавычках…"
Close #MyFile 'Закрываем файл
