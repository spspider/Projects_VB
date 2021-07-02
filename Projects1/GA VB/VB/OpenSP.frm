VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OCX"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2040
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   2040
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo er1:
FileCopy "data\OCX\MCI32.OCX", "C:\WINDOWS\system32\MCI32.OCX"
FileCopy "data\OCX\COMDLG32.OCX", "C:\WINDOWS\system32\COMDLG32.OCX"
FileCopy "data\OCX\MSCOMCT2.OCX", "C:\WINDOWS\system32\MSCOMCT2.OCX"
FileCopy "data\OCX\MSCOMM32.OCX", "C:\WINDOWS\system32\MSCOMM32.OCX"
er1:
Shell "sp.exe", vbNormalFocus
Unload Me
End Sub
