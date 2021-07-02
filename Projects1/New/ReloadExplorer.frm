VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1935
   Icon            =   "ReloadExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell ("C:\windows\system32\tskill explorer")
Shell ("C:\windows\system32\explorer")
Unload Me
End Sub
