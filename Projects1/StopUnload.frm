VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   Icon            =   "StopUnload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell ("C:\windows\system32\shutdown -a"), vbHide
Unload Me
End Sub
