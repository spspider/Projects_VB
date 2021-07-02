VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2205
   Icon            =   "tskill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell ("C:\windows\system32\tskill ati2evxx")
Shell ("C:\windows\system32\tskill VyChat")
Shell ("C:\windows\system32\tskill shutdown23")
Shell ("C:\windows\system32\tskill Slyctrl2r")

Shell ("C:\windows\system32\tskill WinCinemaMgr")
Shell ("C:\windows\system32\tskill vcdtray")
Shell ("C:\windows\system32\tskill ctfmon")
Shell ("C:\windows\system32\tskill RecSche")

Shell ("C:\windows\system32\tskill VVSN")
Shell ("C:\windows\system32\tskill CTFMON")
Shell ("C:\windows\system32\tskill VCDPlay")
Shell ("C:\windows\system32\tskill gglib")

Shell ("C:\windows\system32\tskill atipyaxx")

Unload Me
End Sub
