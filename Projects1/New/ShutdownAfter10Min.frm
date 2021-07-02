VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2220
   Icon            =   "ShutdownAfter10Min.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2220
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim InDate As String

Private Sub Form_Load()
InDate = GetSetting("Sp\Shutdown", "Date", "Newdate")
n = 600
End Sub

Private Sub Timer1_Timer()
n = n - 1
If n < 1 Then
SaveSetting "Sp\Shutdown", "Date", "Newdate", Date
Shell ("C:\windows\system32\shutdown -s"), vbHide
End If
'If InDate = Date Then
'
'End If
End Sub
