VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Shutdown23.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
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
Dim InDate As String
Dim Path As String
Private Sub Form_Load()
InDate = GetSetting("Sp\Shutdown", "Date", "Newdate")
Form1.Visible = False
End Sub

Private Sub Timer1_Timer()
'Text1.Text = Time
 If InDate = Date Then
 Shell ("C:\windows\system32\shutdown -f"), vbHide 's
 End If
If Time = "23:00:00" Then
'SaveSetting "Sp\Shutdown", "Date", "Newdate", Date
Shell ("C:\windows\system32\shutdown -s -t 15"), vbHide 's
End If
End Sub
