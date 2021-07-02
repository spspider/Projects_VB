VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2835
   Icon            =   "UnloadComp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "23:00:00"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q1 As String
Private Sub Form_Load()
On Error GoTo er:
q1 = GetSetting("SP\SPSoft\SP\UnloadComp", "Go", "Go", "HST1")
er:
End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
If Text1.Text = Text2.Text Then
    If q1 = Date Then
    Shell ("C:\windows\system32\shutdown.exe -s"), vbHide
    SaveSetting "SP\SPSoft\SP\UnloadComp", "Go", "Go", Date
    End If
End If
End Sub
