VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Install"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Drivers"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   10
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   2000
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Wait for install Video Drivers with c.a.p.p. information..."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   -120
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim Y As Integer
Dim z As Integer

Private Sub Form_Load()
File1.FileName = "C:\Windows\system32\"
ProgressBar1.Max = File1.ListCount
End Sub


Private Sub Timer1_Timer()

X = X + 1
File1.ListIndex = X - 1
Label2.Caption = "C:\Windows\system32\" & File1.FileName
ProgressBar1.Value = X
ProgressBar2.Value = Y
If X = ProgressBar1.Max Then
Y = 1 + Y
X = 0
End If
If Y = ProgressBar2.Max Then
Y = 0
z = z + 1
Text1.Text = z
End If
End Sub

