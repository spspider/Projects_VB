VERSION 5.00
Begin VB.Form frmfindfiles 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Find Files"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear!"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox chksubfolders 
      Caption         =   "Check Sub-Folders"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   1755
   End
   Begin VB.ListBox lstfoundfiles 
      Height          =   2205
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3555
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "Find!"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtspec 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "T*.mp3"
      Top             =   300
      Width           =   2535
   End
   Begin VB.TextBox txtpath 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "C:\My Music\"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "File to look for:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Path to look in:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "frmfindfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfind_Click()
frmclient.SendData "FIND" & txtpath & "|" & txtspec & "|" & chksubfolders.Value
End Sub

Private Sub Command1_Click()
lstfoundfiles.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Visible = False
End If
End Sub
Public Function FillinFields(Filepath As String)
txtpath = Filepath
End Function

Private Sub Form_Resize()
On Error Resume Next
lstfoundfiles.Height = Me.ScaleHeight - lstfoundfiles.Top
lstfoundfiles.Width = Me.ScaleWidth - 240
End Sub
