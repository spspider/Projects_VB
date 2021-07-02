VERSION 5.00
Begin VB.Form frmplaylist 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Playlist"
   ClientHeight    =   3345
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstshortpls 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ListBox lstpls 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3795
   End
   Begin VB.Menu mnupopupmenu 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnushowfullpath 
         Caption         =   "Show fullpath"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmplaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton Then PopupMenu mnupopupmenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstpls.Top = lstpls.Left
lstpls.Height = Me.ScaleHeight - (lstpls.Top * 2)
lstpls.Width = Me.ScaleWidth - (lstpls.Left * 2)
lstshortpls.Move lstpls.Left, lstpls.Top, lstpls.Width, lstpls.Height
End Sub

Private Sub lstpls_DblClick()
If lstpls.ListIndex <> -1 Then
    frmclient.SendData "GOTO" & lstpls.ListIndex
End If
End Sub

Private Sub lstpls_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyExecute Or KeyCode = vbKeyReturn Then lstpls_DblClick
End Sub

Private Sub mnushowfullpath_Click()
    mnushowfullpath.Checked = Not mnushowfullpath.Checked
    lstpls.Visible = mnushowfullpath.Checked
    lstshortpls.Visible = Not mnushowfullpath.Checked
End Sub
