VERSION 5.00
Begin VB.Form frmorganizeFavs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstpaths 
      Height          =   1230
      Left            =   1140
      TabIndex        =   5
      Top             =   1140
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   435
      Left            =   2940
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdmovedown 
      Caption         =   "Move Down"
      Height          =   435
      Left            =   2940
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdmoveup 
      Caption         =   "Move Up"
      Height          =   435
      Left            =   2940
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
   End
End
Attribute VB_Name = "frmorganizeFavs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
If List1.ListIndex <> -1 Then
    List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub cmdmovedown_Click()
Dim sCaption As String, iPathPos As Integer
Dim NewPos As Integer
NewPos = List1.ListIndex + 1
If List1.ListIndex = List1.ListCount - 1 Then Exit Sub
sCaption = List1.Text
iPathPos = List1.ItemData(List1.ListIndex)
List1.RemoveItem List1.ListIndex
List1.AddItem sCaption, NewPos
List1.ItemData(NewPos) = iPathPos
List1.ListIndex = List1.NewIndex
End Sub

Private Sub cmdmoveup_Click()
Dim sCaption As String, iPathPos As Integer
Dim NewPos As Integer
NewPos = List1.ListIndex - 1
If List1.ListIndex = 0 Then Exit Sub
sCaption = List1.Text
iPathPos = List1.ItemData(List1.ListIndex)
List1.RemoveItem List1.ListIndex
List1.AddItem sCaption, NewPos
List1.ItemData(NewPos) = iPathPos
List1.ListIndex = List1.NewIndex
End Sub

Private Sub Form_Load()
Dim t As Integer
For t = 0 To frmclient.mnuFavS.ubound
If frmclient.mnuFavS(t).Visible Then
    List1.AddItem frmclient.mnuFavS(t).Caption
    lstpaths.AddItem frmclient.mnuFavS(t).Tag
    List1.ItemData(List1.NewIndex) = lstpaths.NewIndex
End If
Next t
End Sub

Private Sub Form_Unload(Cancel As Integer)
With frmclient
For I = 1 To .mnuFavS.ubound
    Unload .mnuFavS(I)
Next I
For I = 0 To List1.ListCount - 1
    Load .mnuFavS(.mnuFavS.ubound + 1)
    .mnuFavS(.mnuFavS.ubound).Caption = List1.List(I)
    .mnuFavS(.mnuFavS.ubound).Tag = lstpaths.List(List1.ItemData(I))
    .mnuFavS(.mnuFavS.ubound).Visible = True
Next I
End With

End Sub
