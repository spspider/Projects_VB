VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   Caption         =   "Winamp Remote Control"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   4335
   Icon            =   "formclient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictray 
      Height          =   315
      Left            =   3300
      ScaleHeight     =   255
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection"
      Height          =   555
      Left            =   60
      TabIndex        =   29
      Top             =   1620
      Width           =   4215
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   1995
         Picture         =   "formclient.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Disconect"
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   1500
         Picture         =   "formclient.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Connect"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   60
         TabIndex        =   30
         Text            =   "128.0.0.1"
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "C"
      Height          =   195
      Left            =   2580
      TabIndex        =   20
      ToolTipText     =   "Center Balance"
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   5
      Left            =   3330
      Picture         =   "formclient.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Repeat Play"
      Top             =   435
      Width           =   495
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   6
      Left            =   2835
      Picture         =   "formclient.frx":0EF7
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Shuffle Play"
      Top             =   435
      Width           =   495
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "VIS"
      Height          =   375
      Index           =   8
      Left            =   3825
      TabIndex        =   26
      ToolTipText     =   "Begin Visualization"
      Top             =   435
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Res. Winamp"
      Height          =   375
      Left            =   2325
      TabIndex        =   24
      Top             =   1215
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Min Winamp"
      Height          =   375
      Left            =   1230
      TabIndex        =   23
      Top             =   1215
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2220
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Clear Playlist"
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   17
      ToolTipText     =   "Clear Playlist"
      Top             =   1215
      Width           =   1230
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   10
      Left            =   3690
      Picture         =   "formclient.frx":1256
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Last Track"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   9
      Left            =   0
      Picture         =   "formclient.frx":15BE
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "First Track"
      Top             =   840
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2940
      Top             =   2340
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3660
      Top             =   3900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Browse for Music File or Playlist file"
      Height          =   3195
      Left            =   60
      TabIndex        =   10
      Top             =   2160
      Width           =   4215
      Begin VB.CommandButton cmdshowPlaylist 
         Height          =   375
         Left            =   1560
         Picture         =   "formclient.frx":1926
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Show and Refresh the Playlist on the Remote's Machine"
         Top             =   2700
         Width           =   375
      End
      Begin VB.CommandButton cmdfind 
         Height          =   375
         Left            =   1080
         Picture         =   "formclient.frx":1C7B
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Find Files on the Remote's Computer"
         Top             =   2700
         Width           =   435
      End
      Begin VB.ComboBox cmbdblclick 
         Height          =   315
         ItemData        =   "formclient.frx":1DC5
         Left            =   3060
         List            =   "formclient.frx":1DD2
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   600
         Picture         =   "formclient.frx":1DEC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Set the path to winamp's Excuteable file, do not use if this program already works."
         Top             =   2700
         Width           =   435
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   120
         Picture         =   "formclient.frx":2176
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Favorites"
         Top             =   2700
         Width           =   435
      End
      Begin VB.ComboBox cmbFiletype 
         Height          =   315
         ItemData        =   "formclient.frx":2700
         Left            =   3120
         List            =   "formclient.frx":271C
         TabIndex        =   16
         Text            =   "*.mp3"
         Top             =   1920
         Width           =   975
      End
      Begin VB.ListBox lstfiles 
         Height          =   1620
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   120
         Picture         =   "formclient.frx":27DA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Browse"
         Top             =   2280
         Width           =   435
      End
      Begin VB.TextBox txtselfile 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   2955
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "R"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   2280
         Width           =   435
      End
      Begin VB.CommandButton cmdaddfile 
         Height          =   375
         Left            =   600
         Picture         =   "formclient.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add File"
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Double Click:"
         Height          =   195
         Left            =   3060
         TabIndex        =   32
         Top             =   2280
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3420
      TabIndex        =   9
      Top             =   1215
      Width           =   885
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   4
      Left            =   2460
      Picture         =   "formclient.frx":31A6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Stop playing"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   3
      Left            =   1845
      Picture         =   "formclient.frx":34E5
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pause/Unpause Song"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   2
      Left            =   1230
      Picture         =   "formclient.frx":3831
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Play Current Song"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   1
      Left            =   3075
      Picture         =   "formclient.frx":3B81
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Next Track"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdActions 
      Height          =   375
      Index           =   0
      Left            =   615
      Picture         =   "formclient.frx":3EDF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Previous Track"
      Top             =   840
      Width           =   615
   End
   Begin ComctlLib.Slider sldVol 
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   480
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin ComctlLib.Slider sldBal 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123/344  2:28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Vol"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Bal"
      Height          =   255
      Left            =   1500
      TabIndex        =   7
      Top             =   480
      Width           =   315
   End
   Begin VB.Menu mnumenuFav 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuAddFav 
         Caption         =   "Add to Favorites"
      End
      Begin VB.Menu mnuOrgFav 
         Caption         =   "Organize Favorites"
      End
      Begin VB.Menu mnuFavSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavS 
         Caption         =   "Fav"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim winamppath As String
Dim PlNum As Long, PlTot As Long, songtime As Long, SongName As String
Private Function RemoveParent(ByVal File As String) As String
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveParent = Mid(File, t + 1)
    Exit Function
End If
Next t
End Function
Private Function RemoveFileName(ByVal File As String) As String
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveFileName = Left(File, t)
    Exit Function
End If
Next t
End Function


Private Sub cmdActions_Click(Index As Integer)
Select Case Index
Case 0
SendData "PREV"
Case 1
SendData "NEXT"
Case 2
SendData "PLAY"
Case 3
SendData "HALT"
Case 4
SendData "STOP"
Case 5
SendData "REPE"
Case 6
SendData "SHUF"
Case 7
SendData "CLER"
Case 8
SendData "VISA"
Case 9
SendData "GBEG"
Case 10
SendData "GEND"

End Select
End Sub

Private Sub cmdaddfile_Click()
SendData "AFLE" & txtselfile
End Sub

Private Sub cmdfind_Click()
frmfindfiles.FillinFields txtselfile
frmfindfiles.Visible = True
End Sub

Private Sub cmdshowPlaylist_Click()
SendData "GPLS"
frmplaylist.Visible = True
End Sub

Private Sub Command10_Click()
SendData "RESW"
End Sub

Private Sub Command11_Click()
PopupMenu mnumenuFav
End Sub

Private Sub cmdReplace_Click()
SendData "LFLE" & txtselfile
End Sub

Private Sub Command3_Click()
Dim x As String
x = InputBox("Please Locate Winamp." & vbCrLf & "DO NOT USE IF THIS PROGRAM ALREADY WORKS!!", "Winamp Location", winamppath)
If x = "" Then Exit Sub
winamppath = x
SendData "WAMP" & winamppath
End Sub

Private Sub Command4_Click()
SendData "CLSW"
End Sub

Private Sub cmdBrowse_Click()
If Right(txtselfile, 1) <> "\" Then
    MsgBox "Can't Browse a file!!!Use Add or Replace instead!", vbCritical, "Error"
    Exit Sub
End If
SendData "BRWD" & Replace(txtselfile & "\" & cmbFiletype.Text, "\\", "\")
End Sub

Private Sub Command6_Click()
Winsock1.Close
Winsock1.Connect txtIP, DefPort
End Sub

Private Sub Command7_Click()
Winsock1.Close
Winsock1.RemoteHost = ""
Winsock1.RemotePort = 0
Winsock1.LocalPort = 0
MsgBox "Disconnected!!"
End Sub

Private Sub Command8_Click()
sldBal.Value = 128
End Sub


Private Sub Command9_Click()
SendData "MINW"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbShiftMask Or vbAltMask Or vbCtrlMask Then
    If KeyCode = vbKeyC Then
        SendData "CHSS"
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim favs As Variant, setting() As String, t As Integer
winamppath = GetSetting("JDS INC", "REMOTEWINAMP", "WinampPath", "")
txtIP = GetSetting("JDS INC", "REMOTEWINAMP", "IP", "128.0.0.1")
cmbdblclick.ListIndex = GetSetting("JDS INC", "REMOTEWINAMP", "DoubleClick", 0)
load_icon pictray, Me.Icon, "Winamp Remote Control"
favs = GetAllSettings("JDS INC", "REMOTEWINAMP\Fav")
If Not IsEmpty(favs) Then
For t = LBound(favs, 1) To UBound(favs, 1)
    Load mnuFavS(mnuFavS.ubound + 1)
    setting = Split(favs(t, 1), "|", 2)
    mnuFavS(mnuFavS.ubound).Caption = setting(1)
    mnuFavS(mnuFavS.ubound).Tag = setting(0)
    mnuFavS(mnuFavS.ubound).Visible = True
Next t
End If
mnumenuFav.Visible = False
Load frmfindfiles
Load frmplaylist
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim t As Integer, I As Integer
SaveSetting "JDS INC", "REMOTEWINAMP", "WinampPath", winamppath
SaveSetting "JDS INC", "REMOTEWINAMP", "IP", txtIP
SaveSetting "JDS INC", "REMOTEWINAMP", "DoubleClick", cmbdblclick.ListIndex
For t = 0 To mnuFavS.ubound
    If mnuFavS(t).Visible Then
        I = I + 1
        SaveSetting "JDS INC", "REMOTEWINAMP\Fav", I, mnuFavS(t).Tag & "|" & mnuFavS(t).Caption
    End If
Next t
Unload_Icon pictray
End
End Sub

Private Sub lblInfo_Click()
Timer2.Enabled = Not Timer2.Enabled
If Timer2.Enabled = True Then SendData ("SNIN")
End Sub

Private Sub lstfiles_Click()
If Mid(lstfiles.Tag & lstfiles.List(lstfiles.ListIndex), 2) = ":\..\" Then
txtselfile = ""
Else
txtselfile = lstfiles.Tag & lstfiles.List(lstfiles.ListIndex)
End If
End Sub

Private Sub lstfiles_DblClick()
If Mid(lstfiles.Tag & lstfiles.List(lstfiles.ListIndex), 2) = ":\..\" Then
txtselfile = ""
Else
txtselfile = lstfiles.Tag & lstfiles.List(lstfiles.ListIndex)
End If
Select Case cmbdblclick.ListIndex
Case 0  'Browse
cmdBrowse_Click
Case 1 'Add
cmdaddfile_Click
Case 2 'Replace
cmdReplace_Click
End Select
End Sub

Private Sub mnuAddFav_Click()
On Error GoTo errhand
Dim x As String
x = InputBox("Please Label your Favorite:", "Label", txtselfile)
If x = "" Then Exit Sub
Load mnuFavS(mnuFavS.ubound + 1)
mnuFavS(mnuFavS.ubound).Caption = x
mnuFavS(mnuFavS.ubound).Tag = txtselfile
mnuFavS(mnuFavS.ubound).Visible = True
Exit Sub
errhand:
MsgBox "ERROR #" & Err.Number & " has occured in mnuAddFav_Click"
End Sub

Private Sub mnuFavS_Click(Index As Integer)
txtselfile = mnuFavS(Index).Tag
End Sub

Private Sub mnuOrgFav_Click()
frmorganizeFavs.Show
End Sub

Private Sub pictray_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo errhand
Dim msg As Long
    msg = x / Screen.TwipsPerPixelX
        Select Case msg
            Case WM_LBUTTONDBLCLK:
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP
            Me.Visible = Not Me.Visible
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
            Me.Visible = Not Me.Visible
        End Select
errhand:
End Sub

Private Sub sldVol_Change()
Call sldVol_Scroll
End Sub

Private Sub sldVol_Scroll()
SendData "VOLU" & sldVol.Value
End Sub
Private Sub sldBal_Change()
Call sldBal_Scroll
End Sub

Private Sub sldBal_Scroll()
'Dim prcnt As Long
SendData "BALN" & sldBal.Value
'prcnt = Int((sldBal.Value - 127) / 1.27)
'If prcnt = 0 Then
'    Timer1.Enabled = True
'Else
'    lrstatus = ""
'    Timer1.Enabled = False
'End If
End Sub


'Private Sub Timer1_Timer()
'If lrstatus.Caption <> "" Then
'lrstatus.Caption = ""
'Timer1.Enabled = False
'End If
'End Sub

Private Sub Timer2_Timer()
Static times As Byte
times = times + 1
If times = 20 Then
SendData ("SNIN")
times = 0
Else
If songtime > 0 Then
songtime = songtime + 500
lblInfo = (PlNum + 1) & "\" & PlTot & " " & cms(CSng(songtime / 1000)) & "   " & SongName
End If
End If
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
MsgBox "Disconnected!!"
End Sub

Private Sub Winsock1_Connect()
MsgBox "Connected!!"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errhand
Dim newdata As String, Linestart As Long, linedata As String
Static lastData As String
Winsock1.PeekData newdata, vbString
If InStr(1, newdata, Chr(1)) = 0 Then
    Exit Sub
End If
Winsock1.GetData newdata, vbString
'//Logdata newdata
again: Linestart = InStr(1, newdata, Chr(1))
If Linestart <> 0 Then
linedata = Mid(newdata, 1, Linestart - 1)
newdata = Mid(newdata, Linestart + 1)
ProcessData linedata
Else
lastData = newdata
newdata = ""
End If
If newdata <> "" And newdata <> Chr(1) Then GoTo again
Exit Sub
errhand:
MsgBox Err.Number & vbCrLf & Err.Description, , "Winsock1_DataArrival"
End Sub
Private Sub ProcessData(linedata As String)
Dim datatype As String, data As String
Dim t As Integer, shortname As String
Dim pathItems() As String
datatype = Left(linedata, 4)
data = Mid(linedata, 5)
Select Case UCase(datatype)
Case "DIRI" 'directory info
pathItems = Split(data, "|")
lstfiles.Clear
If IsNumeric(pathItems(0)) = True Then
lstfiles.Tag = ""
lstfiles.AddItem "ERROR#" & pathItems(0)
lstfiles.AddItem pathItems(1)
lstfiles.AddItem pathItems(2)
lstfiles.AddItem pathItems(3)
Else
lstfiles.Tag = pathItems(0)
For t = 1 To UBound(pathItems)
    lstfiles.AddItem pathItems(t)
Next t
End If
Case "SNIN"
'data=SongTime|TrkNum|#Trk|SongTitle
'where | is CHr(3)
pathItems = Split(data, Chr(3), 4)
lblInfo = (pathItems(1) + 1) & "\" & pathItems(2) & " " & cms(CSng(pathItems(0) / 1000)) & "   " & pathItems(3)
PlNum = pathItems(1)
PlTot = pathItems(2)
songtime = pathItems(0)
SongName = pathItems(3)
Case "CLRP"

Case "PLSE"

Case "FIND"
'data: filename(LB)Filename(LB)...
pathItems = Split(data, vbCrLf)
For t = 0 To UBound(pathItems)
frmfindfiles.lstfoundfiles.AddItem pathItems(t), 0
Next t
Case "GPLS"
'Data=Playlist Entry#1(LB)PLS Entry#2(LB)...
pathItems = Split(data, vbCrLf)
frmplaylist.lstpls.Clear
For t = 0 To UBound(pathItems)
    If pathItems(t) <> "" Then
        shortname = RemoveParent(pathItems(t))
        If t + 1 < 10 Then
            pathItems(t) = "   " & t + 1 & ".  " & pathItems(t)
        ElseIf t + 1 < 100 Then
            pathItems(t) = "  " & t + 1 & ".  " & pathItems(t)
        ElseIf t + 1 < 1000 Then
            pathItems(t) = " " & t + 1 & ".  " & pathItems(t)
        Else
            pathItems(t) = t + 1 & ".  " & pathItems(t)
        End If
        frmplaylist.lstpls.AddItem pathItems(t)
    
        If t + 1 < 10 Then
            pathItems(t) = "   " & t + 1 & ".  " & shortname
        ElseIf t + 1 < 100 Then
            pathItems(t) = "  " & t + 1 & ".  " & shortname
        ElseIf t + 1 < 1000 Then
            pathItems(t) = " " & t + 1 & ".  " & shortname
        Else
            pathItems(t) = t + 1 & ".  " & shortname
        End If
        frmplaylist.lstshortpls.AddItem pathItems(t)
    End If
Next t
End Select

End Sub


Sub SendData(data As String)
If Winsock1.State = sckConnected Then
Winsock1.SendData data & Chr(1)
End If
End Sub

