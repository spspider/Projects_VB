VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1785
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   255
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   1320
   End
   Begin VB.PictureBox RightB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   3360
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   2280
      Width           =   105
   End
   Begin VB.PictureBox LeftB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   3000
      Width           =   105
   End
   Begin VB.PictureBox TB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      Begin VB.PictureBox BTN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   3
         Left            =   1110
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   7
         Top             =   45
         Width           =   330
      End
      Begin VB.PictureBox BTN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   750
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   6
         Top             =   45
         Width           =   330
      End
      Begin VB.PictureBox BTN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   390
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   5
         Top             =   45
         Width           =   330
      End
      Begin VB.PictureBox BTN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   30
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   4
         Top             =   45
         Width           =   330
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BoTN(0 To 3)

Private Sub BTN_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BoTN(Index) = 1
End Sub


Private Sub BTN_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
BoTN(Index) = 1
End If
End Sub

Private Sub BTN_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BoTN(Index) = 0

If X >= 0 And Y >= 0 And X <= 22 And Y <= 11 Then
If Index = 0 Then
Form2.List1.Clear
Surce.List1.Clear
End If

If Index = 1 Then
    With CommonDialog1
        .DialogTitle = "Open List"
        .CancelError = False
        .Filter = "SPL|*.SPL|"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
  LoadList sFile
End If

If Index = 2 Then
    With CommonDialog1
        .DialogTitle = "Save List"
        .CancelError = False
        .FileName = Form1.Caption
        .Filter = "SPL|*.SPL|"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveList sFile
End If

If Index = 3 Then
N = List1.ListIndex
List1.RemoveItem N
Surce.List1.RemoveItem N
End If

End If
End Sub

Private Sub Form_Activate()
Surce.Label1.Caption = 1
End Sub

Private Sub Form_Load()
Form2.Width = TB.Width * 15
For N = 0 To 3
BoTN(N) = 0
Next
LoadList App.Path & "\Defoult.SPL"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveList App.Path & "\Defoult.SPL"
End Sub

Private Sub Form_Resize()
LeftB.Top = Form2.Height / 15 - LeftB.Height
RightB.Left = Form2.Width / 15 - RightB.Width
RightB.Top = Form2.Height / 15 - LeftB.Height
TB.Width = Form1.Width
DarwPC
DarwTB
List1.Width = Form2.Width / 15 - 14
List1.Height = Form2.Height / 15 - 24
End Sub

Sub DarwPC()
For X = 0 To Form2.Width / 15 Step 2
BitBlt Form2.hdc, X, Form2.Height / 15 - 7, 2, 7, Surce.PP.hdc, 0, 14, vbSrcCopy
Next

For Y = 0 To Form2.Height / 15 Step 2
BitBlt Form2.hdc, 0, Y, 7, 2, Surce.PP.hdc, 0, 0, vbSrcCopy
Next

For Y = 0 To Form2.Height / 15 Step 2
BitBlt Form2.hdc, Form2.Width / 15 - 7, Y, 7, 2, Surce.PP.hdc, 0, 7, vbSrcCopy
Next

Form2.Refresh
End Sub


Sub DarwTB()

For X = 0 To Form2.Width / 15 Step 2
BitBlt TB.hdc, X, 0, 2, 17, Surce.PLST2.hdc, 0, 17, vbSrcCopy
Next
BitBlt TB.hdc, 0, 0, 6, 17, Surce.PLST2.hdc, 0, 0, vbSrcCopy
BitBlt TB.hdc, Form2.Width / 15 - 6, 0, 6, 17, Surce.PLST2.hdc, 0, 34, vbSrcCopy

BitBlt TB.hdc, Form2.Width / 15 / 2 - Surce.PLST.Width / 2, 0, Surce.PLST.Width, Surce.PLST.Height, Surce.PLST.hdc, 0, 0, vbSrcCopy

TB.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveList App.Path & "\Defoult.SPL"
End Sub

Private Sub List1_Click()
Surce.List1.ListIndex = List1.ListIndex
Form1.Label1.Caption = List1.text
L1 = Len(Surce.List1.text)
L2 = Len(List1.text)
Form1.Label2.Caption = Mid$(Surce.List1.text, 1, L1 - L2)
End Sub

Private Sub List1_DblClick()
Surce.List1.ListIndex = List1.ListIndex
Form1.Label1.Caption = List1.text
L1 = Len(Surce.List1.text)
L2 = Len(List1.text)
Form1.Label2.Caption = Mid$(Surce.List1.text, 1, L1 - L2)

MClose
OpenFirst Form1.Label2.Caption & Form1.Label1.Caption
PlayFirst
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

txt$ = Data.Files.Item(1)
R$ = BigChr$(Rsh(txt$))
If R$ = ".MP3" Or R$ = ".WAV" Or R$ = ".WMA" Or R$ = ".MID" Or R$ = ".WMV" Then
List1.AddItem FileName(txt$), 0
Surce.List1.AddItem (txt$), 0
Else
MsgBox ("This format is not supported"), , "SoundPlayer"
End If

End Sub

Private Sub RightB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RightB.CurrentX = X
RightB.CurrentY = Y
End Sub

Private Sub RightB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
W = Form2.Width + X - RightB.CurrentX
If W < 80 * 15 Then W = 80 * 15

H = Form2.Height + Y - RightB.CurrentY
If H < Form1.Height Then H = Form1.Height

Form2.Width = W
Form2.Height = H
End If
End Sub

Private Sub TB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TB.CurrentX = X
TB.CurrentY = Y
End Sub

Private Sub TB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = Form2.Left + X - TB.CurrentX
B = Form2.Top + Y - TB.CurrentY
nx = Form1.Left + Form1.Width
nx2 = Form1.Left
ny = Form1.Top + Form1.Height
ny2 = Form1.Top

If RZN(a, nx2) < 2 Then a = nx2
If RZN(B, ny2) < 2 Then B = ny2

If RZN(a / 15, nx / 15) < 2 Then a = nx
If RZN(nx2 / 15, (a + Form2.Width) / 15) < 2 Then a = Form1.Left - Form2.Width
If RZN(B / 15, ny / 15) < 2 Then B = ny
If RZN(ny2 / 15, (B + Form2.Height) / 15) < 2 Then B = Form1.Top - Form2.Height



Form2.Left = a
Form2.Top = B
End If
End Sub


Private Function Rsh(txt$)
Lg = Len(txt$)
For N = 1 To Lg
D$ = Mid$(txt$, Lg - N, 1)
If D$ = "." Then GoTo 2
If N > 5 Then GoTo 2
Next
2
Rsh = Mid$(txt$, Lg - N, Lg - N)
End Function

Private Function FileName(txt$)
Lg = Len(txt$)
For N = 1 To Lg
D$ = Mid$(txt$, Lg - N + 1, 1)
If D$ = "\" Then GoTo 2
Next
2
FileName = Mid$(txt$, Lg - N + 2, Lg)
End Function

Private Function BigChr$(text$)
txt$ = ""
Lg = Len(text$)
For N = 1 To Lg
D$ = Mid$(text$, N, 1)
S$ = D$
If D$ = "a" Then S$ = "A"
If D$ = "b" Then S$ = "B"
If D$ = "c" Then S$ = "C"
If D$ = "d" Then S$ = "D"
If D$ = "e" Then S$ = "E"
If D$ = "f" Then S$ = "F"
If D$ = "g" Then S$ = "G"
If D$ = "h" Then S$ = "H"
If D$ = "i" Then S$ = "I"
If D$ = "j" Then S$ = "J"
If D$ = "k" Then S$ = "K"
If D$ = "l" Then S$ = "L"
If D$ = "m" Then S$ = "M"
If D$ = "n" Then S$ = "N"
If D$ = "o" Then S$ = "O"
If D$ = "p" Then S$ = "P"
If D$ = "q" Then S$ = "Q"
If D$ = "r" Then S$ = "R"
If D$ = "s" Then S$ = "S"
If D$ = "t" Then S$ = "T"
If D$ = "u" Then S$ = "U"
If D$ = "v" Then S$ = "V"
If D$ = "w" Then S$ = "W"
If D$ = "x" Then S$ = "X"
If D$ = "y" Then S$ = "Y"
If D$ = "z" Then S$ = "Z"
txt$ = txt$ & S$
Next
BigChr$ = txt$
End Function

Private Sub Timer1_Timer()
For N = 0 To 3
BitBlt BTN(N).hdc, 0, 0, 22, 11, Surce.PLBTN.hdc, N * 22, BoTN(N) * 11, vbSrcCopy
BTN(N).Refresh
Next
End Sub


Sub SaveList(fPath)
I = 1
Open fPath For Binary As #1
Mx = Surce.List1.ListCount
For N = 0 To Mx - 1
Surce.List1.ListIndex = N
T$ = Surce.List1.text & Chr$(0)
Lg = Len(T$)
For a = 1 To Lg
Put #1, I, Mid$(T$, a, 1)
I = I + 1
Next
Next
Close #1
End Sub



Sub LoadList(fPath)
Lg = Str(FileLen(fPath))
If Lg > 3 Then
Form2.List1.Clear
Surce.List1.Clear
txt$ = ""
S$ = " "
Open fPath For Binary As #1
For I = 1 To Lg
Get #1, I, S$
txt$ = txt$ & S$
Next
Close #1
Lg = Len(txt$)
SN = 1
For N = 1 To Lg
S$ = Mid$(txt$, N, 1)
If S$ = Chr$(0) Then
Form2.List1.AddItem FileName(Mid$(txt$, SN, N - SN)), 0
Surce.List1.AddItem Mid$(txt$, SN, N - SN), 0
SN = N + 1
End If
Next
End If
End Sub


