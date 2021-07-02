VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form index 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Автор исходника: Яценко Александр(Zarak Soft)."
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer vrag 
      Interval        =   10
      Left            =   5400
      Top             =   6360
   End
   Begin VB.Timer other 
      Interval        =   10
      Left            =   480
      Top             =   960
   End
   Begin VB.Timer liftvert 
      Interval        =   10
      Left            =   4320
      Top             =   3360
   End
   Begin VB.Timer lifti 
      Interval        =   1
      Left            =   5280
      Top             =   3480
   End
   Begin VB.Timer dn 
      Interval        =   1
      Left            =   840
      Top             =   5160
   End
   Begin VB.Timer movedown 
      Interval        =   1
      Left            =   840
      Top             =   4560
   End
   Begin VB.Timer moveup 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   3840
   End
   Begin VB.Timer gr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   4320
   End
   Begin VB.Timer gl 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   4320
   End
   Begin MSComctlLib.ImageList I 
      Left            =   7320
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "dl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":031A
            Key             =   "monstr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09F2
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D0C
            Key             =   "zubi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":121D
            Key             =   "finish"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":172C
            Key             =   "dr"
         EndProperty
      EndProperty
   End
   Begin VB.Image enemy 
      Height          =   465
      Left            =   4200
      Top             =   6240
      Width           =   450
   End
   Begin VB.Image clock 
      Height          =   480
      Index           =   4
      Left            =   3840
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image clock 
      Height          =   480
      Index           =   3
      Left            =   8520
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image clock 
      Height          =   480
      Index           =   2
      Left            =   4920
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image clock 
      Height          =   480
      Index           =   1
      Left            =   3240
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image clock 
      Height          =   480
      Index           =   0
      Left            =   11280
      Top             =   2040
      Width           =   480
   End
   Begin VB.Shape stena 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   9135
      Index           =   3
      Left            =   12000
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape stena 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   9015
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image finish 
      Height          =   480
      Left            =   11520
      Picture         =   "Form1.frx":1A46
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image zubi 
      Height          =   330
      Left            =   9000
      Picture         =   "Form1.frx":1F45
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   1725
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   8
      Left            =   10800
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Image dog 
      Height          =   480
      Left            =   1680
      Picture         =   "Form1.frx":2446
      Top             =   8400
      Width           =   480
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   7
      Left            =   2520
      Top             =   6720
      Width           =   4455
   End
   Begin VB.Shape stena 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   1
      Left            =   6240
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   6
      Left            =   8040
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Shape stena 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Index           =   0
      Left            =   8640
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   5
      Left            =   10200
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   4
      Left            =   2040
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Shape plat 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   3
      Left            =   480
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   2
      Left            =   7560
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Shape plat 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   1
      Left            =   6120
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape plat 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   8880
      Width           =   8775
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------Автор исходника: Яценко Александр----------------------
Dim up As Double
Dim jumpup As Boolean
Dim jumpdown As Boolean
Dim go As Boolean
Dim levo As Boolean
Dim vni As Boolean
Dim ogran As Boolean
Dim gor As Boolean
Dim ver As Boolean
Dim sten As Boolean
Dim stenr As Boolean
Dim chas As Long
Dim vraggo As Boolean

Private Sub dn_Timer()
For b = 0 To plat.Count - 1
If plat(b).Top = dog.Top + dog.Height - 20 Then
If dog.Left + dog.Width < plat(b).Left Or dog.Left > plat(b).Left + plat(b).Width Then
movedown.Enabled = True
End If
End If
Next b
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
If jumpup = True Then
jumpup = False
up = 60
moveup.Enabled = True
End If
End If
If KeyCode = 37 Then
dog.Picture = I.ListImages("dl").Picture
gl.Enabled = True
End If
If KeyCode = 39 Then
dog.Picture = I.ListImages("dr").Picture
gr.Enabled = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
gl.Enabled = False
End If
If KeyCode = 39 Then
gr.Enabled = False
End If
End Sub

Private Sub Form_Load()
enemy.Picture = I.ListImages("monstr").Picture
finish.Picture = I.ListImages("finish").Picture
For m = 0 To clock.Count - 1
clock(m).Picture = I.ListImages("clock").Picture
Next m
vraggo = True
chas = clock.Count
ver = True
gor = True
antierror
jumpup = True
ogran = True
End Sub

Private Sub gl_Timer()
stenr = False
For k = 0 To stena.Count - 1
If dog.Left < stena(k).Left + stena(k).Width + 10 And dog.Left > stena(k).Left + stena(k).Width And dog.Top + dog.Height > stena(k).Top And stena(k).Top + stena(k).Height > dog.Top Then
stenr = True
Else
End If
Next k
If stenr = False Then
dog.Left = dog.Left - 20
End If
End Sub
Private Sub gr_Timer()
sten = True
For t = 0 To stena.Count - 1
If dog.Left + dog.Height > stena(t).Left And dog.Left < stena(t).Left + stena(t).Width And dog.Top > stena(t).Top - 100 And dog.Top < stena(t).Top + stena(t).Height Then
sten = False
Else
End If
Next t
If sten = True Then
dog.Left = dog.Left + 20
End If
End Sub

Private Sub Image3_Click(index As Integer)

End Sub

Private Sub lifti_Timer()
'----боковой движок------
If gor = True Then
If plat(1).Left + plat(1).Width < plat(5).Left Then
plat(1).Left = plat(1).Left + 10
Else
gor = False
End If
End If
If gor = False Then
If plat(1).Left > plat(4).Left + plat(4).Width Then
plat(1).Left = plat(1).Left - 10
Else
gor = True
End If
End If
If dog.Top + dog.Height - 20 = plat(1).Top And dog.Left + dog.Width > plat(1).Left And dog.Left < plat(1).Left + plat(1).Width Then
If gl.Enabled = False And gr.Enabled = False Then
If gor = True Then
dog.Left = dog.Left + 10
Else
dog.Left = dog.Left - 10
End If
End If
End If
End Sub

Private Sub liftvert_Timer()
'----верховой движок------
If ver = True Then
If plat(3).Top > 2520 Then
If dog.Top = plat(3).Top - dog.Height + 20 Then
dog.Top = dog.Top - 10
End If
plat(3).Top = plat(3).Top - 10
Else
ver = False
End If
End If
If ver = False Then
If plat(3).Top < 8520 Then
If dog.Top = plat(3).Top - dog.Height Then
dog.Top = dog.Top + 10
End If
plat(3).Top = plat(3).Top + 10
Else
ver = True
End If
End If
End Sub

Private Sub movedown_Timer()
For md = 0 To plat.Count - 1
If dog.Top + dog.Height > plat(md).Top Then
If dog.Top + dog.Height > plat(md).Top + plat(md).Height Then
Else
If dog.Left + dog.Width > plat(md).Left Then
If dog.Left < plat(md).Left + plat(md).Width Then
dog.Top = plat(md).Top - dog.Height
movedown.Enabled = False
jumpup = True
ogran = True
End If
End If
End If
End If
Next md
dog.Top = dog.Top + 20
End Sub

Private Sub moveup_Timer()
For g = 0 To plat.Count - 1
If dog.Left + dog.Width > plat(g).Left And dog.Left < plat(g).Left + plat(g).Width Then
If dog.Top > plat(g).Top And dog.Top < plat(g).Top + plat(g).Height Then
moveup.Enabled = False
movedown.Enabled = True
End If
End If
Next g
If up > 0 Then
up = up - 1
Else
moveup.Enabled = False
movedown.Enabled = True
End If
dog.Top = dog.Top - up
End Sub

Function antierror()
For t = 0 To plat.Count - 1
If t < 10 Then
plat(t).Top = plat(t).Top - t
Else
plat(t).Top = plat(t).Top + t
End If
Next t
End Function

Private Sub other_Timer()
index.Caption = "Осталось собрать " & chas & " будильников. Убейте врага прыжком сверху (как в марио)."
If dog.Top + dog.Height > zubi.Top And dog.Top < zubi.Top + zubi.Height And dog.Left + dog.Width > zubi.Left And dog.Left < zubi.Left + zubi.Width Then
MsgBox "Ты напоролся на шипы и проиграл!!!", 16, "Ты проиграл"
End
End If
For g = 0 To clock.Count - 1
If dog.Top + dog.Height > clock(g).Top And dog.Top < clock(g).Top + clock(g).Height And dog.Left + dog.Width > clock(g).Left And dog.Left < clock(g).Left + clock(g).Width Then
clock(g).Visible = False
clock(g).Top = 20000
chas = chas - 1
End If
Next g
If dog.Top + dog.Height > enemy.Top And dog.Top < enemy.Top + enemy.Height And dog.Left + dog.Width + 50 > enemy.Left And dog.Left - 50 < enemy.Left + enemy.Width Then
MsgBox "Враг тебя поразил и ты проиграл", 16, "Ты проиграл"
End
End If
If dog.Left + dog.Height > enemy.Left And dog.Left < enemy.Left + enemy.Width And dog.Top + dog.Height + 50 > enemy.Top And dog.Top < enemy.Top + enemy.Height Then
enemy.Top = 20000
enemy.Visible = False
End If
If dog.Left + dog.Width > finish.Left And dog.Left < finish.Left + finish.Width And dog.Top + dog.Height > finish.Top And dog.Top < finish.Top + finish.Height Then
If enemy.Visible = False And chas = 0 Then
MsgBox "Ты выйграл", 64, "УРА"
End
End If
End If
End Sub

Private Sub vrag_Timer()
If vraggo = True Then
If enemy.Left + enemy.Width < 6970 Then
enemy.Left = enemy.Left + 10
Else
vraggo = False
End If
End If
If vraggo = False Then
If enemy.Left > 2520 Then
enemy.Left = enemy.Left - 10
Else
vraggo = True
End If
End If
End Sub
