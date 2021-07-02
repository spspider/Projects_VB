VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lucky 7"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   315
      Left            =   480
      TabIndex        =   16
      Top             =   180
      Width           =   735
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   795
      LargeChange     =   25
      Left            =   450
      Max             =   8000
      Min             =   1
      TabIndex        =   14
      Top             =   5580
      Value           =   500
      Width           =   195
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   795
      LargeChange     =   25
      Left            =   675
      Max             =   2500
      Min             =   1
      TabIndex        =   13
      Top             =   5580
      Value           =   1200
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   540
      TabIndex        =   12
      ToolTipText     =   "1 or 3 line play."
      Top             =   5040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "100"
      Top             =   540
      Width           =   645
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   1245
      Picture         =   "Form2.frx":267FA
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   9
      Top             =   4965
      Width           =   3315
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   180
         Left            =   2235
         Top             =   495
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   22140
      Left            =   8640
      Picture         =   "Form2.frx":2C2DC
      ScaleHeight     =   22080
      ScaleWidth      =   960
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   3570
      ScaleHeight     =   2880
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   1590
      Width           =   990
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   22140
      Left            =   7560
      Picture         =   "Form2.frx":4371E
      ScaleHeight     =   22080
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   2400
      ScaleHeight     =   2880
      ScaleWidth      =   960
      TabIndex        =   2
      Top             =   1590
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   22140
      Left            =   6480
      Picture         =   "Form2.frx":5AB60
      ScaleHeight     =   22080
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   1230
      ScaleHeight     =   2880
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   1590
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2500"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   420
      TabIndex        =   15
      ToolTipText     =   "Reels, sound and counter speed."
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2580
      TabIndex        =   10
      Top             =   885
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   8
      Top             =   3885
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   7
      Top             =   2910
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   6
      Top             =   1935
      Width           =   285
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Stoppen1 As Integer
Public Stoppen2 As Integer
Public Stoppen3 As Integer
Public S1 As Integer, S2 As Integer, S3 As Integer
Public OnOff1 As Boolean, OnOff2 As Boolean, OnOff3 As Boolean
Sub Counter(ByVal Amount As Integer, ByVal R As Integer)

    Label4.Caption = "0"
    
    Dim i As Long, y As Long
    Dim Tel As Long, ColorBuffer As Long
    Tel = VScroll1.Value
    Tel = Tel * 200
    ColorBuffer = Label2(R).ForeColor
    Label2(R).ForeColor = &HFFFF&   'Yellow
    Do Until i = Amount
        Text1.Text = Str(Val(Text1.Text) + 1)
        Label4.Caption = Trim(Str(Val(Label4.Caption) - 1))
        BeginPlaySound 60
        y = 0
         Do Until y = Tel
            y = y + 1
         Loop
        i = i + 1
    Loop
    Label2(R).ForeColor = ColorBuffer
    
End Sub

Function Wheel1(R As Integer) As String
    
    Select Case R
        Case 1, 3, 9, 13, 16, 19: Wheel1 = "Sinas"
        Case 0, 7, 11: Wheel1 = "Bel"
        Case 2, 8, 14, 17: Wheel1 = "Druif"
        Case 5, 10: Wheel1 = "Kers"
        Case 6, 12: Wheel1 = "1Bar"
        Case 18: Wheel1 = "2Bar"
        Case 15: Wheel1 = "3Bar"
        Case 4: Wheel1 = "7"
    End Select
    
End Function

Sub RollIt()
'The spinning of the wheels can be speed up by changing the copy steps,
'enlarging the steps increases the speed.

    Command1.Enabled = False
    Static Stap1 As Long, Step1 As Integer
    Dim Teller1 As Integer, Go1 As Boolean
    Static Stap2 As Long, Step2 As Integer
    Dim Teller2 As Integer, Go2 As Boolean
    Static Stap3 As Long, Step3 As Integer
    Dim Teller3 As Integer, Go3 As Boolean
    Dim i  As Long, y As Long, Speed As Long
    Go1 = True
    Go2 = True
    Go3 = True
    Speed = VScroll2.Value  'If you read the value of the ScrollBar during the loop it slows down
    Do Until i >= 1200
        If Go1 = True Then  'Rol 1 start ***************************************************************
            Call BitBlt(hDestDC:=Picture2.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                        hSrcDC:=Picture1.hDC, xSrc:=0, ySrc:=Stap1, _
                        dwRop:=SRCCOPY)
            Stap1 = Stap1 - 8
            Step1 = Step1 + 1
            Teller1 = Teller1 + 1 'it should at least make one turn before it stops.
            If Step1 = (8 * Stoppen1) + 1 And Teller1 > 160 Then    'At least 2 turns!
                Go1 = False
                S1 = Stap1 + 8
                BeginPlaySound 63    '<---- Sound of the wheel stopping.
                Speed = Speed + 7000 'try and keep the speed correct when the others stop
            End If
            If Stap1 <= 0 Then       'If the wheel has been round, Stap should be set back to the beginning,
                Stap1 = 1280         'in that way we know which symbol is in front..
                Step1 = 0            '> go's around 80x
            End If
        End If
        If Go2 = True Then  'Rol 2 start ***************************************************************
            Call BitBlt(hDestDC:=Picture3.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                        hSrcDC:=Picture4.hDC, xSrc:=0, ySrc:=Stap2, _
                        dwRop:=SRCCOPY)
            Stap2 = Stap2 - 8
            Step2 = Step2 + 1
            Teller2 = Teller2 + 1 'it should at least make one turn before it stops.
            If Step2 = (8 * Stoppen2) + 1 And Teller2 > 320 Then    'Only if wheel 1 has stopped, wheel 2 may stop
                Go2 = False                                         'and so on, see wheel 3 also
                S2 = Stap2 + 8
                BeginPlaySound 63   '<---- Sound of the wheel stopping.
                Speed = Speed + 7000 'try to keep speed going correctly!
            End If
            If Stap2 <= 0 Then
                Stap2 = 1280
                Step2 = 0
            End If
        End If
        If Go3 = True Then  'Rol 3 start ***************************************************************
            Call BitBlt(hDestDC:=Picture5.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                        hSrcDC:=Picture6.hDC, xSrc:=0, ySrc:=Stap3, _
                        dwRop:=SRCCOPY)
            Stap3 = Stap3 - 8
            Step3 = Step3 + 1
            Teller3 = Teller3 + 1 'it should at least make one turn before it stops.
            If Step3 = (8 * Stoppen3) + 1 And Teller3 > 480 Then    'Stop only if wheel 2 has stopped
                BeginPlaySound 63   '<---- Sound of the wheel stopping.
                Go3 = False
                S3 = Stap3 + 8
            End If
            If Stap3 <= 0 Then
                Stap3 = 1280
                Step3 = 0
            End If
        End If
        If Go1 = False And Go2 = False And Go3 = False Then
            If Check1.Value = 1 Then
                If Val(Text1.Text) < 4 Then
                    Win1Rol
                Else
                    Win     'Lets see if we won something
                End If
            Else
                Win1Rol
            End If
            Exit Do 'and exit the loop!
        End If
            y = 0                   'This counter is controlled by the scrollbar(see comment above)
            Do Until y >= Speed
              y = y + 1
            Loop
    i = i + 1
    Loop

End Sub

Function Wheel2(R As Integer) As String

    Select Case R
        Case 2, 11, 19: Wheel2 = "Sinas"
        Case 5, 12, 18: Wheel2 = "Bel"
        Case 0, 4, 15: Wheel2 = "Druif"
        Case 1, 3, 6, 8, 13, 16: Wheel2 = "Kers"
        Case 10, 17: Wheel2 = "1Bar"
        Case 7: Wheel2 = "2Bar"
        Case 14: Wheel2 = "3Bar"
        Case 9: Wheel2 = "7"
    End Select
    
End Function

Function Wheel3(R As Integer) As String

    Select Case R
        Case 0, 8, 10, 14: Wheel3 = "Sinas"
        Case 3, 11, 19: Wheel3 = "Bel"
        Case 6, 15, 18: Wheel3 = "Druif"
        Case 1, 2, 7, 9, 12, 16: Wheel3 = "Kers"
        Case 5: Wheel3 = "1Bar"
        Case 17: Wheel3 = "2Bar"
        Case 13: Wheel3 = "3Bar"
        Case 4: Wheel3 = "7"
    End Select
    
End Function


Sub Win()
    
    Dim Winning(0 To 2) As Integer
    Dim B1 As Integer, B2 As Integer, B3 As Integer
    Dim i As Long
    Dim RolA(0 To 2) As String
    Dim RolB(0 To 2) As String
    Dim RolC(0 To 2) As String
    Dim Bar1 As Boolean, Bar2 As Boolean, Bar3 As Boolean
    
    RolA(1) = Wheel1(Stoppen1)
    B1 = Stoppen1 + 1
    If B1 = 20 Then B1 = 0
    RolA(0) = Wheel1(B1)
    B1 = Stoppen1 - 1
    If B1 = -1 Then B1 = 19
    RolA(2) = Wheel1(B1)
    
    RolB(1) = Wheel2(Stoppen2)
    B2 = Stoppen2 + 1
    If B2 = 20 Then B2 = 0
    RolB(0) = Wheel2(B2)
    B2 = Stoppen2 - 1
    If B2 = -1 Then B2 = 19
    RolB(2) = Wheel2(B2)
    
    RolC(1) = Wheel3(Stoppen3)
    B3 = Stoppen3 + 1
    If B3 = 20 Then B3 = 0
    RolC(0) = Wheel3(B3)
    B3 = Stoppen3 - 1
    If B3 = -1 Then B3 = 19
    RolC(2) = Wheel3(B3)

    i = 0
    Do Until i = 3
        If RolA(i) = "Kers" And RolB(i) <> "Kers" Then Winning(i) = 2
        If RolA(i) = "Kers" And RolB(i) = "Kers" And RolC(i) <> "Kers" Then Winning(i) = 5
        If RolA(i) = "Kers" And RolB(i) = "Kers" And RolC(i) = "Kers" Then Winning(i) = 10
        If RolA(i) = "Sinas" And RolB(i) = "Sinas" And RolC(i) = "Sinas" Then Winning(i) = 10
        If RolA(i) = "Sinas" And RolB(i) = "Sinas" And RolC(i) = "3Bar" Then Winning(i) = 10
        If RolA(i) = "Druif" And RolB(i) = "Druif" And RolC(i) = "Druif" Then Winning(i) = 14
        If RolA(i) = "Druif" And RolB(i) = "Druif" And RolC(i) = "3Bar" Then Winning(i) = 14
        If RolA(i) = "Bel" And RolB(i) = "Bel" And RolC(i) = "Bel" Then Winning(i) = 18
        If RolA(i) = "Bel" And RolB(i) = "Bel" And RolC(i) = "3Bar" Then Winning(i) = 18

        If RolA(i) = "1Bar" And RolB(i) = "1Bar" And RolC(i) = "1Bar" Then
            Winning(i) = 50
            Bar1 = True
        Else
            Bar1 = False
        End If
        If RolA(i) = "2Bar" And RolB(i) = "2Bar" And RolC(i) = "2Bar" Then
            Winning(i) = 100
            Bar2 = True
        Else
            Bar2 = False
        End If
        If RolA(i) = "3Bar" And RolB(i) = "3Bar" And RolC(i) = "3Bar" Then
            Winning(i) = 150
            Bar3 = True
        Else
            Bar3 = False
        End If
        If Bar1 = False And Bar2 = False And Bar3 = False Then
            If RolA(i) = "1Bar" Or RolA(i) = "2Bar" Or RolA(i) = "3Bar" Then
               If RolB(i) = "1Bar" Or RolB(i) = "2Bar" Or RolB(i) = "3Bar" Then
                   If RolC(i) = "1Bar" Or RolC(i) = "2Bar" Or RolC(i) = "3Bar" Then
                        Winning(i) = 25
                   End If
               End If
            End If
        End If
        
        If RolA(i) = "7" And RolB(i) = "7" And RolC(i) <> "7" Then Winning(i) = 250
        If RolA(i) = "7" And RolB(i) = "7" And RolC(i) = "7" Then Winning(i) = 500
    i = i + 1
    Loop
    Dim j As Long
    Dim Tel As Long
    Tel = VScroll1.Value
    Tel = Tel * 300
    i = 0
    Do Until i = 3
        If Winning(i) > 0 Then
            Counter Winning(i), i
              j = 0                 'This counter should be controlable (see Command button)
              Do Until j = Tel      'Its here because there should be some time between the win wheels,
                j = j + 1           'in this way it doesn't go right on when there are more win lines.
              Loop
        End If
    i = i + 1
    Loop
    
End Sub
Sub Win1Rol()
    
    Dim Winning As Integer
    Dim B1 As Integer, B2 As Integer, B3 As Integer
    Dim i As Long
    Dim RolA As String
    Dim RolB As String
    Dim RolC As String
    Dim Bar1 As Boolean, Bar2 As Boolean, Bar3 As Boolean
    
    RolA = Wheel1(Stoppen1)
    
    RolB = Wheel2(Stoppen2)
    
    RolC = Wheel3(Stoppen3)

    If RolA = "Kers" And RolB <> "Kers" Then Winning = 2
    If RolA = "Kers" And RolB = "Kers" And RolC <> "Kers" Then Winning = 5
    If RolA = "Kers" And RolB = "Kers" And RolC = "Kers" Then Winning = 10
    If RolA = "Sinas" And RolB = "Sinas" And RolC = "Sinas" Then Winning = 10
    If RolA = "Sinas" And RolB = "Sinas" And RolC = "3Bar" Then Winning = 10
    If RolA = "Druif" And RolB = "Druif" And RolC = "Druif" Then Winning = 14
    If RolA = "Druif" And RolB = "Druif" And RolC = "3Bar" Then Winning = 14
    If RolA = "Bel" And RolB = "Bel" And RolC = "Bel" Then Winning = 18
    If RolA = "Bel" And RolB = "Bel" And RolC = "3Bar" Then Winning = 18

    If RolA = "1Bar" And RolB = "1Bar" And RolC = "1Bar" Then
        Winning = 50
        Bar1 = True
    Else
        Bar1 = False
    End If
    If RolA = "2Bar" And RolB = "2Bar" And RolC = "2Bar" Then
        Winning = 100
        Bar2 = True
    Else
        Bar2 = False
    End If
    If RolA = "3Bar" And RolB = "3Bar" And RolC = "3Bar" Then
        Winning = 150
        Bar3 = True
    Else
        Bar3 = False
    End If
    If Bar1 = False And Bar2 = False And Bar3 = False Then
        If RolA = "1Bar" Or RolA = "2Bar" Or RolA = "3Bar" Then
           If RolB = "1Bar" Or RolB = "2Bar" Or RolB = "3Bar" Then
               If RolC = "1Bar" Or RolC = "2Bar" Or RolC = "3Bar" Then
                    Winning = 25
               End If
           End If
        End If
    End If
    
    If RolA = "7" And RolB = "7" And RolC <> "7" Then Winning = 250
    If RolA = "7" And RolB = "7" And RolC = "7" Then Winning = 500

    Dim j As Long
    Dim Tel As Long
    Tel = VScroll1.Value
    Tel = Tel * 300
    If Winning > 0 Then
        Counter Winning, 1
    End If
    
End Sub


Private Sub Check1_Click()
    
    If Check1.Value = 1 Then
        Label2(0).ForeColor = &HFF&     'red
        Label2(1).ForeColor = &HFF&     'red
        Label2(2).ForeColor = &HFF&     'red
    Else
        Label2(0).ForeColor = &H0&      'darkgray
        Label2(1).ForeColor = &HFF&     'red
        Label2(2).ForeColor = &H0&      'darkgray
    End If

    'darkgray = &H0&
    'red = &H000000FF&
    
End Sub

Private Sub Command1_Click()
                                           
    Randomize CDbl(Now) + Timer
    
    Stoppen1 = Int((20 * Rnd) + 0)
    Stoppen2 = Int((20 * Rnd) + 0)
    Stoppen3 = Int((20 * Rnd) + 0)
'    Stoppen1 = 6       'just to test
'    Stoppen2 = 7
'    Stoppen3 = 5
    Label4.Caption = ""
    Dim i As Long, y As Long
    Dim Tel As Double
    Tel = VScroll1.Value
    Tel = Tel * 100
    
    Dim Counters As Integer
    If Check1.Value = 1 Then
        If Val(Text1.Text) < 4 Then
            Check1.Value = 0
            Check1_Click
            Counters = 1
        Else
            Counters = 4
        End If
    Else
        Counters = 1
    End If
    
    Do Until i = Counters                       'extract the points
         Text1.Text = Str(Val(Text1.Text) - 1)
         BeginPlaySound 62                      'every point a tick
        y = 0
         Do Until y = Tel                       'should be controllable
            y = y + 1
         Loop
        i = i + 1
    Loop
    
    Picture2.AutoRedraw = False
    Picture3.AutoRedraw = False
    Picture5.AutoRedraw = False
    
    BeginPlaySound 61                           'The start sound of the wheels
    RollIt                                      'starting the wheels, we already know where they will stop
    Command1.Enabled = True
                                                'enable the button when rollit finishes
    Picture2.AutoRedraw = True
    Picture3.AutoRedraw = True
    Picture5.AutoRedraw = True
    
End Sub





Private Sub Form_Load()

    With Form2
        .Show
        .Refresh
        '.Width = 5235
    End With
    Shape1.BorderStyle = 0
    'to see something when the program starts we copy the first positions to the picture boxes
    Call BitBlt(hDestDC:=Picture2.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture1.hDC, xSrc:=0, ySrc:=0, _
                dwRop:=SRCCOPY)
    Call BitBlt(hDestDC:=Picture3.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture4.hDC, xSrc:=0, ySrc:=0, _
                dwRop:=SRCCOPY)
    Call BitBlt(hDestDC:=Picture5.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture6.hDC, xSrc:=0, ySrc:=0, _
                dwRop:=SRCCOPY)
                
    Label2(0).ForeColor = &HFF&     'red   'We start with the option "3 Line play",
    Label2(1).ForeColor = &HFF&     'red   'so all arrows are "active"
    Label2(2).ForeColor = &HFF&     'red   '
                
End Sub

Private Sub Form_Paint()
'Well... this doesn't work as it should


    Call BitBlt(hDestDC:=Picture2.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture1.hDC, xSrc:=0, ySrc:=S1, _
                dwRop:=SRCCOPY)
    Call BitBlt(hDestDC:=Picture3.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture4.hDC, xSrc:=0, ySrc:=S2, _
                dwRop:=SRCCOPY)
    Call BitBlt(hDestDC:=Picture5.hDC, x:=0, y:=0, nWidth:=64, nHeight:=192, _
                hSrcDC:=Picture6.hDC, xSrc:=0, ySrc:=S3, _
                dwRop:=SRCCOPY)
    
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Is this a cheat? well you could look at it that way but to get more credit
'by stopping and starting the program would be kind of silly... so I have hidden
'it a bit..

    Static OnOff As Boolean
    If x > Shape1.Left And x < Shape1.Left + Shape1.Width And _
       y > Shape1.Top And y < Shape1.Top + Shape1.Height Then
            OnOff = Not (OnOff)
            Select Case OnOff
                Case True
                    Shape1.BorderStyle = 1                              'Lets show something is happening.
                        Dim Tel As Double, i As Long, iy As Long        '
                        Text1.Text = ""                                 '
                        Tel = VScroll1.Value                            'Correct the counter for the loop
                        Tel = Tel * 100                                 '
                        Do Until i = 50                                 '
                             Text1.Text = Str(Val(Text1.Text) + 10)     'Enlarge the value of the textbox by 10
                             BeginPlaySound 62                          'and play a sound
                             iy = 0                                     '
                             Do Until iy = Tel                          '
                                iy = iy + 1                             '
                             Loop                                       '
                            i = i + 1                                   '
                        Loop                                            '
                    Shape1.BorderStyle = 0                              'Put everything back as it was
                    OnOff = Not (OnOff)                                 'by putting it back here it will never
                Case False                                              'reach the other Case...
                    Shape1.BorderStyle = 0                              'oh well...
            End Select
        End If
    

End Sub


Private Sub VScroll1_Change()
    
    Label1.Caption = Str(VScroll1.Value)
    
End Sub

Private Sub VScroll2_Change()

    Label1.Caption = Str(VScroll2.Value)
    
End Sub


