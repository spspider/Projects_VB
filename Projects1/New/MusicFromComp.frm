VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Beep& Lib "kernel32" _
        (ByVal dwFreq As Long, ByVal dwDuration As Long)
Dim freq As Long
Dim A As Long
Dim B As Long
Private Sub Form_Load()
freq = 37
End Sub

Private Sub Timer1_Timer()
A = A + 1

Static freq&
Static freq2&

freq = Abs(Tan(A / 16) * 100) ''1
Call Beep(freq, 20)

If A = 5 Then ''2
freq2 = 3333
Call Beep(freq2, 20)
End If
If A = 20 Then
freq2 = 3333
Call Beep(freq2, 20)
End If

If A = 40 Then
freq2 = 4444
Call Beep(freq2, 50)
End If
If A = 45 Then
freq2 = 2222
Call Beep(freq2, 30)
End If

If A = 50 Then
freq2 = 200
Call Beep(freq2, 100)
End If
If A = 55 Then
freq2 = 300
Call Beep(freq2, 100)
End If
If A = 60 Then
freq2 = 100
'Call Beep(freq2, 500)
End If
If A = 61 Then ''2
freq2 = 500
Call Beep(freq2, 70)
End If
If A = 62 Then
freq2 = 100
Call Beep(freq2, 70)
End If






Caption = freq

If A > 30 Then A = 0
End Sub
