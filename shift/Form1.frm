VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "ShiftRG"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   1785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Очистить"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ввод"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "www.schemz.narood.ru"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i, k As Integer
Private BufferPort  As String
Private LED_on(8) As Boolean

Private Sub Command1_Click(Index As Integer)
LED_on(Index) = True
Command1(Index).BackColor = vbRed
End Sub

Private Sub Command2_Click()
Call ComPortOut
End Sub

Private Sub Command3_Click()
For i = 0 To 7
LED_on(i) = False
Command1(i).BackColor = &H8000000F
Next i
End Sub

Private Sub Form_Load()
For i = 0 To 7
LED_on(i) = False
Command1(i).BackColor = &H8000000F
Next i
        MSComm1.CommPort = 1
        MSComm1.Settings = "115200,N,8,1"
        MSComm1.PortOpen = True
        MSComm1.RTSEnable = True
        MSComm1.DTREnable = True
End Sub


Sub ComPortOut()    'подпрограмма вывода  на СОМ-ПОРТ
Dim NumDec, Bit(8) As Integer

    MSComm1.DTREnable = True 'срез регистра-защелки
  For i = 0 To 7        ' выдача на порт байта (первым идёт старший бит LED_on(7))
    If LED_on(7 - i) = True Then
        MSComm1.RTSEnable = True 'данные на регистр сдвига
    End If
    If LED_on(7 - i) = False Then
        MSComm1.RTSEnable = False   'данные на регистр сдвига
    End If

    MSComm1.Output = Chr(255) 'импульс сдвига
 Next i
  '------------------------------------------
           Do
        DoEvents                        '
         BufferPort = BufferPort & MSComm1.Input 'скачиваем из буфера порта в строковую переменную BufferPort данные до
       Loop Until MSComm1.InBufferCount = 0 'тех пор, пока в нём не останется ни одного байта (счётчик=0)
        BufferPort = ""
   '------------------------------------------
    MSComm1.DTREnable = False   'фронт- запись в защёлку регистра сдвига

End Sub

