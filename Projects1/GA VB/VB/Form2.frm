VERSION 5.00
Object = "{A8F8E827-06DA-11D2-8D70-00A0C98B28E2}#1.0#0"; "VCMAXB.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin AxBrowse.AxBrowser AxBrowser1 
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закупка"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.B1.Band1 = 0
Form1.B1.Band2 = 0
Form1.B1.Band3 = 0
Form1.B1.Band4 = 0
Form1.B1.Band5 = 0
Form1.B1.Band6 = 0
Form1.B1.Band7 = 0
Form1.B1.Band8 = 0
Form1.B1.Band9 = 0
Form1.B1.Band10 = 0
Form1.B1.Band11 = 0
Form1.B1.Band12 = 0
End Sub

