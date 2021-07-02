VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   885
   ClientTop       =   1125
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Проект1.Graf3D Graf3D1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13361
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Написано Vir <ou953@ipn.ru>

Private Sub Form_Activate()
    Graf3D1.Width = Form1.ScaleWidth
    Graf3D1.Height = Form1.ScaleHeight
    
    Call Graf3D1.AddMass
    Graf3D1.Enabled = True
End Sub


