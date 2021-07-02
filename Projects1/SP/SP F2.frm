VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SP"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "SP F2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Фото"
      Height          =   3615
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.Image Image1 
         Height          =   3270
         Left            =   120
         Picture         =   "SP F2.frx":34CA
         Top             =   240
         Width           =   3000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Информация"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.Label Label1 
         Caption         =   $"SP F2.frx":233FE
         Height          =   2895
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
