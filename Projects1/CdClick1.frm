VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CD"
   ClientHeight    =   0
   ClientLeft      =   -4665
   ClientTop       =   -1500
   ClientWidth     =   1725
   Icon            =   "CdClick1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1725
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim q As String
Private Sub Form_Load()
On Error GoTo er:
q = GetSetting("SP\SPSoft\CD\open_close", "Naw", "Naw")

If q = "False" Then
Status = mciSendString("Set CDAudio Door Open Wait", 0&, 0, 0)
q = "True"
Unload Me
Else
Status = mciSendString("Set CDAudio Door Closed Wait", 0&, 0, 0)
q = "False"
Unload Me
End If

er:
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "SP\SPSoft\CD\open_close", "Naw", "Naw", q
End Sub
