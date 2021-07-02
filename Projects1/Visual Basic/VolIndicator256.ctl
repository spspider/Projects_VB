VERSION 5.00
Begin VB.UserControl VolIndic 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   256
      X2              =   256
      Y1              =   16
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   256
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   256
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line indicator 
      BorderColor     =   &H00004000&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   16
   End
End
Attribute VB_Name = "VolIndic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Имя: Иникатор звука
'Вход: Для 256 бит
'Автор: Рашид Алиев
'Дата: 29 февраль 2004
'Почта: rashid_aliyev@yahoo.com
'Сайт: http://rashid4ever.narod.ru

'Значения для внутренних переменных по-умолчанию:
Const m_def_ILowColor = vbGreen
Const m_def_IMedColor = vbYellow
Const m_def_IHiColor = vbRed
Const m_def_LowMax = 128
Const m_def_MedMax = 192
Const m_def_Value = 0
'Объявление внутренних переменных:
Dim m_ILowColor As OLE_COLOR
Dim m_IMedColor As OLE_COLOR
Dim m_IHiColor As OLE_COLOR
Dim m_LowMax As Long
Dim m_MedMax As Long
Dim m_Value As Long
'Свойства
Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
    Draw
End Property
Public Property Get LowMax() As Long
    LowMax = m_LowMax
End Property
Public Property Let LowMax(ByVal New_LowMax As Long)
    m_LowMax = New_LowMax
    PropertyChanged "LowMax"
    Draw
End Property
Public Property Get MedMax() As Long
    MedMax = m_MedMax
End Property
Public Property Let MedMax(ByVal New_MedMax As Long)
    m_MedMax = New_MedMax
    PropertyChanged "MedMax"
    Draw
End Property
'Цвет фона контроля
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Draw
End Property
Public Property Get ILowColor() As OLE_COLOR
    ILowColor = m_ILowColor
End Property
Public Property Let ILowColor(ByVal New_ILowColor As OLE_COLOR)
    m_ILowColor = New_ILowColor
    PropertyChanged "ILowColor"
    Draw
End Property
Public Property Get IMedColor() As OLE_COLOR
    IMedColor = m_IMedColor
End Property
Public Property Let IMedColor(ByVal New_IMedColor As OLE_COLOR)
    m_IMedColor = New_IMedColor
    PropertyChanged "IMedColor"
    Draw
End Property
Public Property Get IHiColor() As OLE_COLOR
    IHiColor = m_IHiColor
End Property
Public Property Let IHiColor(ByVal New_IHiColor As OLE_COLOR)
    m_IHiColor = New_IHiColor
    PropertyChanged "IHiColor"
    Draw
End Property
Private Sub UserControl_Resize()
    Height = 255
    Width = 3860
    Draw
End Sub
Public Property Get IWidth() As Integer
    IWidth = UserControl.DrawWidth
End Property
Public Property Let IWidth(ByVal New_IWidth As Integer)
    If New_IWidth < 1 Then New_IWidth = 1
    If New_IWidth > 4 Then New_IWidth = 4
    UserControl.DrawWidth() = New_IWidth
    PropertyChanged "IWidth"
    Draw
End Property
'Инициализация свойств контрола
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_ILowColor = m_def_ILowColor
    m_IMedColor = m_def_IMedColor
    m_IHiColor = m_def_IHiColor
    m_LowMax = m_def_LowMax
    m_MedMax = m_def_MedMax
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ILowColor = PropBag.ReadProperty("ILowColor", m_def_ILowColor)
    m_IMedColor = PropBag.ReadProperty("IMedColor", m_def_IMedColor)
    m_IHiColor = PropBag.ReadProperty("IHiColor", m_def_IHiColor)
    UserControl.DrawWidth = PropBag.ReadProperty("IWidth", 2)
    m_LowMax = PropBag.ReadProperty("LowMax", m_def_LowMax)
    m_MedMax = PropBag.ReadProperty("MedMax", m_def_MedMax)
End Sub
Private Sub UserControl_Show()
'перерисовка графики при запуске контрола
    Draw
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ILowColor", m_ILowColor, m_def_ILowColor)
    Call PropBag.WriteProperty("IMedColor", m_IMedColor, m_def_IMedColor)
    Call PropBag.WriteProperty("IHiColor", m_IHiColor, m_def_IHiColor)
    Call PropBag.WriteProperty("IWidth", UserControl.DrawWidth, 2)
    Call PropBag.WriteProperty("LowMax", m_LowMax, m_def_LowMax)
    Call PropBag.WriteProperty("MedMax", m_MedMax, m_def_MedMax)
End Sub
Private Sub Draw()
    Dim Val
    Cls
    'Индикатор уровня
    If Value > 255 Then Value = 255
    If Value < 0 Then Value = 0
    
    Val = Value
    
    If Val > MedMax And Val <= 255 Then
        Line (2, 2)-(MedMax, 14), ILowColor, BF
        Line (LowMax + 1, 2)-(MedMax, 14), IMedColor, BF
        Line (MedMax + 1, 2)-(Val, 14), IHiColor, BF
    End If
    If Val > LowMax And Val <= MedMax Then
        Line (2, 2)-(LowMax, 14), ILowColor, BF
        Line (LowMax + 1, 2)-(Val, 14), IMedColor, BF
    End If
    If Val > 0 And Val <= 128 Then Line (2, 2)-(Val, 14), ILowColor, BF
End Sub
