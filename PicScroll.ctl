VERSION 5.00
Begin VB.UserControl PicScroll 
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   ScaleHeight     =   8025
   ScaleWidth      =   9255
   Begin VB.PictureBox picBack1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   7575
      TabIndex        =   2
      Top             =   2880
      Width           =   7575
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   720
      ScaleHeight     =   1815
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "PicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Declarations
Dim ix As Long
Dim bDrag As Boolean
Dim iMin As Long
Dim iMax As Long
Dim iValue As Long

' Events
Event Change()
Event ErroValue()
Event LeftClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Property Get BackColor() As OLE_COLOR
    BackColor = picBack.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    picBack.BackColor = New_Color
    Call DrawBar
    
    PropertyChanged "BackColor"
End Property

Public Property Get Bar() As Picture
    Set Bar = picBar.Picture
End Property

Public Property Set Bar(ByVal New_Bar As Picture)
    Set picBar.Picture = New_Bar
    UserControl_Resize
    
    Call DrawBar
    PropertyChanged "Bar"
End Property


Private Sub CalcValue()
    iValue = ix / (ToPixels(picBack.Width) - ToPixels(picBar.Width)) * (iMax - iMin) + iMin
End Sub

Private Sub Calcx()
    ix = (iValue - iMin) / (iMax - iMin) * (ToPixels(picBack.Width) - ToPixels(picBar.Width))
End Sub

Private Sub DrawBar(Optional Calculatex As Boolean = True)
    If Calculatex Then Call Calcx
    
    picBack.Cls
    Call BitBlt(picBack.hdc, 0, 0, ix, picBack1.ScaleHeight, picBack1.hdc, 0, 0, vbSrcCopy)
    Dim iYa As Integer
    iYa = ((picBack.Height - picBar.Height) / 2) / Screen.TwipsPerPixelY
    Call BitBlt(picBack.hdc, ix, iYa, picBar.ScaleWidth, picBar.ScaleHeight, picBar.hdc, 0, 0, vbSrcCopy)
    
    picBack.Refresh
    
    UserControl.Refresh
End Sub
Public Property Get Max() As Long
    Max = iMax
End Property

Public Property Let Max(New_Max As Long)
    If New_Max < iValue Then
        'MsgBox "Maximum exceeds value!", vbOKOnly + vbExclamation, "Error"
        Exit Property
    End If
    
    iMax = New_Max
    Call DrawBar
    
    PropertyChanged "Max"
End Property

Public Property Get Min() As Long
    Min = iMin
End Property

Public Property Let Min(New_Min As Long)
    If iMin > iValue Then
        MsgBox "Minimum exceeds value!"
        Exit Property
    End If
    
    iMin = New_Min
    Call DrawBar
    
    PropertyChanged "Min"
End Property
Public Property Get Picture1() As Picture
    Set Picture1 = picBack.Picture
End Property

Public Property Set Picture1(ByVal New_Picture1 As Picture)
    Set picBack.Picture = New_Picture1
    UserControl_Resize
    Call DrawBar
    PropertyChanged "Picture1"
End Property
Public Property Get Picture2() As Picture
    Set Picture2 = picBack1.Picture
End Property

Public Property Set Picture2(ByVal New_Picture2 As Picture)
    Set picBack1.Picture = New_Picture2
    UserControl_Resize
    Call DrawBar
    PropertyChanged "Picture2"
End Property
Private Function ToPixels(ByVal nTwips As Long) As Long
    ToPixels = nTwips / Screen.TwipsPerPixelX
End Function

Public Property Get Value() As Long
    Value = iValue
End Property

Public Property Let Value(New_Value As Long)
    If New_Value < iMin Or New_Value > iMax Then
        'MsgBox "Value exceeds limits!", vbOKOnly + vbExclamation, "Error"
        RaiseEvent ErroValue
        Exit Property
    End If
    
    iValue = New_Value
    Call DrawBar
    
    PropertyChanged "Value"
End Property
Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    RaiseEvent LeftClick
Else
    If ToPixels(x) >= ix And ToPixels(x) <= ix + ToPixels(picBar.Width) And Button = 1 Then
        bDrag = True
    Else
        'MsgBox "Picture Scroller" & vbCrLf & vbCrLf & "Created by PowerSoft Programming" & vbCrLf & "Email: markvr@dsv.nl", vbOKOnly + vbInformation, "Picture Scroller"
        bDrag = True
        ix = ToPixels(x)
        
        If ix > ToPixels(picBack.Width) - (ToPixels(picBar.Width) / 2) Then ix = ToPixels(picBack.Width) - (ToPixels(picBar.Width) / 2)
        If ix < ToPixels(picBar.Width) / 2 Then ix = ToPixels(picBar.Width) / 2
        
        ix = ix - ToPixels(picBar.Width) / 2
        
        Call DrawBar(False)
        Call CalcValue
        Value = iValue
        RaiseEvent Change
        
    End If
    
    RaiseEvent MouseDown(Button, Shift, x, Y)
End If
End Sub


Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bDrag Then
        ix = ToPixels(x)
        
        If ix > ToPixels(picBack.Width) - (ToPixels(picBar.Width) / 2) Then ix = ToPixels(picBack.Width) - (ToPixels(picBar.Width) / 2)
        If ix < ToPixels(picBar.Width) / 2 Then ix = ToPixels(picBar.Width) / 2
        
        ix = ix - ToPixels(picBar.Width) / 2
        
        Call DrawBar(False)
        Call CalcValue
        Value = iValue
        RaiseEvent Change
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    bDrag = False
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub


Private Sub UserControl_Initialize()
    If iMax = 0 Then iMax = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picBack.Picture = PropBag.ReadProperty("Picture1", Nothing)
    picBack1.Picture = PropBag.ReadProperty("Picture2", Nothing)
    picBar.Picture = PropBag.ReadProperty("Bar", Nothing)
    picBack.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    iMin = PropBag.ReadProperty("Min", 0)
    iMax = PropBag.ReadProperty("Max", 100)
    iValue = PropBag.ReadProperty("Value", 0)
    
    Call DrawBar
End Sub

Private Sub UserControl_Resize()
    picBar.Top = 0
    picBar.Left = UserControl.Width + picBar.Width
    picBack1.Top = 0
    picBack1.Left = UserControl.Width + picBack1.Width
    'UserControl.Width = picBack.Width
    'UserControl.Height = picBack.Height
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture1", picBack.Picture, Nothing)
    Call PropBag.WriteProperty("Picture2", picBack1.Picture, Nothing)
    Call PropBag.WriteProperty("Bar", picBar.Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", picBack.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Min", iMin, 0)
    Call PropBag.WriteProperty("Max", iMax, 100)
    Call PropBag.WriteProperty("Value", iValue, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBar,picBack,-1,PaintPicture
Public Sub PaintPicture1(ByVal Picture As Picture, ByVal x1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal x2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    picBar.PaintPicture Picture, x1, Y1, Width1, Height1, x2, Y2, Width2, Height2, Opcode
    picBar.Picture = picBar.Image
    picBar.Height = Height2
    picBar.Width = Width2
    UserControl_Resize
    Call DrawBar
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,PaintPicture
Public Sub PaintPicture2(ByVal Picture As Picture, ByVal x1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal x2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    picBack.PaintPicture Picture, x1, Y1, Width1, Height1, x2, Y2, Width2, Height2, Opcode
    picBack.Picture = picBack.Image
    picBack.Height = Height2
    picBack.Width = Width2
    UserControl.Width = picBack.Width
    UserControl.Height = picBack.Height
    UserControl_Resize
    Call DrawBar
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack1,picBack,-1,PaintPicture
Public Sub PaintPicture3(ByVal Picture As Picture, ByVal x1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal x2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    picBack1.PaintPicture Picture, x1, Y1, Width1, Height1, x2, Y2, Width2, Height2, Opcode
    picBack1.Picture = picBack1.Image
    picBack1.Height = Height2
    picBack1.Width = Width2
    UserControl_Resize
    Call DrawBar
End Sub
