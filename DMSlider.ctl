VERSION 5.00
Begin VB.UserControl DMSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ToolboxBitmap   =   "DMSlider.ctx":0000
   Begin VB.PictureBox HandlePicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Picture         =   "DMSlider.ctx":0312
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox BackPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3600
      Picture         =   "DMSlider.ctx":0600
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox HandleMaskPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3600
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label MaskBool 
      Caption         =   "false"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Strech 
      Caption         =   "true"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label ValueLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label SliderCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slider"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Max 
      Caption         =   "10"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "DMSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest


Dim Deler As Double

Dim Handle As SliderType

Private Type SliderType
    X As Integer
    Y As Integer
End Type
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change(NewValue As Double)


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
If Button = 1 Then
    ValueLabel = Int(X / Deler)
    
    If Val(ValueLabel) < 0 Then ValueLabel = 0
    If Val(ValueLabel) > Val(Max) Then ValueLabel = Max
    
    RaiseEvent Change(Val(ValueLabel))
    
    Handle.X = Val(ValueLabel) * Deler
    If Handle.X < 0 Then Handle.X = 0
    If Handle.X + HandlePicture.Width > UserControl.ScaleWidth Then Handle.X = UserControl.ScaleWidth - HandlePicture.Width
    DrawHandle
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
If Button = 1 Then
    
    ValueLabel = Int(X / Deler)
    
    If Val(ValueLabel) < 0 Then ValueLabel = 0
    If Val(ValueLabel) > Val(Max) Then ValueLabel = Max
    
    RaiseEvent Change(Val(ValueLabel))
    
    Handle.X = Val(ValueLabel) * Deler
    
    If Handle.X < 0 Then Handle.X = 0
    If Handle.X + HandlePicture.Width > UserControl.ScaleWidth Then Handle.X = UserControl.ScaleWidth - HandlePicture.Width
    DrawHandle
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)

If Button = 1 Then
    ValueLabel = Int(X / Deler)
    
    If Val(ValueLabel) < 0 Then ValueLabel = 0
    If Val(ValueLabel) > Val(Max) Then ValueLabel = Max
    
    RaiseEvent Change(Val(ValueLabel))
    
    Handle.X = Val(ValueLabel) * Deler
    
    If Handle.X < 0 Then Handle.X = 0
    If Handle.X + HandlePicture.Width > UserControl.ScaleWidth Then Handle.X = UserControl.ScaleWidth - HandlePicture.Width
    DrawHandle
End If
End Sub

Private Sub DrawHandle()
On Error Resume Next
UserControl.Cls

If CBool(Strech) = True Then
    UserControl.PaintPicture BackPicture.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0, 0, BackPicture.Width, BackPicture.Height
Else
    UserControl.PaintPicture BackPicture.Picture, 0, 0, BackPicture.Width, BackPicture.Height, 0, 0, BackPicture.Width, BackPicture.Height
End If

UserControl.FontName = SliderCaption.FontName
UserControl.FontSize = SliderCaption.FontSize
UserControl.FontBold = SliderCaption.FontBold
UserControl.FontItalic = SliderCaption.FontItalic
UserControl.FontUnderline = SliderCaption.FontUnderline

UserControl.ForeColor = SliderCaption.ForeColor

UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (SliderCaption.Width / 2)
UserControl.CurrentY = (UserControl.ScaleHeight / 2) - (SliderCaption.Height / 2)
UserControl.Print SliderCaption.Caption

If CBool(MaskBool.Caption) = True Then
    BitBlt UserControl.hdc, Handle.X, Handle.Y, HandleMaskPicture.Width, HandleMaskPicture.Height, HandleMaskPicture.hdc, 0, 0, SRCAND
    BitBlt UserControl.hdc, Handle.X, Handle.Y, HandlePicture.Width, HandlePicture.Height, HandlePicture.hdc, 0, 0, SRCPAINT
Else
    BitBlt UserControl.hdc, Handle.X, Handle.Y, HandlePicture.Width, HandlePicture.Height, HandlePicture.hdc, 0, 0, SRCCOPY
End If

UserControl.Refresh

End Sub

Private Sub UserControl_Resize()
Handle.X = Val(ValueLabel * Deler)
Handle.Y = (UserControl.ScaleHeight / 2) - (HandlePicture.Height / 2)
DrawHandle
End Sub

Private Sub UserControl_Show()
Deler = (UserControl.ScaleWidth - HandlePicture.Width) / Val(Max)
Handle.X = Val(ValueLabel) * Deler
DrawHandle
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SliderCaption,SliderCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = SliderCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    SliderCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SliderCaption,SliderCaption,-1,Font
Public Property Get Font() As Font
    Set Font = SliderCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set SliderCaption.Font = New_Font
    PropertyChanged "Font"
    UpdateSlider
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get Border() As Boolean
    Border = UserControl.BorderStyle
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    UserControl.BorderStyle() = IIf(New_Border, 1, 0)
    PropertyChanged "Border"
    UpdateSlider
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SliderCaption,SliderCaption,-1,Caption
Public Property Get Caption() As String
    Caption = SliderCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    SliderCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HandleMaskPicture,HandleMaskPicture,-1,Picture
Public Property Get MaskPicture() As Picture
    Set MaskPicture = HandleMaskPicture.Picture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set HandleMaskPicture.Picture = New_MaskPicture
    PropertyChanged "MaskPicture"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=BackPicture,BackPicture,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = BackPicture.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set BackPicture.Picture = New_Picture
    PropertyChanged "Picture"
    UpdateSlider
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Strech,Strech,-1,Caption
Public Property Get StrechPicture() As Boolean
    StrechPicture = Strech.Caption
End Property

Public Property Let StrechPicture(ByVal New_StrechPicture As Boolean)
    Strech.Caption() = New_StrechPicture
    PropertyChanged "StrechPicture"
    UpdateSlider
End Property


Public Property Get Mask() As Boolean
    Mask = MaskBool.Caption
End Property

Public Property Let Mask(ByVal New_Mask As Boolean)
    MaskBool.Caption() = New_Mask
    PropertyChanged "Mask"
    UpdateSlider
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ValueLabel,ValueLabel,-1,Caption
Public Property Get Value() As Double
    Value = ValueLabel.Caption
End Property

Public Property Let Value(ByVal New_Value As Double)
    If New_Value < 0 Then New_Value = 0
    If New_Value > Val(Max.Caption) Then New_Value = Val(Max.Caption)
    
    ValueLabel.Caption() = New_Value
    
    PropertyChanged "Value"
    UpdateSlider
    
    RaiseEvent Change(New_Value)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Max,Max,-1,Caption
Public Property Get MaxValue() As Double
    MaxValue = Max.Caption
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Double)
    Max.Caption() = New_MaxValue
    PropertyChanged "MaxValue"
    
    Deler = (UserControl.ScaleWidth - HandlePicture.Width) / Val(Max)
    Handle.X = Val(ValueLabel) * Deler
    DrawHandle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HandlePicture,HandlePicture,-1,Picture
Public Property Get HPicture() As Picture
    Set HPicture = HandlePicture.Picture
End Property

Public Property Set HPicture(ByVal New_HPicture As Picture)
    Set HandlePicture.Picture = New_HPicture
    PropertyChanged "HPicture"
    UpdateSlider
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    SliderCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set SliderCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("Border", 0)
    SliderCaption.Caption = PropBag.ReadProperty("Caption", "Slider")
    Set HandleMaskPicture.Picture = PropBag.ReadProperty("MaskPicture", Nothing)
    Strech.Caption = PropBag.ReadProperty("StrechPicture", "true")
    ValueLabel.Caption = PropBag.ReadProperty("Value", "5")
    Max.Caption = PropBag.ReadProperty("MaxValue", "10")
    Set HandlePicture.Picture = PropBag.ReadProperty("HPicture", Nothing)
    MaskBool.Caption = PropBag.ReadProperty("Mask", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", SliderCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", SliderCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Border", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", SliderCaption.Caption, "Slider")
    Call PropBag.WriteProperty("MaskPicture", HandleMaskPicture.Picture, Nothing)
    Call PropBag.WriteProperty("Picture", BackPicture.Picture, Nothing)
    Call PropBag.WriteProperty("StrechPicture", Strech.Caption, "true")
    Call PropBag.WriteProperty("Value", ValueLabel.Caption, "5")
    Call PropBag.WriteProperty("MaxValue", Max.Caption, "10")
    Call PropBag.WriteProperty("Mask", MaskBool.Caption, False)
End Sub


Public Sub UpdateSlider()
Handle.X = Val(ValueLabel * Deler)
Handle.Y = (UserControl.ScaleHeight / 2) - (HandlePicture.Height / 2)
DrawHandle
End Sub
