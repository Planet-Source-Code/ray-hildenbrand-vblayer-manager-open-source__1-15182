VERSION 5.00
Begin VB.UserControl LayerItem 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   LockControls    =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   3600
   ToolboxBitmap   =   "LayerItem.ctx":0000
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   105
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picVisible 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   90
      MouseIcon       =   "LayerItem.ctx":0312
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   90
      Width           =   285
      Begin VB.Image Eye 
         Height          =   480
         Left            =   -90
         MouseIcon       =   "LayerItem.ctx":0BDC
         MousePointer    =   99  'Custom
         Picture         =   "LayerItem.ctx":14A6
         Top             =   -105
         Width           =   480
      End
   End
   Begin VB.Image imgClosed 
      Height          =   480
      Left            =   1320
      Picture         =   "LayerItem.ctx":1D70
      Top             =   1635
      Width           =   480
   End
   Begin VB.Image imgOpen 
      Height          =   480
      Left            =   825
      Picture         =   "LayerItem.ctx":263A
      Top             =   1635
      Width           =   480
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Object 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   510
      TabIndex        =   2
      Top             =   105
      Width           =   3045
   End
   Begin VB.Shape shpBorder 
      Height          =   450
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3600
   End
   Begin VB.Shape shpInnerBorder 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   450
      Top             =   90
      Width           =   3075
   End
   Begin VB.Label lblBG 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "LayerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event Click() 'MappingInfo=lblBG,lblBG,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lblBG,lblBG,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblBG,lblBG,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblBG,lblBG,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblBG,lblBG,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Public EyeCleared As Boolean
Public OwnedControl As Control
Public Event isSelected(Button As Integer)
Private mHighlightColor As OLE_COLOR
Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR
Private mForeHighlightColor As OLE_COLOR

Private nSelected As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblBG,lblBG,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mBackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = mForeColor
    
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    mForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get ForeHighlightColor() As OLE_COLOR
    ForeHighlightColor = mForeHighlightColor
    
End Property

Public Property Let ForeHighlightColor(ByVal New_ForehighlightColor As OLE_COLOR)
    
    mForeHighlightColor = New_ForehighlightColor
    PropertyChanged "ForeHighlightColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblBG,lblBG,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = lblBG.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    lblBG.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Eye_Click()
    If Eye.Picture = imgOpen.Picture Then
        Set Eye.Picture = imgClosed.Picture
        If Not TypeName(OwnedControl) = "Nothing" Then
            OwnedControl.Visible = False
        End If
    Else
        Set Eye.Picture = imgOpen.Picture
        If Not TypeName(OwnedControl) = "Nothing" Then
            OwnedControl.Visible = True
        End If
    End If
    
    
    
End Sub

Private Sub Eye_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    EyeCleared = False
    DrawDownButton picVisible, picVisible.Height / Screen.TwipsPerPixelY, picVisible.Width / Screen.TwipsPerPixelX
End Sub

Private Sub Eye_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    EyeCleared = False
    DrawButton picVisible, picVisible.Height / Screen.TwipsPerPixelY, picVisible.Width / Screen.TwipsPerPixelX
End Sub

Private Sub Eye_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not EyeCleared Then
        Cls
        picVisible.Refresh
    End If
    EyeCleared = False
    DrawButton picVisible, picVisible.Height / Screen.TwipsPerPixelY, picVisible.Width / Screen.TwipsPerPixelX
End Sub


Private Sub lblBG_Click()
    'RaiseEvent isselected button
    RaiseEvent Click
End Sub

Private Sub lblBG_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = mHighlightColor
End Property
Public Property Let HighlightColor(vnewColor As OLE_COLOR)
    mHighlightColor = vnewColor
    
End Property

Private Sub lblCaption_DblClick()
    txtEdit.Text = lblCaption.Caption
    
    txtEdit.Visible = True
    txtEdit.BackColor = mHighlightColor
    
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent isSelected(Button)
    RaiseEvent MouseDown(Button, Shift, x, y)
    lblBG.BackColor = HighlightColor
    lblCaption.ForeColor = mForeHighlightColor
    lblCaption.FontBold = True
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not EyeCleared Then
        Cls
        picVisible.Refresh
        EyeCleared = True
    End If
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Caption = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
Caption = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub UserControl_Initialize()
    Set Eye.Picture = imgOpen.Picture
    NotSelected = True
End Sub

Private Sub UserControl_InitProperties()
    lblCaption.FontBold = False
    lblCaption.FontItalic = False
    
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lblBG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent isSelected(Button)
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    lblBG.BackColor = mHighlightColor
    lblCaption.ForeColor = mForeHighlightColor
    lblCaption.FontBold = True
End Sub

Private Sub lblBG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
    If Not EyeCleared Then
        Cls
        picVisible.Refresh
        EyeCleared = True
    End If
End Sub

Private Sub lblBG_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
   
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
    picVisible.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not EyeCleared Then
        Cls
        picVisible.Refresh
        EyeCleared = True
    End If
End Sub

Private Sub UserControl_Resize()
    lblBG.Width = UserControl.Width - lblBG.Left - 90
    shpBorder.Width = UserControl.Width - shpBorder.Left - 15
    shpInnerBorder.Width = UserControl.Width - shpInnerBorder.Left - 125
    lblCaption.Width = UserControl.Width - lblCaption.Left - 100
    
    RaiseEvent Resize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mHighlightColor = PropBag.ReadProperty("HighlightColor", &H8000000D)
    mBackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    'lblBG.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    mForeColor = PropBag.ReadProperty("ForeColor", mForeColor)
    mForeHighlightColor = PropBag.ReadProperty("ForeHighlightColor", vbWhite)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblBG.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Object 1")
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", 8)
    lblCaption.FontName = PropBag.ReadProperty("FontName", "Tahoma")
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", False)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", False)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HighlightColor", HighlightColor, &H8000000D)
    Call PropBag.WriteProperty("ForeHighlightColor", mForeHighlightColor, vbWhite)
    Call PropBag.WriteProperty("BackColor", mBackColor, &HC0C0C0)
    Call PropBag.WriteProperty("ForeColor", mForeColor, vbBlack)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", lblBG.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, -2147483640)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Object 1")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   
    
End Sub



Public Property Get NotSelected() As Boolean
    NotSelected = nSelected
End Property

Public Property Let NotSelected(ByVal vNewValue As Boolean)
    nSelected = vNewValue
    If vNewValue = True Then
        
        lblBG.BackColor = vbWhite
        lblCaption.ForeColor = mForeColor
        
        lblCaption.FontBold = False
    Else
        lblBG.BackColor = mHighlightColor
        lblCaption.ForeColor = mForeHighlightColor
        
        lblCaption.FontBold = True
    End If
End Property

Public Sub IsVisible(bIs As Boolean)
      If Not bIs Then
        Set Eye.Picture = imgClosed.Picture
        
    Else
        Set Eye.Picture = imgOpen.Picture
'        If Not TypeName(OwnedControl) = "Nothing" Then
'            OwnedControl.Visible = True
'        End If
    End If
    
    
End Sub
