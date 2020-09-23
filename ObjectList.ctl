VERSION 5.00
Begin VB.UserControl LayerContainer 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   3900
   Begin VB.VScrollBar VScroll 
      Height          =   3735
      Left            =   3630
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picTBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   3900
      TabIndex        =   1
      Top             =   3750
      Width           =   3900
      Begin VB.PictureBox picButtons 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   540
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   30
         Width           =   225
         Begin VB.Image imgZorder 
            Height          =   480
            Left            =   -135
            Picture         =   "ObjectList.ctx":0000
            ToolTipText     =   "Bring To Front"
            Top             =   -120
            Width           =   480
         End
      End
      Begin VB.PictureBox picButtons 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   285
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   3
         Top             =   30
         Width           =   225
         Begin VB.Image imgDup 
            Height          =   480
            Left            =   -150
            Picture         =   "ObjectList.ctx":08CA
            ToolTipText     =   "Duplicate Layer"
            Top             =   -120
            Width           =   480
         End
      End
      Begin VB.PictureBox picButtons 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   45
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   2
         Top             =   30
         Width           =   225
         Begin VB.Image imgTrash 
            Height          =   480
            Left            =   -135
            Picture         =   "ObjectList.ctx":1194
            ToolTipText     =   "Delete Layer"
            Top             =   -135
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picScroll 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   3615
      TabIndex        =   5
      Top             =   0
      Width           =   3615
      Begin vbpLayerItem.LayerItem Layer 
         DragIcon        =   "ObjectList.ctx":1A5E
         Height          =   480
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Visible         =   0   'False
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   847
         HighlightColor  =   11493445
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         FontName        =   "Tahoma"
         MousePointer    =   99
         MouseIcon       =   "ObjectList.ctx":2328
      End
      Begin VB.Image imgDrop 
         Height          =   480
         Left            =   1620
         Picture         =   "ObjectList.ctx":2C02
         Top             =   1425
         Width           =   480
      End
   End
End
Attribute VB_Name = "LayerContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public TotalLayers As Integer
Private ButtonCleared As Boolean
Dim i As Integer
Public SelectedLayer As Object
Private mBackColor As OLE_COLOR
Private mHighlightColor As OLE_COLOR
'Event Declarations:
Event Click() 'MappingInfo=picScroll,picScroll,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=picScroll,picScroll,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picScroll,picScroll,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picScroll,picScroll,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picScroll,picScroll,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Scroll() 'MappingInfo=VScroll,VScroll,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public ActiveDesigner As Form

Public Event LayerDuplicated(ByRef objectToDuplicate As Object)


''All of the layer manager and layer item code found in this example is copyright Ray Hildenbrand 2001''''
''''this code was posted as an example and no liability can be had by its author
'''' if you use any of this code inan application, you will need to ensure that the proper credit has been given to ray hildenbrand on an about form or something like that.
''''' please do not steal this code and say it was yours''''

''''''ray hildenbrand '''''''''''''''

Public Sub DuplicateLayer(ByRef objWhich As Object, Optional Selectedindex As Integer = -1)
   If Not Selectedindex = -1 Then
    
        If Not TotalLayers = 0 Then
            TotalLayers = TotalLayers + 1
            Load Layer((Layer.UBound) + 1)
            Layer(Layer.UBound).Move Layer(Layer.UBound - 1).Left, Layer(Layer.UBound - 1).Top + Layer(Layer.UBound - 1).Height + 15
            Layer(Layer.UBound).Visible = True
            Set Layer(Layer.UBound).OwnedControl = objWhich.OwnedControl
            Layer(Layer.UBound).Caption = "Copy of " & objWhich.OwnedControl.Name & "(" & Selectedindex & ")"
            RaiseEvent LayerDuplicated(objWhich)
        Else
            TotalLayers = TotalLayers + 1
            Layer(Layer.UBound).Move 0, 0
            Layer(Layer.UBound).Visible = True
            Set Layer(Layer.UBound).OwnedControl = objWhich.OwnedControl
            Layer(Layer.UBound).Caption = "Copy of " & objWhich.OwnedControl.Name & "(" & Selectedindex & ")"
        End If
        picScroll.Height = Layer(Layer.UBound).Top + Layer(Layer.UBound).Height
        
        If Layer(Layer.UBound).Top + Layer(Layer.UBound).Height > UserControl.Height Then
            VScroll.Visible = True
            VScroll.Move UserControl.Width - VScroll.Width, 0, VScroll.Width, UserControl.Height
                VScroll.Max = picScroll.Height - UserControl.Height + picTBar.Height
                VScroll.LargeChange = 145
                VScroll.SmallChange = 60
        Else
            VScroll.Visible = False
        End If
        
        Layer(Layer.UBound).DragMode = vbManual

    Else
              If Not TotalLayers = 0 Then
            TotalLayers = TotalLayers + 1
            Load Layer((Layer.UBound) + 1)
            Layer(Layer.UBound).Move Layer(Layer.UBound - 1).Left, Layer(Layer.UBound - 1).Top + Layer(Layer.UBound - 1).Height + 15
            Layer(Layer.UBound).Visible = True
            Set Layer(Layer.UBound).OwnedControl = objWhich
            Layer(Layer.UBound).Caption = "Copy of " & objWhich.OwnedControl.Name & "(" & objWhich.Index & ")"
            
        Else
            TotalLayers = TotalLayers + 1
            Layer(Layer.UBound).Move 0, 0
            Layer(Layer.UBound).Visible = True
            Set Layer(Layer.UBound).OwnedControl = objWhich
            Layer(Layer.UBound).Caption = "Copy of " & objWhich.OwnedControl.Name & "(" & objWhich.Index & ")"
        End If
        picScroll.Height = Layer(Layer.UBound).Top + Layer(Layer.UBound).Height
        
        If Layer(Layer.UBound).Top + Layer(Layer.UBound).Height > UserControl.Height Then
            VScroll.Visible = True
            VScroll.Move UserControl.Width - VScroll.Width, 0, VScroll.Width, UserControl.Height
                VScroll.Max = picScroll.Height - UserControl.Height + picTBar.Height
                VScroll.LargeChange = 145
                VScroll.SmallChange = 60
        Else
            VScroll.Visible = False
        End If
        
        Layer(Layer.UBound).DragMode = vbManual
    End If
End Sub
Public Sub AddLayerItem(ByRef objWhich As Object, Caption As String)
    If Not TotalLayers = 0 Then
        TotalLayers = TotalLayers + 1
        Load Layer((Layer.UBound) + 1)
        Layer(Layer.UBound).Move Layer(Layer.UBound - 1).Left, Layer(Layer.UBound - 1).Top + Layer(Layer.UBound - 1).Height + 15
        Layer(Layer.UBound).Visible = True
        Set Layer(Layer.UBound).OwnedControl = objWhich
        Layer(Layer.UBound).Caption = Caption
        
    Else
        TotalLayers = TotalLayers + 1
        Layer(Layer.UBound).Move 0, 0
        Layer(Layer.UBound).Visible = True
        Set Layer(Layer.UBound).OwnedControl = objWhich
        Layer(Layer.UBound).Caption = Caption
    End If
    picScroll.Height = Layer(Layer.UBound).Top + Layer(Layer.UBound).Height
    
    If Layer(Layer.UBound).Top + Layer(Layer.UBound).Height > (UserControl.Height - picTBar.Height) Then
        VScroll.Visible = True
        VScroll.Move UserControl.Width - VScroll.Width, 0, VScroll.Width, UserControl.Height - VScroll.Top - picTBar.Height
            VScroll.Max = picScroll.Height - UserControl.Height + picTBar.Height
            VScroll.LargeChange = 450
            VScroll.SmallChange = 190
    Else
        VScroll.Visible = False
    End If
    
    Layer(Layer.UBound).DragMode = vbManual

End Sub

Private Sub imgDup_Click()
    If Not TypeName(SelectedLayer) = "Nothing" Then
            Dim i As Integer
            For i = 0 To Layer.UBound
                If Not Layer(i).NotSelected Then
                    DuplicateLayer SelectedLayer, i
                    Exit For
                End If
                
            Next
        End If
        
End Sub

Private Sub imgDup_DragDrop(Source As Control, x As Single, y As Single)
    
    DuplicateLayer Source
End Sub

Private Sub imgDup_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ButtonCleared = False
    For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    DrawDownButton picButtons(1), picButtons(1).Height / Screen.TwipsPerPixelY, picButtons(1).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgDup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    DrawDownButton picButtons(1), picButtons(1).Height / Screen.TwipsPerPixelY, picButtons(1).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgDup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    DrawButton picButtons(1), picButtons(1).Height / Screen.TwipsPerPixelY, picButtons(1).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgDup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonCleared Then
        For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    End If
    ButtonCleared = False
    DrawButton picButtons(1), picButtons(1).Height / Screen.TwipsPerPixelY, picButtons(1).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgTrash_Click()
    If Not TypeName(SelectedLayer) = "Nothing" Then
        Dim i As Integer
        For i = 0 To Layer.UBound
            If Not Layer(i).NotSelected Then
                DoDelete SelectedLayer, i
                Exit For
            End If
            
        Next
    End If
        
End Sub

Private Sub imgTrash_DragDrop(Source As Control, x As Single, y As Single)
   DoDelete Source
 
End Sub

Private Sub imgTrash_DragOver(Source As Control, x As Single, y As Single, State As Integer)
 ButtonCleared = False
    picButtons(1).Cls
    DrawDownButton picButtons(0), picButtons(0).Height / Screen.TwipsPerPixelY, picButtons(0).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgTrash_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    picButtons(1).Cls
    DrawDownButton picButtons(0), picButtons(0).Height / Screen.TwipsPerPixelY, picButtons(0).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgTrash_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    picButtons(1).Cls
    DrawButton picButtons(0), picButtons(0).Height / Screen.TwipsPerPixelY, picButtons(0).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgTrash_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonCleared Then
        For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    End If
    ButtonCleared = False
    DrawButton picButtons(0), picButtons(0).Height / Screen.TwipsPerPixelY, picButtons(0).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgZorder_Click()
    
            Dim i As Integer
            For i = 0 To Layer.UBound
                If Layer(i).NotSelected Then
                    Layer(i).ZOrder 1
                Else
                        'Layer(i).ZOrder 1

                End If

            Next
 

  '  ZorderObjects
End Sub

Private Sub imgZorder_DragDrop(Source As Control, x As Single, y As Single)
    'Source.OwnedControl.ZOrder 0
    ZorderObjects
End Sub

Private Sub imgZorder_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ButtonCleared = False
    picButtons(1).Cls
    DrawDownButton picButtons(2), picButtons(2).Height / Screen.TwipsPerPixelY, picButtons(2).Width / Screen.TwipsPerPixelX
End Sub

Private Sub Layer_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Dim tmpVisible As Boolean
    Dim tmpOwned As Object
    Dim tmpCaption As String
    Dim tmpelement As Integer
    
    
    tmpelement = Source.Index
    If Not Index = tmpelement Then
        tmpCaption = Layer(Index).Caption
        Set tmpOwned = Layer(Index).OwnedControl
        tmpVisible = Layer(Index).OwnedControl.Visible
        Layer(Index).Caption = Layer(tmpelement).Caption
        Set Layer(Index).OwnedControl = Layer(tmpelement).OwnedControl
        Layer(Index).OwnedControl.Visible = Layer(tmpelement).OwnedControl.Visible
        Layer(Index).IsVisible Layer(tmpelement).OwnedControl.Visible
        
        Layer(tmpelement).Caption = tmpCaption
        Set Layer(tmpelement).OwnedControl = tmpOwned
        Layer(tmpelement).OwnedControl.Visible = tmpVisible
        Layer(tmpelement).IsVisible tmpVisible
    
    End If
    
    
    ZorderObjects
End Sub

Private Sub Layer_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  ' If Not Source.Index = Index Then Source.DragIcon = imgDrop.Picture
End Sub


Private Sub Layer_isSelected(Index As Integer, Button As Integer)
    If Button = 1 Then
        Dim i As Integer
        Set SelectedLayer = Layer(Index)
        Layer(Index).NotSelected = False
        For i = 0 To Layer.UBound
            If i <> Index Then Layer(i).NotSelected = True
        Next
        Dim tmpObj As Object
        Set tmpObj = Layer(Index).OwnedControl
        ActiveDesigner.startdrag tmpObj
        Set tmpObj = Nothing
    End If
End Sub

Private Sub Layer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Layer(Index).DragMode = vbAutomatic
    Else
        Layer(Index).DragMode = vbManual
    End If
End Sub

Private Sub Layer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    For i = 0 To Layer.UBound
        If Not Layer(i).EyeCleared Then
            If Not i = Index Then Layer(i).Cls
        End If
    Next
End Sub

Private Sub Layer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Layer(Index).DragMode = vbManual
        Layer(Index).Drag
    End If
End Sub



Private Sub picScroll_Click()
    RaiseEvent Click
    Set SelectedLayer = Nothing
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mBackColor = New_BackColor
    picScroll.BackColor = mBackColor
    UserControl.BackColor = mBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = mHighlightColor
End Property

Public Property Let HighlightColor(ByVal New_highlightColor As OLE_COLOR)
    mHighlightColor = New_highlightColor
    For i = 0 To Layer.UBound
        Layer(i).HighlightColor = mHighlightColor
    Next
    PropertyChanged "highlightColor"
End Property
Private Sub picScroll_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
        
    For i = 0 To Layer.UBound
        If Not Layer(i).EyeCleared Then
             Layer(i).Cls
        End If
    Next
End Sub

Private Sub picScroll_Resize()
    For i = 0 To Layer.UBound
        Layer(i).Width = picScroll.Width
    Next
End Sub

Private Sub picTBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Not ButtonCleared Then
        For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
                ButtonCleared = True
        Next
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonCleared Then
        For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
                ButtonCleared = True
        Next
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    HighlightColor = PropBag.ReadProperty("highlightColor", &H8000000D)
    
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    'If UserControl.Width < 3900 Then UserControl.Width = 3900
    
    ResizeScrollBar
End Sub

Private Sub UserControl_Terminate()
    Set SelectedLayer = Nothing
    
    
End Sub

Private Sub VScroll_Change()
    picScroll.Top = -VScroll.Value
End Sub

Public Sub ZorderObjects()
    For i = 0 To Layer.UBound
        On Error Resume Next
        If i = 0 Then
            Layer(i).OwnedControl.ZOrder
        Else
            Layer(i).OwnedControl.ZOrder i
        End If
    Next
End Sub

Private Sub imgzorder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    picButtons(1).Cls
    DrawDownButton picButtons(2), picButtons(2).Height / Screen.TwipsPerPixelY, picButtons(2).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgzorder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ButtonCleared = False
    picButtons(0).Cls
    DrawButton picButtons(2), picButtons(2).Height / Screen.TwipsPerPixelY, picButtons(2).Width / Screen.TwipsPerPixelX
End Sub

Private Sub imgzorder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonCleared Then
        For i = 0 To picButtons.UBound
                picButtons(i).Cls
                picButtons(i).Refresh
        Next
    End If
    ButtonCleared = False
    DrawButton picButtons(2), picButtons(2).Height / Screen.TwipsPerPixelY, picButtons(2).Width / Screen.TwipsPerPixelX
End Sub

Public Sub DoDelete(mSource As Object, Optional Selectedindex As Integer = -1)
    Dim tmpIndex As Integer
    Dim tmpCount As Integer
    If Selectedindex = -1 Then
        tmpIndex = mSource.Index
        tmpCount = Layer.UBound
    
        If mSource.Index = 0 Then Exit Sub
        TotalLayers = TotalLayers - 1
        If TotalLayers = 0 Then Exit Sub
        For i = tmpIndex To Layer.UBound
            If Not i + 1 > tmpCount Then
              Layer(i).OwnedControl = Layer(i + 1).OwnedControl
              Layer(i).Caption = Layer(i + 1).Caption
            End If
        Next
        
        Unload Layer(Layer.UBound)
        'Set Layer(tmpCount) = Nothing
        For i = 0 To Layer.UBound
            If i = 0 Then
                Layer(i).Move 0, 0
            Else
                Layer(i).Move Layer(i).Left, Layer(i - 1).Top + Layer(i - 1).Height + 15
            End If
        Next
        Set SelectedLayer = Layer(0)
    Else
        tmpIndex = Selectedindex
        tmpCount = Layer.UBound
    
        If Selectedindex = -1 Or Selectedindex = 0 Then Exit Sub
        TotalLayers = TotalLayers - 1
        If TotalLayers = 0 Then Exit Sub
        For i = tmpIndex To Layer.UBound
            If Not i + 1 > tmpCount Then
              Layer(i).OwnedControl = Layer(i + 1).OwnedControl
              Layer(i).Caption = Layer(i + 1).Caption
            End If
        Next
        
        Unload Layer(Layer.UBound)
        'Set Layer(tmpCount) = Nothing
        For i = 0 To Layer.UBound
            If i = 0 Then
                Layer(i).Move 0, 0
            Else
                Layer(i).Move Layer(i).Left, Layer(i - 1).Top + Layer(i - 1).Height + 15
            End If
        Next
        Layer(Layer.UBound).NotSelected = False
        Set SelectedLayer = Layer(Layer.UBound)
    End If
    
    ResizeScrollBar
End Sub

Public Sub ResizeScrollBar()
    picScroll.Height = Layer(Layer.UBound).Top + Layer(Layer.UBound).Height
    picScroll.Width = UserControl.Width - VScroll.Width
    If Layer(Layer.UBound).Top + Layer(Layer.UBound).Height > UserControl.Height Then
        VScroll.Visible = True
        On Error Resume Next
        VScroll.Move UserControl.Width - VScroll.Width, 0, VScroll.Width, UserControl.Height - VScroll.Top - picTBar.Height
        VScroll.Max = picScroll.Height - UserControl.Height + picTBar.Height
        VScroll.LargeChange = VScroll.Max / TotalLayers
        VScroll.SmallChange = VScroll.Max / (TotalLayers / 2)
    Else
        VScroll.Visible = False
    End If
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

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
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "Returns a handle (from Microsoft Windows) to the window a UserControl is contained in."
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Private Sub picScroll_DblClick()
    RaiseEvent DblClick
End Sub

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

Private Sub picScroll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub picScroll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub VScroll_Scroll()
    RaiseEvent Scroll
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("highlightColor", mHighlightColor, &H8000000D)
    Call PropBag.WriteProperty("Enabled", mBackColor, &H8000000F)
End Sub


Public Sub SelectLayer(iIndex As Integer)
    If Not iIndex > Layer.UBound Then
        For i = 0 To Layer.UBound
            If Not iIndex = i Then
               Layer(i).NotSelected = True
            Else
                Layer(i).NotSelected = False
            End If
        Next
        
    End If
    
End Sub

Public Sub DeselectAll()
    For i = i To Layer.UBound
        Layer(i).NotSelected = True
    Next
End Sub
