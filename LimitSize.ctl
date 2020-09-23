VERSION 5.00
Begin VB.UserControl LimitSizing 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   345
   InvisibleAtRuntime=   -1  'True
   Picture         =   "LimitSize.ctx":0000
   ScaleHeight     =   315
   ScaleWidth      =   345
   ToolboxBitmap   =   "LimitSize.ctx":0672
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Size Limiter"
      Height          =   195
      Left            =   3390
      TabIndex        =   0
      Top             =   3105
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "LimitSizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
Width = 345
Height = 325
End Sub

Public Property Get Formhwnd() As Long
Formhwnd = cForm
End Property

Public Property Let Formhwnd(ByVal vNewValue As Long)
unhook
cForm = vNewValue
hook
End Property

Public Property Get Enabled() As Boolean
Enabled = cenabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
cenabled = vNewValue
If cenabled = False Then
  unhook
Else
  hook
End If
End Property
Public Property Get MinWidth() As Single
MinWidth = cMinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Single)
  cMinWidth = New_MinWidth
  PropertyChanged "MinWidth"
End Property

Public Property Get MinHeight() As Single
  MinHeight = cMinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Single)
  cMinHeight = New_MinHeight
  PropertyChanged "MinHeight"
End Property

Public Property Get MaxWidth() As Single
  MaxWidth = cMaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Single)
  cMaxWidth = New_MaxWidth
  PropertyChanged "MaxWidth"
End Property

Public Property Get MaxHeight() As Single
  MaxHeight = cMaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Single)
  cMaxHeight = New_MaxHeight
  PropertyChanged "MaxHeight"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  cMinWidth = PropBag.ReadProperty("MinWidth", cMinWidth)
  cMinHeight = PropBag.ReadProperty("MinHeight", cMinHeight)
  cMaxWidth = PropBag.ReadProperty("MaxWidth", cMaxWidth)
  cMaxHeight = PropBag.ReadProperty("MaxHeight", cMaxHeight)
  
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("MinWidth", cMinWidth)
  Call PropBag.WriteProperty("MinHeight", cMinHeight)
  Call PropBag.WriteProperty("MaxWidth", cMaxWidth)
  Call PropBag.WriteProperty("MaxHeight", cMaxHeight)
End Sub

Public Sub About()

End Sub
