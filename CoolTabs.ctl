VERSION 5.00
Begin VB.UserControl CoolTabs 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   3900
   Begin VB.Image ImageSingle 
      Height          =   285
      Left            =   1260
      Picture         =   "CoolTabs.ctx":0000
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image ImageOrg 
      Height          =   285
      Left            =   2460
      Picture         =   "CoolTabs.ctx":0950
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image ImageReplace 
      Height          =   285
      Left            =   120
      Picture         =   "CoolTabs.ctx":0DAB
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image btnRound 
      Height          =   225
      Left            =   3600
      Picture         =   "CoolTabs.ctx":16FB
      Top             =   30
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   4140
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2580
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   3
      Left            =   2520
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1680
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   1
      Left            =   840
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image GreyBorder 
      Height          =   285
      Left            =   0
      Picture         =   "CoolTabs.ctx":1A8D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4035
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4140
      Y1              =   290
      Y2              =   290
   End
End
Attribute VB_Name = "CoolTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event RoundBtnClick(ActiveTabIs As Integer)
Event TabClick(Index As Integer)

Const m_def_VisibleTabs = 3
Const m_def_ActiveTab = 0
Const m_def_Licensecode = ""

Dim m_VisibleTabs As Long
Dim m_ActiveTab As Integer
Dim m_Licensecode As String

Enum Tabbs
    Tab1 = 0
    Tab2 = 1
    Tab3 = 2
    Tab4 = 3
End Enum



Public Property Get Licensecode() As String
    Licensecode = m_Licensecode
End Property
Public Property Let Licensecode(ByVal New_Licensecode As String)
    m_Licensecode = New_Licensecode
    PropertyChanged "Licensecode"
End Property

Sub SetCapBold(Index As Integer)
Dim i As Long

SetOrgPics

For i = 0 To m_VisibleTabs
    Label1(i).FontBold = False
    Label1(i).ForeColor = &H80000011
    Label1(i).Visible = True
Next

If Index <> m_VisibleTabs Then
    Image1(Index).Picture = ImageReplace.Picture
    Image1(Index).ZOrder
    DoEvents
Else
    Image1(Index).Picture = ImageSingle.Picture
    Image1(Index).ZOrder
    DoEvents
End If

Label1(Index).FontBold = True
Label1(Index).ForeColor = vbBlack
Label1(Index).ZOrder

End Sub


Sub SetOrgPics()
Dim i As Long


For i = 0 To m_VisibleTabs
    Image1(i).Picture = ImageOrg.Picture
    Image1(i).ZOrder
    Image1(i).Visible = True
    Label1(i).ZOrder
    DoEvents
Next

End Sub

Private Sub btnRound_Click()
RaiseEvent RoundBtnClick(ActiveTab)
End Sub

Private Sub btnRound_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
btnRound.Left = btnRound.Left + 15
btnRound.Top = btnRound.Top + 15

DoEvents

End Sub


Private Sub btnRound_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
btnRound.Left = (Width - btnRound.Width) - 70
btnRound.Top = 30
DoEvents

End Sub


Private Sub Label1_Click(Index As Integer)
m_ActiveTab = Index
SetCapBold Index
RaiseEvent TabClick(Index)

End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Label1(Index).BackStyle = 0
Label1(Index).BackColor = &H8000000D
Label1(Index).ForeColor = vbBlack

End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Label1(Index).BackStyle = 0
Label1(Index).BackColor = &H8000000F
Label1(Index).ForeColor = vbBlack

End Sub


Private Sub UserControl_AmbientChanged(PropertyName As String)

If PropertyName = "BackColor" Then
    Parent.BackColor = &HE0E0E0
End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1(0).Caption = PropBag.ReadProperty("CaptionTab1", "Label1")
    Label1(1).Caption = PropBag.ReadProperty("CaptionTab2", "Label2")
    Label1(2).Caption = PropBag.ReadProperty("CaptionTab3", "Label3")
    Label1(3).Caption = PropBag.ReadProperty("CaptionTab4", "Label4")
    m_VisibleTabs = PropBag.ReadProperty("VisibleTabs", m_def_VisibleTabs)
    Label1(0).Caption = PropBag.ReadProperty("CaptionTab1", "Label1")
    Label1(1).Caption = PropBag.ReadProperty("CaptionTab2", "Label1")
    Label1(2).Caption = PropBag.ReadProperty("CaptionTab3", "Label1")
    Label1(3).Caption = PropBag.ReadProperty("CaptionTab4", "Label1")
    Label1(0).Caption = PropBag.ReadProperty("CaptionTab1", "Label1")
    m_Licensecode = PropBag.ReadProperty("Licensecode", m_def_Licensecode)
    m_ActiveTab = PropBag.ReadProperty("ActiveTab", m_def_ActiveTab)

End Sub

Private Sub UserControl_Resize()
On Error Resume Next

Line1.X2 = Width
Line2.X2 = Width
GreyBorder.Width = Width
btnRound.Left = (Width - btnRound.Width) - 70

If Width < 3855 Then Width = 3855


End Sub

Private Sub UserControl_Show()

Parent.BackColor = &HE0E0E0
SetOrgPics
Label1_Click (0)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CaptionTab1", Label1(0).Caption, "Label1")
    Call PropBag.WriteProperty("CaptionTab2", Label1(1).Caption, "Label2")
    Call PropBag.WriteProperty("CaptionTab3", Label1(2).Caption, "Label3")
    Call PropBag.WriteProperty("CaptionTab4", Label1(3).Caption, "Label4")
    Call PropBag.WriteProperty("VisibleTabs", m_VisibleTabs, m_def_VisibleTabs)
    Call PropBag.WriteProperty("Licensecode", m_Licensecode, m_def_Licensecode)
    Call PropBag.WriteProperty("ActiveTab", m_ActiveTab, m_def_ActiveTab)

    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,3
Public Property Get VisibleTabs() As Tabbs
Attribute VisibleTabs.VB_Description = "Sets the total amount of visible tabs..."
    VisibleTabs = m_VisibleTabs
End Property
Public Property Get ActiveTab() As Tabbs
    ActiveTab = m_ActiveTab
End Property
Public Property Let VisibleTabs(ByVal New_VisibleTabs As Tabbs)
    Dim i As Long
    
    If New_VisibleTabs < 4 Then
        m_VisibleTabs = New_VisibleTabs
        PropertyChanged "VisibleTabs"
        
        For i = 0 To 3
            Label1(i).Visible = False
            Image1(i).Visible = False
        Next
        
        Label1_Click (0)
    Else
        MsgBox "Invalid number...", 48
    End If


    
End Property
Public Property Let ActiveTab(ByVal New_ActibeTabs As Tabbs)
    m_ActiveTab = New_ActibeTabs
    PropertyChanged "ActiveTab"
    Label1_Click m_ActiveTab
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_CaptionTab1 = m_def_CaptionTab1
'    m_CaptionTab2 = m_def_CaptionTab2
'    m_CaptionTab3 = m_def_CaptionTab3
'    m_CaptionTab4 = m_def_CaptionTab4
    m_VisibleTabs = m_def_VisibleTabs
    
    m_Licensecode = m_def_Licensecode
    'm_Licensecode = InputBox("Ange din licenskod:", "Licenskontroll för denna kontroll")

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(1),Label1,1,Caption
Public Property Get CaptionTab2() As String
Attribute CaptionTab2.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionTab2 = Label1(1).Caption
End Property

Public Property Let CaptionTab2(ByVal New_CaptionTab2 As String)
    Label1(1).Caption() = New_CaptionTab2
    PropertyChanged "CaptionTab2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(2),Label1,2,Caption
Public Property Get CaptionTab3() As String
Attribute CaptionTab3.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionTab3 = Label1(2).Caption
End Property

Public Property Let CaptionTab3(ByVal New_CaptionTab3 As String)
    Label1(2).Caption() = New_CaptionTab3
    PropertyChanged "CaptionTab3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(3),Label1,3,Caption
Public Property Get CaptionTab4() As String
Attribute CaptionTab4.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionTab4 = Label1(3).Caption
End Property

Public Property Let CaptionTab4(ByVal New_CaptionTab4 As String)
    Label1(3).Caption() = New_CaptionTab4
    PropertyChanged "CaptionTab4"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1(0),Label1,0,Caption
Public Property Get CaptionTab1() As String
Attribute CaptionTab1.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CaptionTab1 = Label1(0).Caption
End Property

Public Property Let CaptionTab1(ByVal New_CaptionTab1 As String)
    Label1(0).Caption() = New_CaptionTab1
    PropertyChanged "CaptionTab1"
End Property

