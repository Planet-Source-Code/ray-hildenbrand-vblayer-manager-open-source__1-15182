VERSION 5.00
Begin VB.Form frmObjects 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Contact Sheet Designer "
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "frmObjects.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHandle 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   0
      Left            =   30
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   -15
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   3480
      Picture         =   "frmObjects.frx":0322
      ScaleHeight     =   600
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   3750
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   4
      Left            =   6435
      Picture         =   "frmObjects.frx":10D4
      ScaleHeight     =   825
      ScaleWidth      =   2070
      TabIndex        =   1
      Top             =   2715
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1365
      Index           =   3
      Left            =   750
      Picture         =   "frmObjects.frx":2344
      ScaleHeight     =   1365
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   2250
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1470
      Index           =   1
      Left            =   285
      Picture         =   "frmObjects.frx":3B43
      ScaleHeight     =   1470
      ScaleWidth      =   2040
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   2
      Left            =   3345
      Picture         =   "frmObjects.frx":5A74
      ScaleHeight     =   945
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   615
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1290
      Index           =   0
      Left            =   3450
      Picture         =   "frmObjects.frx":71D2
      ScaleHeight     =   1290
      ScaleWidth      =   1890
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   800
      Index           =   6
      Left            =   1875
      Picture         =   "frmObjects.frx":8A78
      ScaleHeight     =   795
      ScaleWidth      =   2055
      TabIndex        =   6
      Top             =   5085
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   800
      Index           =   7
      Left            =   6405
      Picture         =   "frmObjects.frx":A021
      ScaleHeight     =   795
      ScaleWidth      =   2025
      TabIndex        =   7
      Top             =   4500
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   8
      Left            =   3735
      Picture         =   "frmObjects.frx":B52F
      ScaleHeight     =   1020
      ScaleWidth      =   2025
      TabIndex        =   8
      Top             =   6510
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   9
      Left            =   6330
      Picture         =   "frmObjects.frx":CDC6
      ScaleHeight     =   1440
      ScaleWidth      =   2175
      TabIndex        =   9
      Top             =   6210
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mTrans(9) As clsTransForm
Dim Initialized As Boolean
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Windows declarations
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI
Dim i As Integer
Private Sub Form_Activate()
    
    If Not Initialized Then
        Set frmTbar.LayerContainer1.ActiveDesigner = Me
        Load frmTbar
        For i = 0 To Picture1.UBound
              Dim tmpObj As Control
              Set tmpObj = frmObjects.Picture1(i)
              frmTbar.LayerContainer1.AddLayerItem tmpObj, "Image " & tmpObj.Index + 1 & ""
              Set mTrans(i) = New clsTransForm
              mTrans(i).ShapeMe RGB(255, 255, 255), True, , tmpObj
              DoEvents
              Picture1(i).Visible = True
        Next
          
        Set tmpObj = Nothing
   
        
        frmTbar.Show , Me
        Initialized = True
       ' Me.BackColor = &H87B568
    End If
    
End Sub

Private Sub Form_Load()
    DragInit
    
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

    If Button = vbLeftButton Then  'And m_bDesignMode
        'Hit test over light-weight (non-windowed) controls
        For i = 0 To (Controls.Count - 1)
            'Check for visible, non-menu controls
            '[Note 1]
            'If any of the sizing handle controls are under the mouse
            'pointer, then they must not be visible or else they would
            'have already intercepted the MouseDown event.
            '[Note 2]
            'This code will fail if you have a control such as the
            'Timer control which has no Visible property. You will
            'either need to make sure your form has no such controls
            'or add code to handle them.
            If Not TypeOf Controls(i) Is Menu And Controls(i).Visible Then
                m_DragRect.SetRectToCtrl Controls(i)
                If m_DragRect.PtInRect(X, Y) Then
                    DragBegin Controls(i)
                    Exit Sub
                End If
            End If
        Next i
        'No control is active
        Set m_CurrCtl = Nothing
        'Hide sizing handles
        ShowHandles False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nWidth As Single, nHeight As Single
    Dim pt As POINTAPI

    If m_DragState = StateDragging Then
        'Save dimensions before modifying rectangle
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Update drag rectangle coordinates
        m_DragRect.Left = pt.X - m_DragPoint.X
        m_DragRect.Top = pt.Y - m_DragPoint.Y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
        'Draw new rectangle
        DrawDragRect
    ElseIf m_DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Action depends on handle being dragged
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = pt.X
                m_DragRect.Top = pt.Y
            Case 1
                m_DragRect.Top = pt.Y
            Case 2
                m_DragRect.Right = pt.X
                m_DragRect.Top = pt.Y
            Case 3
                m_DragRect.Right = pt.X
            Case 4
                m_DragRect.Right = pt.X
                m_DragRect.Bottom = pt.Y
            Case 5
                m_DragRect.Bottom = pt.Y
            Case 6
                m_DragRect.Left = pt.X
                m_DragRect.Bottom = pt.Y
            Case 7
                m_DragRect.Left = pt.X
        End Select
        'Draw new rectangle
        DrawDragRect
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If m_DragState = StateDragging Or m_DragState = StateSizing Then
            'Hide drag rectangle
            DrawDragRect
            'Move control to new location
            On Error Resume Next ''when in doubt
            m_DragRect.ScreenToTwips m_CurrCtl
            m_DragRect.SetCtrlToRect m_CurrCtl
            'Restore sizing handles
            ShowHandles True
            'Free mouse movement
            ClipCursor ByVal 0&
            'Release mouse capture
            ReleaseCapture
            'Reset drag state
            m_DragState = StateNothing
        End If
    End If
End Sub

Private Sub DragInit()
    Dim i As Integer, xHandle As Single, yHandle As Single

    'Use black Picture box controls for 8 sizing handles
    'Calculate size of each handle
    xHandle = 7 * Screen.TwipsPerPixelX
    yHandle = 7 * Screen.TwipsPerPixelY
    'Load array of handles until we have 8
    For i = 0 To 7
        If i <> 0 Then
            Load picHandle(i)
           
            
            
        End If
        picHandle(i).Width = xHandle
        picHandle(i).Height = yHandle
        'Must be in front of other controls
        picHandle(i).ZOrder
        
    Next i
    'Set mousepointers for each sizing handle
    picHandle(0).MousePointer = vbSizeNWSE
    picHandle(1).MousePointer = vbSizeNS
    picHandle(2).MousePointer = vbSizeNESW
    picHandle(3).MousePointer = vbSizeWE
    picHandle(4).MousePointer = vbSizeNWSE
    picHandle(5).MousePointer = vbSizeNS
    picHandle(6).MousePointer = vbSizeNESW
    picHandle(7).MousePointer = vbSizeWE
    'Initialize current control
    Set m_CurrCtl = Nothing
End Sub
Public Sub DragBegin(ctl As Control)
    Dim rc As RECT

    'Hide any visible handles
    ShowHandles False
    'Save reference to control being dragged
    Set m_CurrCtl = ctl
    'Store initial mouse position
    GetCursorPos m_DragPoint
    'Save control position (in screen coordinates)
    'Note: control might not have a window handle
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    'Make initial mouse position relative to control
    m_DragPoint.X = m_DragPoint.X - m_DragRect.Left
    m_DragPoint.Y = m_DragPoint.Y - m_DragRect.Top
    'Force redraw of form without sizing handles
    'before drawing dragging rectangle
    Refresh
    'Show dragging rectangle
    DrawDragRect
    'Indicate dragging under way
    m_DragState = StateDragging
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    ReleaseCapture  'This appears needed before calling SetCapture
    SetCapture hwnd
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub
Private Sub DragEnd()
    'm_CurrCtl.Visible = True
    Set m_CurrCtl = Nothing
    ShowHandles False
    m_DragState = StateNothing
End Sub

Private Sub Form_Resize()
    CheckerBoard
End Sub

Private Sub LayerContainer1_Click()

End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
        'Me.Visible = False
        frmTbar.LayerContainer1.SelectLayer Index
        DragBegin Picture1(Index)
    End If
End Sub
'Process MouseDown over handles
Private Sub picHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim rc As RECT

    'Handles should only be visible when a control is selected
    Debug.Assert (Not m_CurrCtl Is Nothing)
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    'Track index handle
    m_DragHandle = Index
    'Hide sizing handles
    ShowHandles False
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    'Indicate sizing is under way
    m_DragState = StateSizing
    'Show sizing rectangle
    DrawDragRect
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hwnd
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub

'Display or hide the sizing handles and arrange them for the current rectangld
Private Sub ShowHandles(Optional bShowHandles As Boolean = True)
    Dim i As Integer
    Dim xFudge As Long, yFudge As Long
    Dim nWidth As Long, nHeight As Long

    If bShowHandles And Not m_CurrCtl Is Nothing Then
        With m_DragRect
            'Save some calculations in variables for speed
            nWidth = (picHandle(0).Width \ 2)
            nHeight = (picHandle(0).Height \ 2)
            xFudge = (0.5 * Screen.TwipsPerPixelX)
            yFudge = (0.5 * Screen.TwipsPerPixelY)
            'Top Left
            picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge
            'Bottom right
            picHandle(4).Move (.Left + .Width) - nWidth - xFudge, .Top + .Height - nHeight - yFudge
            'Top center
            picHandle(1).Move .Left + (.Width / 2) - nWidth, .Top - nHeight + yFudge
            'Bottom center
            picHandle(5).Move .Left + (.Width / 2) - nWidth, .Top + .Height - nHeight - yFudge
            'Top right
            picHandle(2).Move .Left + .Width - nWidth - xFudge, .Top - nHeight + yFudge
            'Bottom left
            picHandle(6).Move .Left - nWidth + xFudge, .Top + .Height - nHeight - yFudge
            'Center right
            picHandle(3).Move .Left + .Width - nWidth - xFudge, .Top + (.Height / 2) - nHeight
            'Center left
            picHandle(7).Move .Left - nWidth + xFudge, .Top + (.Height / 2) - nHeight
        End With
    End If
    'Show or hide each handle
    For i = 0 To 7
        picHandle(i).Visible = bShowHandles
       
    Next i
End Sub

'Draw drag rectangle. The API is used for efficiency and also
'because drag rectangle must be drawn on the screen DC in
'order to appear on top of all controls
Private Sub DrawDragRect()
    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    Dim hScreenDC As Long, nDrawMode As Long

    'Get DC of entire screen in order to
    'draw on top of all controls
    hScreenDC = GetDC(0)
    'Select GDI object
    hPen = CreatePen(PS_SOLID, 2, 0)
    hOldPen = SelectObject(hScreenDC, hPen)
    hBrush = GetStockObject(NULL_BRUSH)
    hOldBrush = SelectObject(hScreenDC, hBrush)
    nDrawMode = SetROP2(hScreenDC, R2_NOT)
    'Draw rectangle
    Rectangle hScreenDC, m_DragRect.Left, m_DragRect.Top, _
        m_DragRect.Right, m_DragRect.Bottom
    'Restore DC
    SetROP2 hScreenDC, nDrawMode
    SelectObject hScreenDC, hOldBrush
    SelectObject hScreenDC, hOldPen
    ReleaseDC 0, hScreenDC
    'Delete GDI objects
    DeleteObject hPen
End Sub



Public Sub StartDrag(ByRef objWhich As Object)
      DragBegin objWhich
End Sub


Public Sub CheckerBoard()
    Dim X As Integer
    Dim Y As Integer
    Dim counter As Integer
    'Dim MainCounter As Integer
    Dim MainCounter As Long
    Me.Cls
  For Y = 0 To Me.ScaleHeight 'Step ((Me.ScaleWidth / Screen.TwipsPerPixelX) / 16)
    MainCounter = MainCounter + 1
    For X = 0 To Me.ScaleWidth 'Step (Me.ScaleWidth / Screen.TwipsPerPixelX) / 16
        
        If CInt(MainCounter) Mod 2 = 0 Then
            If CInt(counter) Mod 2 = 0 Then
                Me.Line (X, Y)-(X + (Screen.TwipsPerPixelX * 15), Y + (Screen.TwipsPerPixelY * 15)), vbWhite, BF
            Else
                Me.Line (X, Y)-(X + (Screen.TwipsPerPixelX * 15), Y + (Screen.TwipsPerPixelY * 15)), &HE0E0E0, BF
            End If
        Else
            If CInt(counter) Mod 2 = 0 Then
                Me.Line (X, Y)-(X + (Screen.TwipsPerPixelX * 15), Y + (Screen.TwipsPerPixelY * 15)), &HE0E0E0, BF
            Else
                Me.Line (X, Y)-(X + (Screen.TwipsPerPixelX * 15), Y + (Screen.TwipsPerPixelY * 15)), vbWhite, BF
            End If
        
        End If
        counter = counter + 1
        X = X + Screen.TwipsPerPixelX * 16
      
    Next
        counter = 0
        Y = Y + Screen.TwipsPerPixelX * 16
     
  Next
End Sub
