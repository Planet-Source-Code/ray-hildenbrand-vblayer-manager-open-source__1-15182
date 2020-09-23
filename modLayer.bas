Attribute VB_Name = "modLayer"
Option Explicit
Const BF_SOFT = &H1000
'Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)


Public Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type
Const BDR_INNER = &HC
Const BDR_OUTER = &H3
Const BDR_RAISED = &H5
Const BDR_RAISEDINNER = &H4
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKEN = &HA
Const BDR_SUNKENINNER = &H8
Const BDR_SUNKENOUTER = &H2
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Public Sub DrawButton(objWhich As Object, iheight As Integer, iwidth As Integer)
    Dim RECT As RECT
    
    RECT.Left = 0
    RECT.Top = 0
    RECT.Right = iwidth
    RECT.Bottom = iheight
    DrawEdge objWhich.hdc, RECT, BDR_RAISEDOUTER, BF_SOFT Or BF_RECT  'Or BF_MIDDLE
End Sub

Public Sub DrawDownButton(objWhich As Object, iheight As Integer, iwidth As Integer)
    Dim RECT As RECT
    
    RECT.Left = 0
    RECT.Top = 0
    RECT.Right = iwidth
    RECT.Bottom = iheight
    DrawEdge objWhich.hdc, RECT, BDR_SUNKEN, BF_SOFT Or BF_RECT  'Or BF_MIDDLE
End Sub

'Sub Main()
'    Dim tmpDate As Variant
'    tmpDate = CStr(GetSetting("LayerMan", "Environment", "FR", ""))
'
'    If tmpDate = "12/29/1899" Then
'       SaveSetting "LayerMan", "Environment", "FR", CStr(Date)
'    Else
'        tmpDate = CDate(tmpDate)
'        If (tmpDate + 30) < Date Then
'            MsgBox "Controls are not ina release state yet and an update should be made. Email Ray.Hildenbrand@verizon.net to obtain the most current version.", vbInformation, App.Title
'            'End
'        End If
'
'
'    End If
'End Sub
