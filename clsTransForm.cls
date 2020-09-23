VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Doug Gaede
'October 31, 2000
'----------------------------------------------------------
'Notes:
'Originally based on code by Chris Yates (Automatic Form Shaper) from www.planetsourcecode.com,
'but modified completely and only shares about 5 lines of code with the original.  This version runs
'much faster and will work on either a Form or PictureBox (for making custom-shaped controls).
'Should work with any object that has an hDC, but I have restricted it for my own use.
'My test form image (not the one in Form1 now, but about the same size and complexity) took
'~35 seconds to create on a PIII-650 with Chris's code.
'I cut it down to ~3 seconds (10X gain) by cutting the CombineRgn calls down from many, many thousands
'to a few hundred in many cases.
'This was done by searching for adjacent points and combining them into one line.
'I also replaced the code to move the form because it didn't work in NT.
'I have only tested this code in NT4, SP4.

'You are free to use, modify and distribute this code.
'Email me at dgaede@home.com with comments or questions.
'----------------------------------------------------------

'This module causes a Form or PictureBox to become shaped depending on the picture
'that is assigned to the Picture property.  One color in the picture
'will become the transparent color, depending on the color values passed to ShapeMe.
'The DragForm sub allows the user to drag a form that doesn't have a title bar.

'Note that you MUST set the Form or PictureBox's BorderStyle property
'equal to 0, manually.  This can not be set in code because it is read-only at runtime.

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const RGN_DIFF = 4

Dim CurRgn As Long, TempRgn As Long  ' Region variables

'For dragging the form
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Sub ShapeMe(Color As Long, HorizontalScan As Boolean, Optional Name1 As Form = Nothing, Optional Name2 As PictureBox = Nothing)

'Color = the color to convert to transparent (easiest to use RGB function to pass in this value)
'HorizontalScan = scan for transparent lines horizonally or vertically.  Try both during development and pick the fastest one.
'Name1 = a Form name.  If used, leave Name2 blank.
'Name2 = a PictureBox name.  If used, leave Name1 blank.

Dim X As Integer, Y As Integer 'points on form
Dim dblHeight As Double, dblWidth As Double 'height and width of object
Dim lngHDC As Long 'the hDC property of the object
Dim booMiddleOfSet As Boolean 'used during the gathering of transparent points
Dim colPoints As Collection 'this will hold all usrPoints
Set colPoints = New Collection
Dim Z As Variant 'used during iteration through collection
Dim dblTransY As Double 'these 3 variables hold each point that will be made transparent
Dim dblTransStartX As Double
Dim dblTransEndX As Double
Dim Name As Object 'will hold the name of the object.  Late-bound and slower, but allows different types (in this case Form or PictureBox)

'check out the name or names passed into the subroutine
If Name1 Is Nothing Xor Name2 Is Nothing Then 'we know there is a name in one of them
    If Name1 Is Nothing Then 'set the name
        Set Name = Name2
    Else
        Set Name = Name1
    End If
Else 'both or none hold valid names
    MsgBox "Must pass in the name of either a Form OR a PictureBox.  TransForm received NONE or BOTH.  Function failed.", vbOKOnly, "ShapeMe Subroutine"
    Exit Sub
End If

'initialization
With Name
    .AutoRedraw = True 'object must have this setting
    .ScaleMode = 3 'object must have this setting
    lngHDC = .hdc 'faster to use a variable; VB help recommends using the property, but I didn't encounter any problems
    If HorizontalScan = True Then 'look for lines of transparency horizontally
        dblHeight = .ScaleHeight 'faster to use a variable
        dblWidth = .ScaleWidth 'faster to use a variable
    Else 'look vertically (note that the names "dblHeight" and "dblWidth" are non-sensical now, but this was an easy way to do this
        dblHeight = .ScaleWidth 'faster to use a variable
        dblWidth = .ScaleHeight 'faster to use a variable
    End If 'HorizontalScan = True
End With
booMiddleOfSet = False

'gather all points that need to be made transparent
For Y = 0 To dblHeight  ' Go through each column of pixels on form
    dblTransY = Y
    For X = 0 To dblWidth  ' Go through each line of pixels on form
        'note that using GetPixel appears to be faster than using VB's Point
        If TypeOf Name Is Form Then 'check to see if this is a form and use GetPixel function which is a little faster
            If GetPixel(lngHDC, X, Y) = Color Then  ' If the pixel's color is the transparency color, record it
                If booMiddleOfSet = False Then
                    dblTransStartX = X
                    dblTransEndX = X
                    booMiddleOfSet = True
                Else
                    dblTransEndX = X
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'GetPixel(lngHDC, X, Y) = Color
         ElseIf TypeOf Name Is PictureBox Then 'if a PictureBox then use Point; a little slower but works when GetPixel doesn't
            If Name.Point(X, Y) = Color Then
                If booMiddleOfSet = False Then
                    dblTransStartX = X
                    dblTransEndX = X
                    booMiddleOfSet = True
                Else
                    dblTransEndX = X
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'Name.Point(X, Y) = Color
        End If 'TypeOf Name Is Form
        
    Next X
Next Y

CurRgn = CreateRectRgn(0, 0, dblWidth, dblHeight)  ' Create base region which is the current whole window

For Each Z In colPoints 'now make it transparent
    TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)  ' Create a temporary pixel region for this pixel
    CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
    DeleteObject (TempRgn)  ' Delete the temporary region and free resources
Next

SetWindowRgn Name.hwnd, CurRgn, True  ' Finally set the windows region to the final product
'I do not use DeleteObject on the CurRgn, going with the advice in Dan Appleman's book:
'once set to a window using SetWindowRgn, do not delete the region.

Set colPoints = Nothing

End Sub
Public Sub DragForm(Who As Form, intButton As Integer)

On Error Resume Next

If intButton = vbLeftButton Then
    'Move the borderless form...
    Call ReleaseCapture
    Call SendMessage(Who.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If

End Sub





