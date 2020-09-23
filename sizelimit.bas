Attribute VB_Name = "SizeLimitMod"
Option Explicit
Public OldWindowProc As Long  ' Original window proc
Public cForm As Long
Public cenabled As Boolean

Public cMinWidth As Single
Public cMinHeight As Single
Public cMaxWidth As Single
Public cMaxHeight As Single

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Function to copy an object/variable/structure passed by reference onto a variable of your own
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Const WM_GETMINMAXINFO = &H24
Private Const GWL_WNDPROC = (-4)

Type POINTAPI
     x As Long
     y As Long
End Type
Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
' Watch for the message to come in
If Msg = WM_GETMINMAXINFO Then
  
  Dim MinMax As MINMAXINFO
  
  CopyMemory MinMax, ByVal lp, Len(MinMax)
  If cMinWidth > 0 Then MinMax.ptMinTrackSize.x = cMinWidth
  If cMinHeight > 0 Then MinMax.ptMinTrackSize.y = cMinHeight
  If cMaxWidth > 0 Then MinMax.ptMaxTrackSize.x = cMaxWidth
  If cMaxHeight > 0 Then MinMax.ptMaxTrackSize.y = cMaxHeight
  
  CopyMemory ByVal lp, MinMax, Len(MinMax)
  
  ' This tells Windows that the message was handled successfully
  WndProc = 1
  Exit Function

End If

' Forward all messages on to the default message handler as well
WndProc = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)

End Function

Public Sub hook()
If cenabled = True And OldWindowProc = 0 Then
  OldWindowProc = GetWindowLong(cForm, GWL_WNDPROC)
  SetWindowLong cForm, GWL_WNDPROC, AddressOf WndProc
End If
End Sub

Public Sub unhook()
If OldWindowProc = 0 Then Exit Sub
SetWindowLong cForm, GWL_WNDPROC, OldWindowProc
OldWindowProc = 0
End Sub
