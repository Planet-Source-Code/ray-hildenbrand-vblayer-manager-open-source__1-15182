Attribute VB_Name = "modMail"
Public Const email = "ray.hildenbrand@verizon.net"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1


Public Sub sendemail()
Dim Success As Long
    Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Public Sub AccessWeb()
    Dim Success As Long
    Success = ShellExecute(0&, vbNullString, "http://www.vbaccelerator.com", vbNullString, "C:\", SW_SHOWNORMAL)
End Sub
