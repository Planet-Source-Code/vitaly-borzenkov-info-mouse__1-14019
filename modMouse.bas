Attribute VB_Name = "modMouse"
'API calls
'this call is for finding out the current positon of the mouse cursor *
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'this call lets you find out the hwnd property of the current window *
Public Declare Function WindowFromPoint Lib "user32" _
(ByVal xPoint As Long, ByVal yPoint As Long) As Long
'call for getting the title or name of a window
Public Declare Function GetClassName Lib "user32" _
Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long

'type that is used by 'GetCursorPos'
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'* what I have mentioned is not all that that function can do, but this is what
'I use it for in this code
