Attribute VB_Name = "Mdlapi"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Global Const EM_SETPASSWORDCHAR = &HCC
Global Const Chr_Star_Ascii = 42 'ascii value of the * char
Global pt2 As POINTAPI
Global Wnd
Global Etick
Global bEnd As Boolean
