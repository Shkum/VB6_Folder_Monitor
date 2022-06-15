Attribute VB_Name = "WinFromPnt"
Option Explicit
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Function WinFromCursor() As Long
Dim s As POINTAPI
GetCursorPos s
WinFromCursor = WindowFromPoint(s.x, s.y)
End Function
