Attribute VB_Name = "OnTopMod"
Option Explicit
Private Declare Sub SetWindowPos Lib "User32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Function OnTop(Hwnd As Long, Top As Boolean)
SetWindowPos Hwnd, Abs(Top) - 2, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Function

