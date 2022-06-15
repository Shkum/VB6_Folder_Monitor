Attribute VB_Name = "GetWinText_Class"
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Const WM_GETTEXT = &HD

Public Function GetWinText(Win_hWnd As Long)
Dim cl As String * 10240
GetWinText = Left(cl, GetWindowText(Win_hWnd, cl, 256))
End Function

Public Function GetWinClass(Win_hWnd As Long)
Dim cl As String * 10240
GetWinClass = Left(cl, GetClassName(Win_hWnd, cl, 256))
End Function

Public Function GetWinText_Msg(Win_hWnd As Long)
Dim cl As String * 10240
GetWinText_Msg = Left(cl, SendMessage(Win_hWnd, WM_GETTEXT, ByVal 10240, ByVal cl))
End Function
