Attribute VB_Name = "ChildWnd"
Option Explicit
Private Declare Function EnumChildWindows Lib "User32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private WinArr As String
Private Function EnumChildProc(ByVal Hwnd As Long, ByVal lParam As Long) As Long
 If WinArr = "" Then
  WinArr = Hwnd
 Else
  WinArr = WinArr & vbCrLf & Hwnd
 End If
EnumChildProc = 1
End Function

Public Function GetChildWin(Win_hwnd As Long) As String
WinArr = ""
EnumChildWindows Win_hwnd, AddressOf EnumChildProc, ByVal 0&
GetChildWin = WinArr
End Function

