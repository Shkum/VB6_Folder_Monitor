VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Enabler"
   ClientHeight    =   6720
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   15750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   15750
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Hide when start"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hide selected window"
      Height          =   735
      Left            =   6840
      TabIndex        =   30
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show selected window"
      Height          =   735
      Left            =   5520
      TabIndex        =   29
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2175
      Left            =   4260
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disable selected window"
      Height          =   735
      Left            =   4200
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable selected window"
      Height          =   735
      Left            =   2880
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   150
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   690
      Width           =   3000
   End
   Begin VB.TextBox Text3 
      Height          =   1905
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Form1.frx":08CA
      Top             =   1230
      Width           =   3000
   End
   Begin VB.TextBox Text4 
      Height          =   2325
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Form1.frx":08D0
      Top             =   3330
      Width           =   3000
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "Form1.frx":08D6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   5760
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST button"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8760
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4905
      ItemData        =   "Form1.frx":0BE0
      Left            =   8040
      List            =   "Form1.frx":0BE2
      TabIndex        =   7
      Top             =   120
      Width           =   7575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   5400
      Width           =   2000
   End
   Begin VB.TextBox Text6 
      Height          =   2580
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":0BE4
      Top             =   2640
      Width           =   2000
   End
   Begin VB.TextBox Text5 
      Height          =   2145
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0BEA
      Top             =   360
      Width           =   2000
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   255
      Left            =   12240
      TabIndex        =   27
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   255
      Left            =   12240
      TabIndex        =   26
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   12240
      TabIndex        =   25
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Text:"
      Height          =   255
      Left            =   10680
      TabIndex        =   24
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Class:"
      Height          =   255
      Left            =   10680
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Selected window:"
      Height          =   255
      Left            =   10680
      TabIndex        =   22
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "WINDOWLONG:"
      Height          =   315
      Left            =   6000
      TabIndex        =   19
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label6 
      Caption         =   "CHILD:"
      Height          =   315
      Left            =   5040
      TabIndex        =   18
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "GWL_ID:"
      Height          =   315
      Left            =   4920
      TabIndex        =   17
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000017&
      Caption         =   "Mouse Up to stop"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000017&
      Caption         =   "Mouse Down to start"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   5400
      Width           =   5235
   End
   Begin VB.Label Label4 
      Caption         =   "SENDMSG:"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "GETTEXT:"
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   1230
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "CLASS:"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "HWND:"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "User32" (ByVal Hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const GWL_HINSTANCE = -6
Private Const GWL_HWNDPARENT = -8
Private Const GWL_ID = -12
Private Const GWL_STYLE = -16
Private Const GWL_USERDATA = -21
Private Const GWL_WNDPROC = -4
Const WM_GETTEXT = &HD
Private Const LB_SETTABSTOPS = &H192


Private Sub Command1_Click()
MsgBox "test"
End Sub

Private Sub Command2_Click()
On Error Resume Next
EnableWindow Label14, True
End Sub

Private Sub Command3_Click()
On Error Resume Next
EnableWindow Label14, False
End Sub

Private Sub Command4_Click()
On Error Resume Next
ShowWindow Label14, 5
End Sub

Private Sub Command5_Click()
On Error Resume Next
ShowWindow Label14, 0
End Sub

Private Sub Form_Load()
Dim mTabs(0) As Long
mTabs(0) = 100
SendMessage List1.Hwnd, LB_SETTABSTOPS, 1, mTabs(0)
OnTop Me.Hwnd, True
Text1 = "0"
Text8 = "gwl_exstyle:" & vbCrLf & "gwl_hinstance:" & vbCrLf & "gwl_hwndparent:" & vbCrLf & "gwl_id:" & vbCrLf & "gwl_style:" & vbCrLf & "gwl_userdata:" & vbCrLf & "gwl_wndproc:"
End Sub


Private Sub List1_Click()
Label14 = Left(List1.Text, InStr(1, List1.Text, vbTab, vbTextCompare) - 1)
Label15 = GetWinClass(Label14)
Label16 = GetWinText_Msg(Label14)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LstArr As String
Picture1.MousePointer = Button + 1
If Button = 1 Then
Dim s As Long, cl As String * 10240
s = WinFromCursor
Text2 = GetWinClass(s)
Text3 = GetWinText(s)
Text7 = GetWindowLong(s, GWL_ID)
Text4 = Left(cl, SendMessage(s, WM_GETTEXT, ByVal 10240, ByVal cl))
Text5 = GetWindowLong(s, GWL_EXSTYLE) & vbCrLf & GetWindowLong(s, GWL_HINSTANCE) & vbCrLf & GetWindowLong(s, GWL_HWNDPARENT) & vbCrLf & GetWindowLong(s, GWL_ID) & vbCrLf & GetWindowLong(s, GWL_STYLE) & vbCrLf & GetWindowLong(s, GWL_USERDATA) & vbCrLf & GetWindowLong(s, GWL_WNDPROC)
 If Text1 <> s Then
  List1.Clear
  Text6 = GetChildWin(s)
  Dim WinArr() As String, Cntr As Integer, HW As Long
 If GetChildWin(s) <> "" Then
 WinArr() = Split(GetChildWin(s), vbCrLf)
    For Cntr = LBound(WinArr) To UBound(WinArr)
    HW = WinArr(Cntr)
    List1.AddItem HW & vbTab & GetWinText(HW) & vbTab & GetWinClass(HW)
    Next
 End If
 End If
Text1 = s
Label8 = "ChildWindows: " & List1.ListCount
End If
Label14 = Text1
Label15 = Text2
Label16 = Text4
If Check1.Value = 1 And Button = 1 Then Me.WindowState = 1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 0
End Sub
