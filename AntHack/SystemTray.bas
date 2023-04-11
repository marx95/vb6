Attribute VB_Name = "SystemTray"
Option Explicit

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
"Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204

Public Sub VaiPraTray(hWnd As Long, Icone As Long, ToolTip As String)
Dim IconeTray As NOTIFYICONDATA
IconeTray.cbSize = Len(IconeTray)
IconeTray.hWnd = hWnd
IconeTray.uID = 1&
IconeTray.uFlags = NIF_DOALL
IconeTray.uCallbackMessage = WM_MOUSEMOVE
IconeTray.hIcon = Icone
IconeTray.szTip = ToolTip & Chr$(0)
Call Shell_NotifyIcon(NIM_ADD, IconeTray)
End Sub

Public Sub ForaDaTray(hWnd As Long)
Dim IconeTray As NOTIFYICONDATA
IconeTray.cbSize = Len(IconeTray)
IconeTray.hWnd = hWnd
IconeTray.uID = 1&
Call Shell_NotifyIcon(NIM_DELETE, IconeTray)
End Sub



