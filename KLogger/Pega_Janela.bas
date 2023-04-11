Attribute VB_Name = "Pega_Janela"
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpszClassName As String, ByVal lpszWindow As String) As Long

Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String
    Dim TextLength As Long
    TextLength = GetWindowTextLength(WindowHandle)
    Buffer$ = String(TextLength, 0)
    Call GetWindowText(WindowHandle, Buffer, TextLength + 1)
    GetCaption$ = Buffer$
End Function
