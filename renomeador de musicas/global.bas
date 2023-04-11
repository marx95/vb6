Attribute VB_Name = "Module1"
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpszClassName As String, ByVal lpszWindow As String) As Long
Global Travar As Integer
