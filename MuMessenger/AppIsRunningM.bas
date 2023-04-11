Attribute VB_Name = "AppIsRunningM"
Option Explicit
Private Const TH32CS_SNAPPROCESS As Long = 2
Private Const MAX_PATH As Long = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, _
                                                                  ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
                                                        typProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
                                                       typProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function AppIsRunning(ByVal AppName As String) As Boolean
    Dim Process As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim r As Long
    AppName = LCase$(AppName)
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot <> -1 Then
        Process.dwSize = Len(Process)
        r = Process32First(hSnapShot, Process)
        Do While r
            If LCase$(Left$(Process.szExeFile, InStr(1, Process.szExeFile, vbNullChar) - 1)) = AppName Then
                AppIsRunning = True
                r = False
            End If
            r = Process32Next(hSnapShot, Process)
        Loop
        CloseHandle hSnapShot
    End If
End Function



