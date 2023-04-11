Attribute VB_Name = "OtmizaProcesso"
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Const HIGH_PRIORITY_CLASS = &H80
Const IDLE_PRIORITY_CLASS = &H40
Const NORMAL_PRIORITY_CLASS = &H20
Const REALTIME_PRIORITY_CLASS = &H100
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hthread As _
Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As _
Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hthread As _
Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As _
Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Function Otimizar()
If Environ("NUMBER_OF_PROCESSORS") = 1 Then
    'retrieve the current thread and process
    hthread = GetCurrentThread
    hProcess = GetCurrentProcess
    'set the new thread priority to "lowest"
    SetThreadPriority hthread, THREAD_PRIORITY_ABOVE_NORMAL
    'set the new priority class to "idle"
    SetPriorityClass hProcess, HIGH_PRIORITY_CLASS
End If
End Function




