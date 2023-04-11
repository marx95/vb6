Attribute VB_Name = "Processos"

Public Const PROCESS_ALL_ACCESS As Long = 4096
Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Long = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Public Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal blnheritHandle As Long, _
    ByVal dwAppProcessId As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" ( _
    ByVal hSnapshot As Long, _
    uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" ( _
    ByVal hSnapshot As Long, _
    uProcess As PROCESSENTRY32) As Long

Public Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
    ByVal lFlags As Long, _
    lProcessID As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32.dll" ( _
    ByVal ApphProcess As Long, _
    ByVal uExitCode As Long) As Long



Public Function ProcessoExiste(ByVal processo As String) As Boolean
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim pos As Integer

  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
  
    Do While rProcessFound
        pos = InStr(1, uProcess.szexeFile, vbNullChar)
        szExename = Left$(uProcess.szexeFile, pos - 1)
        
        If UCase$(szExename) = UCase$(processo) Then
            ProcessoExiste = True
            Exit Do
        End If
      
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
End Function








