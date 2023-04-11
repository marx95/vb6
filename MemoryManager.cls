VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This module could be used for a number of things,
' mostly it is used for memory based hacks and add-
' ons for games and applications. Somewhat messy,
' it has been hacked together throughout the years
'
' - rancid ( www.r-cid.com )

Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private hPid As Long
Private hProcess As Long

Public Property Let pid(dwPid As Long)
    hPid = dwPid
    hProcess = OpenProcess(&H1F0FFF, False, hPid)
End Property

Public Property Get pid() As Long
    pid = hPid
End Property

Public Sub writeByte(lngAddress As Long, bVal As Byte)
    WritePtr lngAddress, VarPtr(bVal), 1
End Sub

Public Sub writeInteger(lngAddress As Long, intValue As Integer)
    WritePtr lngAddress, VarPtr(intValue), 2
End Sub

Public Sub writeSingle(lngAddress As Long, sngValue As Single)
    WritePtr lngAddress, VarPtr(sngValue), 4
End Sub

Public Sub writeLong(lngAddress As Long, lngValue As Long)
    WritePtr lngAddress, VarPtr(lngValue), 4
End Sub

Public Sub WriteString(lngAddress As Long, strValue As String)
    Dim lngBytesWritten As Long
    WriteProcessMemory hProcess, ByVal lngAddress, ByVal strValue, Len(strValue), lngBytesWritten
End Sub

Public Function WritePtr(lngAddress As Long, destPtr As Long, length As Long) As Long
    Dim lngBytesWritten As Long
    WriteProcessMemory hProcess, ByVal lngAddress, ByVal destPtr, length, lngBytesWritten
    WritePtr = lngBytesWritten
End Function

Public Function copyToType(destPtr As Long, length As Long, lngAddr As Long) As Boolean
    copyToType = (ReadProcessMemory(hProcess, ByVal lngAddr, ByVal destPtr, length, &H0) <> 0)
End Function

Public Function readLong(lngAddr As Long) As Long
    copyToType VarPtr(readLong), 4, lngAddr
End Function

Public Function readInteger(lngAddr As Long) As Integer
    copyToType VarPtr(readInteger), 2, lngAddr
End Function

Public Function readSingle(lngAddr As Long) As Single
    copyToType VarPtr(readSingle), 4, lngAddr
End Function

Public Function readByte(lngAddr As Long) As Byte
    copyToType VarPtr(readByte), 1, lngAddr
End Function

'//strings ends at null characters
Public Function readString(length As Long, lngAddr As Long) As String
    Dim strTmp As String, i As Long
     
    strTmp = String(length, 0)

    If (ReadProcessMemory(hProcess, ByVal lngAddr, ByVal strTmp, length, &H0) <> 0) Then
        i = InStr(strTmp, Chr(0))
         
        If i > 0 Then strTmp = Left(strTmp, i - 1)
         
        readString = strTmp
         
      '  lngBaseAddr = lngBaseAddr + length
    End If
End Function

'//string contains null characters
Public Function readStringN(length As Long, lngAddr As Long) As String
    Dim strTmp As String, i As Long
     
    strTmp = String(length, 0)

    If (ReadProcessMemory(hProcess, ByVal lngAddr, ByVal strTmp, length, &H0) <> 0) Then
        readStringN = strTmp
    End If
End Function

