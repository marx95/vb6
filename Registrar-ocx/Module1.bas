Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0

Public Function RegistrarComponente(hWnd As Long, LibreriaDeseada As String, Registrar As Boolean) As Boolean
 Dim libAddress As Long, procAddress As Long
 
 ' Obtener la direccion de la libreria indicada
 libAddress = LoadLibrary(LibreriaDeseada)
 If libAddress = 0 Then
 RegistrarComponente = False
 Exit Function
 End If
 
 ' Obtener de la libreria, la direccion del server para Registrar o DesRegistrar
 If Registrar Then
 procAddress = GetProcAddress(libAddress, "DllRegisterServer")
 Else
 procAddress = GetProcAddress(libAddress, "DllUnregisterServer")
 End If
 ' Ejecutar el Registro/DesRegistro de la libreria
 
 If CallWindowProc(procAddress, hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
 RegistrarComponente = True
 Else
 RegistrarComponente = False
 End If

 ' Liberar la libreria utilizada
 FreeLibrary libAddress
 
End Function
