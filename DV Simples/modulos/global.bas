Attribute VB_Name = "global"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Global Preload As Long
Global Source As String
Global LinkNavegador As String

Global LinksDV(1 To 32767) As String
Global TotalLinks As Long
Global LinkAgora As Long

Public Function Navegar(URL As String)
    On Error Resume Next
    Divulgador.m_WebControl.object.navigate URL
    Divulgador.SalvarBT.Enabled = False
    Divulgador.LinkInfo.Caption = URL
End Function

Public Function Carregar_Links()
    TotalLinks = ReadINI(App.Path & "/Data/LinksDB.ini", "0", "Total")
    
    If TotalLinks = 0 Then Exit Function
    Dim Total As Long
    
    For i = 1 To TotalLinks
        Dim Link As String
        Link = ReadINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L")
        If Len(Link) > 6 Then
            If LinkAgora = 0 Then LinkAgora = 1
            Total = Total + 1
            LinksDV(CInt(Total)) = ReadINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L")
            
        End If
    Next i

    'TotalLinks = Total
End Function

Public Function Adicionar_Link()
    Divulgador.SalvarBT.Enabled = False

    For i = 1 To TotalLinks
        Dim Link As String
        Link = ReadINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L")
        If Len(Link) < 6 Then
            Call WriteINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L", LinkNavegador)
            Exit Function
        End If
    Next i
    
    TotalLinks = TotalLinks + 1
    Call WriteINI(App.Path & "/Data/LinksDB.ini", "0", "Total", CStr(TotalLinks))
    Call WriteINI(App.Path & "/Data/LinksDB.ini", CStr(TotalLinks), "L", LinkNavegador)
End Function
