Attribute VB_Name = "Global"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global LeftInicial As Long
Global Delay_Desaparecer As Long
Global WinMode As Long
Global UltimaMsg As Long
Global MuMsgLink As String
Global Source As String

Public Function Sumir()
    MuMsg.Deslizar.Enabled = False
    MuMsg.Visible = False
End Function

Public Function Tranparencia()
    Dim bytOpacity As Byte       'Set the transparency level
    bytOpacity = 180
    Call SetWindowLong(MuMsg.hWnd, GWL_EXSTYLE, GetWindowLong(MuMsg.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(MuMsg.hWnd, 0, bytOpacity, LWA_ALPHA)
    SetWindowPos MuMsg.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
End Function

Public Function AbrirLauncher()
    Dim ShellInfo As Integer
    ShellInfo = ShellExecute(hWnd, vbNullString, "Launcher.exe", vbNullString, vbNullString, 1)
    
    Select Case ShellInfo
    Case 2
        MsgBox "Falha ao executar o Launcher.exe!", vbCritical, "MuMessenger!"
    End Select
End Function

Public Function Navegar(URL As String)
    On Error Resume Next
    MuMsg.m_WebControl.object.Navigate URL
End Function

Public Function VerificarSource()
    If MuMsg.Deslizar.Enabled = True Then Exit Function
    
    Dim TmpSrc() As String
    On Error GoTo ErroA
    TmpSrc = Split(Source, "#")
    
    If CInt(TmpSrc(2)) = UltimaMsg Then Exit Function
    UltimaMsg = CInt(TmpSrc(2))
    
    Call Mostrar_Msg(TmpSrc(0), TmpSrc(1))
    Exit Function
ErroA:
End Function

Public Function Mostrar_Msg(Titulo As String, Msg As String)
    
    WinMode = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "WindowMode")
    If WinMode = 0 Then Exit Function
    
    MuMsg.Titulo.Caption = Titulo
    MuMsg.Msg.Caption = Msg
    MuMsg.Visible = True
    
    LeftInicial = Screen.Width - (MuMsg.Width + 200)
    MuMsg.Left = Screen.Width - (MuMsg.Width + 200)
    MuMsg.Top = Screen.Height - ((MuMsg.Height + 800))
    
    Call Tip
    
    MuMsg.Deslizar.Interval = Delay_Desaparecer
    MuMsg.Deslizar.Enabled = True
    
End Function
