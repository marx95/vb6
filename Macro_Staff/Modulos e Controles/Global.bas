Attribute VB_Name = "Global"
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global Configurado As Boolean
Global Ligado As Boolean
Global PreLoad As Long
Global MsgDV(0 To 16) As String
Global MsgID As Long
Global MsgMax As Long
Global XatMax As Long
Global TotalMsgs As Long
Global XatEnviar As Long

Public Function Sumir_Xats()
    For i = 0 To XatMax
        DoEvents
        Macro.Shock(i).Left = Macro.Width + 15
    Next i
End Function
Public Function Mostrar_Xat(Index As Integer)
    If Index < 0 Or Index > XatMax Then Exit Function
    
    For i = 0 To XatMax
        DoEvents
        If i <> Index Then
            Macro.Shock(i).Left = Macro.Width + 15
        End If
    Next i
    
    If Macro.Shock(Index).Visible = False Then Macro.Shock(Index).Visible = True
    Macro.Shock(Index).Left = 0
End Function

Public Sub PegarMsgDV()
    If MsgID > MsgMax Then MsgID = 0
    MsgID = MsgID + 1
End Sub

Public Function Enviar_Msg()
    
    Dim UltimoXat As String
    Dim i As Integer
    
    UltimoXat = Macro.Xats.Text
    
    If XatEnviar = 0 Then XatEnviar = 1
    If XatEnviar > XatMax Then
        XatEnviar = 1
        Call PegarMsgDV
        
        While Len(MsgDV(MsgID)) < 4
            Call PegarMsgDV
        Wend
    End If
    
    Call Mostrar_Xat(CInt(XatEnviar))
    Macro.Xats.ListIndex = XatEnviar
        
    While Macro.Shock(XatEnviar).Left > 0
        Call Sleep(5)
    Wend
        
    Call Mouse(170, 375)
    Call Clicar_Esquerdo_Duplo(4)

    For i = 1 To Len(MsgDV(MsgID))
        If Ligado = False Then Exit Function
        On Error Resume Next
        Call SendKeys("{" & Mid(MsgDV(MsgID), i, 1) & "}")
        DoEvents
        Call Sleep((5 + Rnd(11)))
    Next i
    Call Enter
        
    TotalMsgs = TotalMsgs + 1
    Macro.Status.Caption = "Msg's Enviadas: " & TotalMsgs
    XatEnviar = XatEnviar + 1
    
    'Call Mouse(610, 280)
End Function

Public Function Desligar()
    Macro.dvBT.Caption = "Divulgar"
    Macro.Info.Caption = "Esc em Divulgar"
        
    Macro.Xats.Enabled = True
    Macro.Tempo.Enabled = True
    Macro.FecharBT.Enabled = True
    Macro.Timer1.Enabled = False
    Ligado = False
End Function

Public Function Ligar()
    Macro.dvBT.Caption = "Parar"
    Macro.Info.Caption = "Esc para Parar!"
        
    Macro.Xats.Enabled = False
    Macro.Tempo.Enabled = False
    Macro.FecharBT.Enabled = False
    Macro.Timer1.Interval = 1
    Macro.Timer1.Enabled = True
    Ligado = True
End Function
