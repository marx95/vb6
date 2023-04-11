Attribute VB_Name = "Global"
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Global PreLoad As Long
Global MsgDV(0 To 16) As String
Global MsgID As Long
Global MsgMax As Long
Global XatMax As Long
Global TotalSessoes As Long
Global TotalMsgs As Long

Public Function Setar_MsgDV()
    If MsgID > MsgMax Then MsgID = 0
    
    Clipboard.Clear
    Clipboard.SetText MsgDV(MsgID)
    MsgID = MsgID + 1
End Function

Public Function Enviar_Msg()
    Call Setar_MsgDV
    
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For i = 0 To XatMax
        X = (Macro.Shock(i).Left / 15) + 50 + 15
        Y = (Macro.Shock(i).Top / 15) + 135
        Call Mouse(X, Y)
        Call Clicar_Esquerdo_Duplo(4)
        Call Control_V
        
        X = (Macro.Shock(i).Left / 15) + 142
        Y = (Macro.Shock(i).Top / 15) + 135
        Call Mouse(X, Y)
        Call Clicar_Esquerdo_Duplo(4)

        TotalMsgs = TotalMsgs + 1
        Macro.Status.Caption = "Msg's Enviadas: " & TotalMsgs & " - Total de Sessões: " & TotalSessoes
    Next i
    
    Call Focalizar_BTControle
    
    TotalSessoes = TotalSessoes + 1
    Macro.Status.Caption = "Msg's Enviadas: " & TotalMsgs & " - Total de Sessões: " & TotalSessoes
End Function

Public Function Desligar()
    Macro.dvBT.Caption = "Divulgar"
    Macro.Info.Caption = "Clique em Divulgar"
        
    Macro.Tempo.Enabled = True
    Macro.FecharBT.Enabled = True
    Macro.Timer1.Enabled = False
End Function

Public Function Ligar()
    Macro.dvBT.Caption = "Parar"
    Macro.Info.Caption = "Clique para Parar!"
        
    Macro.Tempo.Enabled = False
    Macro.FecharBT.Enabled = False
    Macro.Timer1.Interval = 1
    Macro.Timer1.Enabled = True
End Function

Public Function Organizar_Botoes(Index As Long)
    Macro.dvBT.Top = Macro.Shock(Index).Top + 15 + Macro.Shock(0).Height
    Macro.dvBT.Left = Macro.Shock(Index).Left
    Macro.dvBT.Width = Macro.Shock(0).Width / 2
    Macro.FecharBT.Top = Macro.dvBT.Top
    Macro.FecharBT.Left = Macro.dvBT.Left + 15 + Macro.dvBT.Width
    Macro.FecharBT.Width = Macro.Shock(0).Width / 2
    
    Macro.Info.Top = Macro.dvBT.Top + Macro.dvBT.Height + 15
    Macro.Info.Left = Macro.dvBT.Left
    Macro.Info.Width = Macro.dvBT.Width + Macro.FecharBT.Width + 15
End Function

Public Function Organizar_Text()
    Macro.Label2.Visible = True
    Macro.Label2.Left = Macro.Shock(4).Left + 15
    Macro.Label2.Top = Macro.Shock(4).Top + 45 + Macro.Shock(4).Height
    
    Macro.Tempo.Visible = True
    Macro.Tempo.Left = Macro.Label2.Left + Macro.Label2.Width + 60
    Macro.Tempo.Top = Macro.Shock(4).Top + 15 + Macro.Shock(4).Height
    
    Macro.Label1.Visible = True
    Macro.Label1.Left = Macro.Tempo.Left + Macro.Tempo.Width
    Macro.Label1.Top = Macro.Label2.Top
    
    Macro.Status.Visible = True
    Macro.Status.Left = Macro.Label1.Left + 15 + Macro.Label1.Width
    Macro.Status.Top = Macro.Label2.Top
End Function
