Attribute VB_Name = "Global"
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Global PreLoad As Long
Global MsgDV(0 To 16) As String
Global MsgID As Long
Global MsgMax As Long
Global Xat(0 To 64) As String
Global TotalSessoes As Long
Global TotalMsgs As Long
Global Aviso As Long

Public Function Setar_MsgDV()
    If MsgID > MsgMax Then MsgID = 0
    
    Dim Sorteado As Long
    Sorteado = Rnd(MsgMax + 1)
    While Sorteado = MsgID
        Sorteado = Rnd(MsgMax + 1)
    Wend
    
    Dim MsgSaida As String
    Dim Fim(0 To 4) As String
    Fim(0) = "BORA"
    Fim(1) = "BORAA"
    Fim(2) = "GOGO"
    Fim(3) = "GOGOGO"
    Fim(4) = "VAMOS LA"
    
    MsgSaida = Replace(MsgDV(Sorteado), "MuNovus", vbNullString)
    MsgSaida = MsgDV(MsgID) & MsgSaida & " - " & Fim(Rnd(4))
    'MsgSaida = "[" & Replace(MsgSaida, " - ", "] [") & "]"

    Clipboard.Clear
    Clipboard.SetText MsgSaida
    MsgID = MsgID + 1
End Function

Public Function Enviar_Msg()
    Call Setar_MsgDV
    
    Call Mouse(180, 306)
    Call Clicar_Esquerdo(2)
    Call Sleep(200)
    Call Control_V
    Call Mouse(342, 306)
    Call Clicar_Esquerdo(2)
    Call Mouse(((Macro.Divulgar.Left / 15) + (Macro.Divulgar.Width / 15) / 2), ((Macro.Divulgar.Top / 15) + (Macro.Divulgar.Height / 15) / 2))
End Function

Public Function Navegar(URL As String)
    On Error Resume Next
    Macro.m_WebControl.object.Navigate URL
End Function
