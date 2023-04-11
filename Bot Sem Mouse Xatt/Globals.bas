Attribute VB_Name = "Globals"
    Global Usuario As String
    Global Senha As String
    Global Source As String
    Global Grupo As String
    
    Global MsgID As Long
    Global MsgMax As Long
    Global MsgDV(0 To 64) As String
    
Public Function Navegar(URL As String)
    On Error Resume Next
    Bot.m_WebControl.object.navigate URL
End Function

Public Function Entrar_Xat()
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(0).Value = Usuario
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(1).Value = Senha
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(2).Value = Bot.Xats.Text
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(4).Click
End Function

Public Function Enviar_Msg()
    If MsgID > MsgMax Then MsgID = 0
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(0).Value = MsgDV(MsgID)
    On Error Resume Next
    Bot.m_WebControl.object.Document.getElementsByTagName("input")(1).Click
    MsgID = MsgID + 1
End Function

