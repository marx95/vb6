Attribute VB_Name = "Funçoes"
Global PackPronto() As String
Global Linhas As Integer
Public Function AddLog(Texto As String)

If Linhas >= 400 Then
Linhas = 0
Server.Logs.Text = ""
End If

If Server.Logs.Text = "" Then
Server.Logs.Text = "[" & Now & "] " & Texto
Else
Server.Logs.Text = Server.Logs.Text + vbNewLine + "[" & Now & "] " & Texto
End If
Linhas = Linhas + 1
End Function

Public Function MontarPacote(Tipo As Integer, Tamanho As Integer, Index As Integer, data As String)
For i = 1 To Tamanho
PackPronto(i) = ""
Next i
End Function

Public Function StringToHex(ByVal StrToHex As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim i         As Long
    For i = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        If strReturn = "" Then
        strReturn = strReturn & strTemp
        Else
        strReturn = strReturn & Space$(1) & strTemp
        End If
    Next i
    StringToHex = strReturn
End Function

Public Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim i         As Long
    For i = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, i, 2)))
        strReturn = strReturn & strTemp
    Next i
    HexToString = strReturn
End Function

Public Function Ligar()
'server.Sock(0).LocalPort = ReadINI(App.Path & "/Data/Config.ini", "Config", "GameServer_Porta")
Server.Sock(0).LocalPort = 55901
On Error GoTo Erro_Ligar1
Server.Sock(0).Listen
AddLog ("Servidor Online - " & Server.Sock(0).LocalIP & ":" & Server.Sock(0).LocalPort)
Call AtualizaTitulo
Server.Atualizador.Enabled = True

Exit Function
Erro_Ligar1: AddLog ("Não foi possivel iniciar o servidor na porta " & Server.Sock(0).LocalIP)
End Function

Public Function CalculaVersao()
VersaoOrig = ReadINI(App.Path & "/Data/Config.ini", "Config", "Versao_do_Main")
VersaoFinal = Replace(VersaoOrig, ".", "")
AddLog ("Versão: " & VersaoOrig & " -> " & VersaoFinal)
End Function

Public Function PegaSerial()
Serial = ReadINI(App.Path & "/Data/Config.ini", "Config", "Serial_do_Main")
AddLog ("Serial do Server: " & Serial)
End Function

Public Function LimiteDeUsuarios()
UsuariosMaximos = ReadINI(App.Path & "/Data/Config.ini", "Config", "Jogadores_Maximos")
For i = 1 To UsuariosMaximos
Load Server.Sock(i)
Next i
AddLog ("Usuarios Máximos: " & UsuariosMaximos)
End Function

Public Function AtualizaTitulo()
Dim TempUsers As Integer
For i = 1 To UsuariosMaximos
If Server.Sock(i).State = sckConnected Then
TempUsers = TempUsers + 1
End If
UsuariosConectados = TempUsers
Next i
Server.Caption = "MuOnline Emulador - Usuarios: [" & UsuariosConectados & "/" & UsuariosMaximos & "]"
End Function

Public Function IniciaMssql(ID As Integer)
Mssql_Ip = ReadINI(App.Path & "/Data/Config.ini", "Config", "Mssql_IP")
Mssql_Usuario = ReadINI(App.Path & "/Data/Config.ini", "Config", "Mssql_Usuario")
Mssql_Senha = ReadINI(App.Path & "/Data/Config.ini", "Config", "Mssql_Senha")
Mssql_Db = ReadINI(App.Path & "/Data/Config.ini", "Config", "Mssql_DB")

StringDeConexao = "Provider=SQLOLEDB.1;Password =" & Mssql_Senha & ";Persist Security Info=False;User ID=" & Mssql_Usuario & ";Initial Catalog=" & Mssql_Db & ";Data Source=" & Mssql_Ip

On Error Resume Next
Mssql_Rst(ID).Close

On Error Resume Next
MSSQL_Cnn(ID).Close

On erro GoTo Erro_Mssql_1
MSSQL_Cnn(ID).Open StringDeConexao

If ID = 0 Then
AddLog ("Conectado no MSSQL (" & Mssql_Ip & ":1433)")
End If

Exit Function
Erro_Mssql_1:
AddLog ("Falha ao conectar no MSSQL (" & Mssql_Ip & ":1433]")
End Function

Public Function VerificaLogin(ID As Integer, Resultado As String)
Resultado = Replace(Resultado, HexToString(&H0), "")
Dim Resultados() As String
Dim Login As String
Dim Senha As String
Dim SerialRecebido As String
Dim VersaoRecebida As String

Resultados = Split(Resultado, "=")
Login = Resultados(0)
Senha = Resultados(1)
VersaoRecebida = Resultados(2)
SerialRecebido = Resultados(3)
TentativaLogin(ID) = TentativaLogin(ID) + 1

If TentativaLogin(ID) = 4 Then
Call PacoteContaLogin(ID, 7)
Server.Sock(ID).CloseSck
Exit Function
End If

If VersaoFinal <> VersaoRecebida Then
AddLog ("Cliente(" & ID & ") disconectado por Versão diferente [" & VersaoRecebida & "]")
Server.Sock(ID).CloseSck
Exit Function
End If

If Serial <> SerialRecebido Then
AddLog ("Cliente(" & ID & ") disconectado por Serial diferente [" & SerialRecebido & "]")
Call PacoteContaLogin(ID, 7)
Server.Sock(ID).CloseSck
Exit Function
End If

Dim TotalOn As Integer
For i = 1 To UsuariosMaximos
If Server.Sock(i).State = sckConnected Then
TotalOn = TotalOn + 1
End If
Next i

If TotalOn >= UsuariosMaximos Then
AddLog ("Cliente(" & ID & ") disconectado - Servidor cheio!")
Call PacoteContaLogin(ID, 4)
Server.Sock(ID).CloseSck
Exit Function
End If

On Error Resume Next
Mssql_Rst(ID).Close

On Error GoTo VerificaLogin_AccInexistente
Mssql_Rst(ID).Open "SELECT Login, Senha, Bloqueado FROM Contas WHERE Login='" & Login & "'", StringDeConexao, adOpenKeyset, adLockOptimistic, adCmdText

If Mssql_Rst(ID).Fields(1) <> Senha Then
AddLog ("Cliente(" & ID & ") disconectado - Senha incorreta - Login: " & Login)
Call PacoteContaLogin(ID, 0)
Exit Function
End If

If Mssql_Rst(ID).Fields(2) = 1 Then
AddLog ("Cliente(" & ID & ") disconectado - Conta bloqueada - Login: " & Login)
Call PacoteContaLogin(ID, 5)
Server.Sock(ID).CloseSck
Exit Function
End If

Call PacoteContaLogin(ID, 1)
AddLog ("Cliente(" & ID & ") logado - Login: " & Login)
Personagem(ID).Login = Login

Exit Function
VerificaLogin_AccInexistente:
AddLog ("Cliente(" & ID & ") disconectado - Login: " & Login & " inexistente!")
Call PacoteContaLogin(ID, 2)
End Function

Public Function ToHex(ByRef pstrMessage As String) As String
    
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrHex As String
    
    llngMaxIndex = LenB(pstrMessage)
    
    For llngIndex = 1 To llngMaxIndex
        lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(pstrMessage, llngIndex, 1))), 2)
    Next
    
    ToHex = lstrHex
    
End Function

Public Function PegaItemDoInventario(ByRef pstrMessage As String, Idx As Integer) As Byte
  Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrHex As String
    Dim ToHex As String
    
    llngMaxIndex = LenB(pstrMessage)
    
    For llngIndex = 1 To llngMaxIndex
        lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(pstrMessage, llngIndex, 1))), 2)
    Next
    
    ToHex = lstrHex
    
Dim TempLen As Integer
TempLen = Len(ToHex)
Idx = Idx + 1
PegaItemDoInventario = Val("&H" & Mid(ToHex, Idx * 2 - 1, 2))

End Function

Public Function Enviar(TempPacote As String, TamanhoPack As Integer)

End Function
