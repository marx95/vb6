Attribute VB_Name = "Pacotes"
Public Function PacoteLogin(ID As Integer)
Call IniciaMssql(ID)
Call AtualizaTitulo

Dim Pacote(11) As Byte
Pacote(0) = &HC1
Pacote(1) = &HC
Pacote(2) = &HF1
Pacote(3) = &H0
Pacote(4) = &H1
Pacote(5) = &H12
Pacote(6) = &HC0
Pacote(7) = Val("&H" & StringToHex(Mid(VersaoFinal, 1, 1)))
Pacote(8) = Val("&H" & StringToHex(Mid(VersaoFinal, 2, 1)))
Pacote(9) = Val("&H" & StringToHex(Mid(VersaoFinal, 3, 1)))
Pacote(10) = Val("&H" & StringToHex(Mid(VersaoFinal, 4, 1)))
Pacote(11) = Val("&H" & StringToHex(Mid(VersaoFinal, 5, 1)))

Personagem(ID).IP = Server.Sock(ID).RemoteHostIP
Personagem(ID).Porta = Server.Sock(ID).RemotePort

Server.Sock(ID).SendData Pacote
End Function

Public Function PacoteContaLogin(ID As Integer, Resultado As Integer)
Dim Pacote(4) As Byte
Pacote(0) = &HC1
Pacote(1) = &H5
Pacote(2) = &HF1
Pacote(3) = &H1

Select Case Resultado

Case 0: 'Senha incorreta
Pacote(4) = &H0

Case 1: 'login com sucesso
Pacote(4) = &H1

Case 2: 'conta inexistente
Pacote(4) = &H2

Case 3: 'conta conectada
Pacote(4) = &H3

Case 4: 'Server Cheio
Pacote(4) = &H4

Case 5: 'conta bloqueada
Pacote(4) = &H5

Case 6: 'nova versao requerida
Pacote(4) = &H6

Case 7: 'erro de conexao
Pacote(4) = &H7
End Select

Server.Sock(ID).SendData Pacote
End Function


Public Function EnviaListadDePersonagens(ID As Integer)

AddLog ("Cliente(" & ID & ") Pedido lista de char - Login: " & Personagem(ID).Login)

On Error Resume Next
Mssql_Rst(ID).Close
'On Error GoTo ContaSemPersonagens
Mssql_Rst(ID).Open "SELECT Nome, pLevel, Classe, Tipo, Inventario FROM Personagens WHERE Login='" & Personagem(ID).Login & "'", StringDeConexao, adOpenKeyset, adLockOptimistic, adCmdText

'########################################################################################################
Dim PacoteEmString As String
Dim temp As String
Dim Pacote() As Byte
Dim TempPack As Byte
Dim TotalDePersonagens As Integer
Dim TamanhoPack As Integer
Dim Chars(1 To 5) As String
Dim Levels(1 To 5) As Integer
Dim Classes(1 To 5) As Integer
Dim Tipos(1 To 5) As Integer
Dim Inventario(1 To 5) As String
Dim Ultimo As Integer
Dim LiberaMG As Integer
Dim TamanhoChar As Integer

'########################################################################################################
While (Mssql_Rst(ID).EOF) = False
TotalDePersonagens = TotalDePersonagens + 1
Chars(TotalDePersonagens) = Mssql_Rst(ID).Fields(0)
Levels(TotalDePersonagens) = Mssql_Rst(ID).Fields(1)
Classes(TotalDePersonagens) = Mssql_Rst(ID).Fields(2)
Tipos(TotalDePersonagens) = Mssql_Rst(ID).Fields(3)
Inventario(TotalDePersonagens) = Mssql_Rst(ID).Fields(4)

Mssql_Rst(ID).MoveNext
Wend

'########################################################################################################
TamanhoPack = 4 + (26 * TotalDePersonagens)
ReDim Pacote(TamanhoPack)

Pacote(0) = &HC1
Pacote(1) = TamanhoPack + 1
Pacote(2) = &HF3
Pacote(3) = &H0
Pacote(4) = TotalDePersonagens

'########################################################################################################
For i = 1 To TotalDePersonagens
TamanhoChar = Len(Chars(i))

    For h = 1 To TamanhoChar
        Select Case i
        Case 1: Pacote(5 + h) = Val("&H" & StringToHex(Mid(Chars(i), h, 1)))
        Case 2: Pacote(31 + h) = Val("&H" & StringToHex(Mid(Chars(i), h, 1)))
        Case 3: Pacote(57 + h) = Val("&H" & StringToHex(Mid(Chars(i), h, 1)))
        Case 4: Pacote(83 + h) = Val("&H" & StringToHex(Mid(Chars(i), h, 1)))
        Case 5: Pacote(109 + h) = Val("&H" & StringToHex(Mid(Chars(i), h, 1)))
        End Select
    Next h

'Inventario Personagem 0
Pacote(5) = &H0
Pacote(17) = Levels(i) 'level
Pacote(19) = Tipos(i) 'ctlcode
Pacote(20) = Classes(i) 'classe
Pacote(21) = PegaItemDoInventario(Inventario(i), 0)
Pacote(22) = PegaItemDoInventario(Inventario(i), 10)
Pacote(23) = Val("&H" & PegaItemDoInventario(Inventario(i), 10)) + Val("&H" & PegaItemDoInventario(Inventario(i), 20)) + Val("&H" & PegaItemDoInventario(Inventario(i), 30)) + Val("&H" & PegaItemDoInventario(Inventario(i), 40))
Pacote(24) = &HF
Pacote(25) = &H3
Pacote(29) = &H0
Pacote(30) = &H0

'inventario personagem 1
Pacote(31) = &H1
Pacote(43) = Levels(i) 'level
Pacote(45) = Tipos(i) 'ctlcode
Pacote(46) = Classes(i) 'classe

'inventario personagem 2
Pacote(57) = &H2
Pacote(69) = Levels(i) 'level
Pacote(71) = Tipos(i) 'ctlcode
Pacote(72) = Classes(i) 'classe

'inventario personagem 3
Pacote(83) = &H3
Pacote(95) = Levels(i) 'level
Pacote(97) = Tipos(i) 'ctlcode
Pacote(98) = Classes(i) 'classe

'inventario personagem 4
Pacote(109) = &H4
Pacote(121) = Levels(i) 'level
Pacote(123) = Tipos(i) 'ctlcode
Pacote(124) = Classes(i) 'classe
Next i
'########################################################################################################
Server.Sock(ID).SendData Pacote
'Server.Sock(ID).SendData Personagem(100).
Exit Function

'########################################################################################################
ContaSemPersonagens:
Dim NoCharPack(4) As Byte
NoCharPack(0) = &HC1
NoCharPack(1) = &H5
NoCharPack(2) = &HF3
NoCharPack(3) = &H0
NoCharPack(4) = &H0
AddLog "Sem personagens"
Server.Sock(ID).SendData NoCharPack
Exit Function
End Function

Public Function EnviaInfosDoPersonagem(ID As Integer)
AddLog Personagem(ID).Nome
On Error Resume Next
Mssql_Rst(ID).Close

'On Error GoTo EnviaInfosDoPersonagem_Inexistente
Mssql_Rst(ID).Open "SELECT PosX, PosY, Mapa, Experiencia, Pontos, Forca, Agilidade, Vitalidade, Energia, Vida, VidaMaxima, Mana, ManaMaxima, Estamina, EstaminaMaxima, Zen, PKLevel, Tipo FROM Personagens WHERE Nome='" & Personagem(ID).Nome & "' AND Login='" & Personagem(ID).Login & "'", StringDeConexao, adOpenKeyset, adLockOptimistic, adCmdText

'########################################################################################################
Dim Pacote(78) As Byte
Dim PacoteFinal As String

Pacote(0) = &HC3
Pacote(1) = &H4F
Pacote(2) = &HD0
Pacote(3) = &H1
Pacote(4) = &H39

Pacote(5) = &H59
Pacote(6) = &H48
Pacote(7) = &HDB

Pacote(8) = &H1
Pacote(9) = &HA4
Pacote(10) = &H36
Pacote(11) = &H9A

Pacote(12) = &HAF
Pacote(13) = &HF4
Pacote(14) = &H8A
Pacote(15) = &H2A
Pacote(16) = &H6F
Pacote(17) = &H41
Pacote(18) = &H1B
Pacote(19) = &HE0
Pacote(20) = &H80
Pacote(21) = &H6A
Pacote(22) = &HA8
Pacote(23) = &H9D
Pacote(24) = &H2B
Pacote(25) = &HC9
Pacote(26) = &H19
Pacote(27) = &H23
Pacote(28) = &H1F
Pacote(29) = &HC8
Pacote(30) = &H46
Pacote(31) = &HB0
Pacote(32) = &H45
Pacote(33) = &H60
Pacote(34) = &H55
Pacote(35) = &HF9
Pacote(36) = &H79
Pacote(37) = &H34
Pacote(38) = &HD
Pacote(39) = &H3
Pacote(40) = &HFC
Pacote(41) = &H63
Pacote(42) = &HA0
Pacote(43) = &HE6
Pacote(44) = &HF3
Pacote(45) = &HC6
Pacote(46) = &HF7
Pacote(47) = &H5F
Pacote(48) = &H3E
Pacote(49) = &H1F
Pacote(50) = &H4A
Pacote(51) = &HBE
Pacote(52) = &H72
Pacote(53) = &H10
Pacote(54) = &H6D
Pacote(55) = &HE
Pacote(56) = &H3B
Pacote(57) = &H58
Pacote(58) = &HC2
Pacote(59) = &HE
Pacote(60) = &HAD
Pacote(61) = &HC5
Pacote(62) = &H54
Pacote(63) = &H57
Pacote(64) = &HC3
Pacote(65) = &H4D
Pacote(66) = &HEC
Pacote(67) = &HD9
Pacote(68) = &H27
Pacote(69) = &HB2
Pacote(70) = &H1E
Pacote(71) = &HA
Pacote(72) = &H9E
Pacote(73) = &H35
Pacote(74) = &HC2
Pacote(75) = &HD
Pacote(76) = &H2A
Pacote(77) = &H80
Pacote(78) = &HB5

MsgBox StringToHex(Dec(5, Val("&H" & "6003")))

Server.Sock(ID).SendData Pacote

Exit Function
EnviaInfosDoPersonagem_Inexistente:
AddLog ("Cliente(" & ID & ") - Personagem inexistente: " & Personagem(ID).Nome)
End Function

