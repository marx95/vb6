Attribute VB_Name = "Globals"
Global PacoteHacks As String
Global MSSQL_Cnn(0 To 1) As New ADODB.Connection
Global Mssql_Rst(0 To 1) As New ADODB.Recordset
Global StringDeConexao1 As String
Global StringDeConexao2 As String
Global Mssql_Ip As String
Global Mssql_Usuario As String
Global Mssql_Senha As String
Global Mssql_Db As String
Global Mssql_Db2 As String
Global BanirAoDetectar As Integer
Global TotalLinhas As Integer
Global ModoManutencao As Integer
Global ClienteMinimo As Integer
Global ClienteMaximo As Integer
Global Cliente() As Info
Global JaCarregado As Integer

Public Type Info
    IP As String
    Tempo As Integer
End Type

Public Function DescarregaTudo()
Call SalvarLog
Server.LogTXT = ""
PacoteHacks = ""
TotalLinhas = 0

For i = 0 To ClienteMaximo ' - inicia no 0 para desligar o server tb
    On Error Resume Next
    Server.Sock(i).Close
Next i

On Error Resume Next
MSSQL_Cnn(0).Close
On Error Resume Next
MSSQL_Cnn(1).Close
End Function

Public Function CarregaTudo()
BanirAoDetectar = ReadINI(App.Path & "/config.ini", "Config", "Banir_ao_detectar_Hack")
ClienteMinimo = 1
ClienteMaximo = 2500

If JaCarregado = 0 Then
    JaCarregado = 1
    ReDim Preserve Cliente(ClienteMinimo To ClienteMaximo)
    
    For i = ClienteMinimo To ClienteMaximo
        Load Server.Sock(i)
    Next i
End If

Call AddLog(">> [Iniciando servidor AntiHack] <<")

If BanirAoDetectar = 0 Then
    Call AddLog("O usuario será NÃO banido ao detectar Hack!")
Else
    Call AddLog("Atenção: O usuario será banido ao detectar Hack!")
End If


Open App.Path & "\Hacks.txt" For Input As #2
Dim TmpFile() As String
TmpFile = Split(Input(FileLen(App.Path & "\Hacks.txt"), #2), vbNewLine)
For i = 0 To UBound(TmpFile)
    If PacoteHacks = vbNullString Then
        PacoteHacks = TmpFile(i)
    Else
        PacoteHacks = PacoteHacks + "#" + TmpFile(i)
    End If
Next i

Dim TmpCrc As String
TmpCrc = ReadINI(App.Path & "/config.ini", "Config", "Main_Crc")
PacoteHacks = TmpCrc & "@" & PacoteHacks        ' Aqui emenda a CRC do main com os Hacks
    If TmpCrc = "0" Then
        Call AddLog("Atenção: A CRC do Main.exe é 0!")
    Else
        Call AddLog("CRC do main checado com sucesso!")
    End If
Call AddLog("Carregado " & UBound(TmpFile) + 1 & " hacks parciais!")
Close #2

Call IniciaMssql(0)
Call IniciaMssql(1)

Server.Sock(0).LocalPort = 90
Server.Sock(0).Listen
Call AddLog("Servidor Anti-Hack Online")
End Function

Public Function AddLog(Texto As String)
If TotalLinhas > 30 Then
    Call SalvarLog
End If

If Server.LogTXT.Caption = vbNullString Then
    On Error GoTo Erro
    Server.LogTXT.Caption = "[" & Now & "] " & Texto
Else
    On Error GoTo Erro
    Server.LogTXT.Caption = Server.LogTXT.Caption + vbNewLine + "[" & Now & "] " & Texto
End If
TotalLinhas = TotalLinhas + 1

Exit Function
Erro:
Call SalvarLog
Server.LogTXT.Caption = "[" & Now & "] Falha ao Adicionar Log"
End Function

Public Function SalvarLog()
If ModoManutencao = 1 Then ' - NAO DEIXA SALVAR LOG, PARA NAO LOTAR O Logs.txt
    'Exit Function
End If

Dim tmpLog As String
Open App.Path & "\Logs.txt" For Input As #1
tmpLog = Input(FileLen(App.Path & "\Logs.txt"), #1)
tmpLog = tmpLog + vbNewLine + Server.LogTXT.Caption
Server.LogTXT.Caption = ""
Close #1
TotalLinhas = 0

Open App.Path & "\Logs.txt" For Output As #1
Print #1, tmpLog
Close #1
End Function

Public Function IniciaMssql(ID As Integer)
Mssql_Ip = ReadINI(App.Path & "/Config.ini", "Config", "Mssql_IP")
Mssql_Usuario = ReadINI(App.Path & "/Config.ini", "Config", "Mssql_Usuario")
Mssql_Senha = ReadINI(App.Path & "/Config.ini", "Config", "Mssql_Senha")
Mssql_Db = ReadINI(App.Path & "/Config.ini", "Config", "Mssql_DB")
Mssql_Db2 = ReadINI(App.Path & "/Config.ini", "Config", "Mssql_DB2")

If ID = 0 Then
StringDeConexao1 = "Provider=SQLOLEDB.1;Password =" & Mssql_Senha & ";Persist Security Info=False;User ID=" & Mssql_Usuario & ";Initial Catalog=" & Mssql_Db & ";Data Source=" & Mssql_Ip

On Error GoTo Erro_Mssql_DB1
MSSQL_Cnn(ID).Open StringDeConexao1
AddLog ("Conectado no MSSQL (" & Mssql_Ip & ":1433) - " & Mssql_Db)
Else

StringDeConexao2 = "Provider=SQLOLEDB.1;Password =" & Mssql_Senha & ";Persist Security Info=False;User ID=" & Mssql_Usuario & ";Initial Catalog=" & Mssql_Db2 & ";Data Source=" & Mssql_Ip
On Error GoTo Erro_Mssql_DB2
MSSQL_Cnn(ID).Open StringDeConexao2
AddLog ("Conectado no MSSQL (" & Mssql_Ip & ":1433) - " & Mssql_Db2)
End If


Exit Function

Erro_Mssql_DB1:
Call AddLog("Falha ao conectar no MSSQL (" & Mssql_Ip & ":1433) - " & Mssql_Db)
Exit Function

Erro_Mssql_DB2:
Call AddLog("Falha ao conectar no MSSQL (" & Mssql_Ip & ":1433) - " & Mssql_Db2)
Exit Function
End Function

Public Function AddDbHackLog(Pacote As String, IP As String, Index As Integer)
Dim TmpPack() As String
Dim HackCapturado As String
Dim LoginCapturado As String

If InStr(1, Pacote, "@") Then
    TmpPack = Split(Pacote, "@")
    TmpPack(1) = Replace(TmpPack(1), "%2", "")
    
    HackCapturado = TmpPack(0)
    LoginCapturado = TmpPack(1)

    If HackCapturado = vbNullString Then ' - Pois nao possui o hack capturado
        Call AddLog("Pacote Inválido [" & Pacote & "]: " & IP)
        Server.Sock(Index).Close
        Exit Function
    End If
    
    If LoginCapturado <> vbNullString Then
        Call AddLog("Hack Detectado - Capturado [" & HackCapturado & "] Login [" & LoginCapturado & "]: " & IP)
    Else
        Call AddLog("Hack Detectado - Capturado [" & HackCapturado & "]: " & IP)
    End If
Else
    Call AddLog("Pacote Inválido [" & Pacote & "]: " & IP)
    Server.Sock(Index).Close
    Exit Function ' - Pois nao possui o hack capturado
End If

On Error Resume Next
Mssql_Rst(1).Close
Mssql_Rst(1).Open "SELECT * from AntiHack", StringDeConexao2, adOpenKeyset, adLockOptimistic, adCmdText
Mssql_Rst(1).AddNew
Mssql_Rst(1).Fields("hack").Value = HackCapturado
Mssql_Rst(1).Fields("login").Value = LoginCapturado
Mssql_Rst(1).Fields("data").Value = Now
Mssql_Rst(1).Fields("ip").Value = IP
Mssql_Rst(1).Update
Mssql_Rst(1).Close

If BanirAoDetectar = 0 Then
    Exit Function
End If

If LoginCapturado <> vbNullString Then
    On Error Resume Next
    Mssql_Rst(0).Close
    Mssql_Rst(0).Open "SELECT bloc_code from MEMB_INFO WHERE memb___id='" & LoginCapturado & "'", StringDeConexao1, adOpenKeyset, adLockOptimistic, adCmdText
    Mssql_Rst(0).Fields(0) = 1
    Mssql_Rst(0).Update
    Mssql_Rst(0).Close
End If
Exit Function

erro_Add:
    Call AddLog("Pacote Inválido [" & Pacote & "]: " & IP)
    Server.Sock(Index).Close
    Exit Function
End Function
