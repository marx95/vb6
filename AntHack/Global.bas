Attribute VB_Name = "Globals"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Global Hacks(0 To 35)       As String       ' Array da lista de hacks
Global HackID               As Integer      ' ID da Array da lista de hacks
Global Janela               As String       ' Nome da janela do MU(main.exe) para findwindow
Global JogoAberto           As Integer      ' 1 para o jogo aberto

Global MainCRC              As String       ' CRC do Main.exe
Global MainVerificado       As Integer      ' 1 para quando o main.exe ser verificado por crc
Global ListaHacksCarregado  As Integer      ' 1 para lista de hacks recebida carregada
Global mcrc                 As New CRC
Global Login                As String       ' Diz o login do usuario, é pego pelo registro, se existente
Global TempoExcedido        As Integer      ' Variavel que recebe o valor do tempo excedido
Global Tentativas           As Integer      ' - Numero de tentativas
Global Sucesso              As Integer      ' - Sucesso para quando conectar
Global ExecutavelDoMain     As String
Global AutoClick            As Integer

Public Function Reportar_MainModificado()
    Dim TmpInfo As String
    Call PararTudo
    
    TmpInfo = "O AntiHack detectou uma modificação e o jogo foi finalizado!" + vbNewLine + "Foi enviado um relatório de erro ao servidor, será feito uma analize sobre o ocorrido!" + vbNewLine + vbNewLine + "Fazer modificações no cliente pode causar o banimento eterno de sua conta!" + vbNewLine + "Seus dados foram enviados (Login, Modificação detectada e IP)"
    HackEncontradoForm.Caption = "Uma modificação foi detectada!"
    HackEncontradoForm.Info.Text = TmpInfo
    
    While Main.Visible = True
        Main.Visible = False
    Wend
    
    While HackEncontradoForm.Visible = False
        HackEncontradoForm.Show
    Wend
End Function

Public Function PararTudo()
    Main.Loading.Enabled = False
    Main.Timer_Hacks.Enabled = False
End Function

Public Function CarregaHacks()
    Hacks(0) = "bot"
    Hacks(1) = "blaster"
    Hacks(2) = "mupie"
    Hacks(3) = "capote"
    Hacks(4) = "cheat"
    Hacks(5) = "hack"
    Hacks(6) = "speed"
    Hacks(7) = "hex"
    Hacks(8) = "hit"
    Hacks(9) = "Catastrophe"
    Hacks(10) = "Pinnacle"
    Hacks(11) = "Destroyer"
    Hacks(12) = "proxy"
    Hacks(13) = "hasty"
    Hacks(14) = "Lipsum"
    Hacks(15) = "buff"
    Hacks(16) = "Utilidades"
    Hacks(17) = "godlike"
    Hacks(18) = "packet"
    Hacks(19) = "combo"
    Hacks(20) = "process"
    Hacks(21) = "inject"
    Hacks(22) = "wall"
    Hacks(23) = "dope"
    Hacks(24) = "decompiler"
    Hacks(25) = "ant ban"
    Hacks(26) = "doshttp"
    Hacks(27) = "trapaceador"
    Hacks(28) = "saruen"
    Hacks(29) = "ViCtor"
    Hacks(30) = "ocult"
    Hacks(31) = "editor"
    Hacks(32) = "ollydbg"
    Hacks(32) = "cmdexe"
    Hacks(33) = "prompt"
    Hacks(34) = "gerenciador"
    Hacks(35) = "wpe"
End Function

Public Function HackEncontrado()
    Dim TmpInfo As String
    
    TmpInfo = "Um programa Hacker foi encontrado e o jogo foi finalizado!" + vbNewLine + "Foi enviado um relatório de erro ao servidor, será feito uma analize sobre o ocorrido!" + vbNewLine + vbNewLine + "O uso de hacks pode causar o banimento eterno de sua conta!" + vbNewLine + "Seus dados foram enviados (Login, Programa detectado e IP)"
    HackEncontradoForm.Caption = "Um hack foi detectado!"
    HackEncontradoForm.Info.Text = TmpInfo
    
    While HackEncontradoForm.Visible = False
        HackEncontradoForm.Show
    Wend
End Function

Public Function MoverFormSemCaption(Botao As Integer, FormHwnd As Long) As Long
    If Botao = 1 Then
    Call ReleaseCapture
    MoverFormSemCaption = SendMessage(FormHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Function
