VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Server 
   Caption         =   "MuOnline Emulador"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   Icon            =   "Emulador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleMode       =   0  'User
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Logs 
      Height          =   6375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   10935
   End
   Begin VB.Timer Atualizador 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2400
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Usuarios Online"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin Project1.Socket Sock 
      Index           =   0
      Left            =   2880
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSAdodcLib.Adodc Mssql 
      Height          =   330
      Index           =   0
      Left            =   3360
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Atualizador_Timer()
Call AtualizaTitulo
End Sub

Private Sub Command1_Click()
Administrar_Jogadores.Show
End Sub

Private Sub Form_load()
Call CalculaVersao
Call PegaSerial
Call LimiteDeUsuarios
Call IniciaMssql(0)
Call Ligar
End Sub

Private Sub Form_Resize()
On Error Resume Next
Logs.Width = Me.Width - 355
On Error Resume Next
Logs.Height = Me.Height - 1235
End Sub

Public Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Integer

For i = 1 To UsuariosMaximos
If Sock(i).State = sckClosed Then
Sock(i).Accept requestID
UsuariosConectados = UsuariosConectados + 1
AddLog ("Cliente(" & i & ") Conectado - " & Sock(i).RemoteHostIP & ":" & Sock(i).RemotePort)

Call PacoteLogin(i)
Exit Sub
End If
Next i

AddLog ("Tentativa de conexão - Excedeu o numero de usuarios máximos")
End Sub

Public Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim PacoteTemporario As String
Dim Tamanho As Integer
Dim Pacote(1024) As String
Dim Resultado As String

Sock(Index).GetData PacoteTemporario
Tamanho = Len(PacoteTemporario)

'ISTO DEFINE O PACOTE, SUPER IMPORTANTE
For i = 0 To Tamanho - 1
Pacote(i) = StringToHex(Mid(PacoteTemporario, i + 1, 1))
Next i

'####################################################################################################################
Select Case Pacote(0)
Case "C1":
    Select Case Pacote(2)
    Case "F3":
            Select Case Pacote(3)
            Case "7A": 'Lista de personagens
                Call EnviaListadDePersonagens(Index)
                Exit Sub
            Case "79": 'Entrar no Mapa (Carrega infos do personagem)
                Personagem(Index).Nome = Dec(4, PacoteTemporario)
                Call EnviaInfosDoPersonagem(Index)
                Exit Sub
            End Select
    End Select
'####################################################################################################################
Case "C2":

'####################################################################################################################
Case "C3":
On Error Resume Next
Resultado = Dec(1, PacoteTemporario)
Resultado = Replace(Resultado, HexToString(&H0), "")

    Select Case Resultado
    Case 241:
    Resultado = Dec(2, PacoteTemporario)
    Resultado = Replace(Resultado, HexToString(&H0), "")

        Select Case Resultado
        Case 1:
            Resultado = Dec(3, PacoteTemporario)
            Call VerificaLogin(Index, Resultado)
            Exit Sub
        End Select
    End Select



Case "C4":
End Select
End Sub
