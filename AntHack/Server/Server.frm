VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Server 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Servidor Anti-Hacker"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleMode       =   0  'User
   ScaleWidth      =   13882.55
   StartUpPosition =   2  'CenterScreen
   Begin MarxD.jcbutton TravarCnn 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Modo Manutenção - Offline"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MarxD.jcbutton infos 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Infos"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MarxD.jcbutton reiniciar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Reiniciar Servidor"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Timer Deslogar 
      Interval        =   1000
      Left            =   5160
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   5640
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LogTXT 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13935
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Deslogar_Timer()
For i = ClienteMinimo To ClienteMaximo
    If Sock(i).State > 0 Then
        If Cliente(i).Tempo = 30 Then
            On Error GoTo Erro1
            Sock(i).Close
            Call AddLog("Tempo Excedido: " & Cliente(i).IP)
        Else
            Cliente(i).Tempo = Cliente(i).Tempo + 1
        End If
    End If
Next i
Exit Sub

Erro1:
    Exit Sub
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Já em execução", vbCritical, "ERRO"
        End
    End If
    
    Call CarregaTudo
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    LogTXT.Width = Me.Width - 400
    On Error Resume Next
    LogTXT.Height = Me.Height - 1050
End Sub

Private Sub infos_Click()
MsgBox "AntiHack Server 1.01" + vbNewLine + "Sistema feito por MarxD" + vbNewLine + "www.munovus.net", vbInformation, "Informações"
End Sub

Private Sub jcbutton1_Click()
    L_Negra.Show
End Sub

Private Sub reiniciar_Click()
    Call DescarregaTudo
    Call CarregaTudo
End Sub

Public Sub Sock_ConnectionRequest(Index As Integer, ByVal RequestID As Long)

For i = ClienteMinimo To ClienteMaximo
    If Sock(i).State = 0 Then
        Sock(i).Accept RequestID
        
        Cliente(i).IP = Sock(i).RemoteHostIP
        Cliente(i).Tempo = 1
        On Error GoTo Erro1
        Sock(i).SendData "%0"
        Call AddLog("Conexão aceita: " & Cliente(i).IP)
        Exit Sub
    End If
Next i
Exit Sub

Erro1:
Call AddLog("Falha ao enviar pacote: " & Cliente(i).IP)
Exit Sub

Erro2:
Exit Sub
End Sub

Public Sub Sock_DataArrival(Index As Integer, ByVal Bytes As Long)
   Dim Pacotes As String
   Sock(Index).GetData Pacotes
    
    Select Case Pacotes
    Case "%0" ' - Desconectar o cliente
        On Error GoTo Erro1
        Sock(Index).Close
        Call AddLog("Conexão fechada: " & Cliente(Index).IP)
        Exit Sub
    
    Case "%1" ' - Pacotes de lista de hacks
        If ModoManutencao = 0 Then
            On Error GoTo Erro1
            Sock(Index).SendData PacoteHacks
            Call AddLog("Enviado lista de Hacks para: " & Cliente(Index).IP)
        Else
            On Error GoTo Erro1
            Sock(Index).SendData "%1"
            Call AddLog("Enviado Status de manutenção para: " & Cliente(Index).IP)
        End If
        Exit Sub
    End Select
    
    Call AddDbHackLog(Pacotes, Cliente(Index).IP, Index)
    Exit Sub
    
Erro1:
    Exit Sub
    
Erro2:
    Call AddLog("Falha ao enviar pacote: " & Cliente(Index).IP)
Exit Sub

End Sub

Private Sub TravarCnn_Click()
If TravarCnn.Caption = "Modo Manutenção - Offline" Then
    ModoManutencao = 1
    TravarCnn.Caption = "Modo Manutenção - Online"
    Call AddLog("Manutenção Online - Conexão travada!")
    Exit Sub
End If

If TravarCnn.Caption = "Modo Manutenção - Online" Then
    ModoManutencao = 0
    TravarCnn.Caption = "Modo Manutenção - Offline"
    Call AddLog("Manutenção Offline - Conexão liberada!")
    Exit Sub
End If
End Sub
