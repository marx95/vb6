VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simulador de Conexão do MarxD"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Conectar BOT"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtPorta2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "55901"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtPorta 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "44405"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "mxsv.sytes.net"
      Top             =   360
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   120
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bot por MarxD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Porta GS"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Porta CS"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Logar()

Dim merda(67) As Byte
merda(0) = &HC3
merda(1) = &H44
merda(2) = &H2F
merda(3) = &H29
merda(4) = &H3
merda(5) = &H97
merda(6) = &H1F
merda(7) = &HBB
merda(8) = &HE2
merda(9) = &H4C
merda(10) = &H9C
merda(11) = &H82
merda(12) = &HB7
merda(13) = &HA3
merda(14) = &HA3
merda(15) = &H65
merda(16) = &HEB
merda(17) = &H5
merda(18) = &HAD
merda(19) = &H33
merda(20) = &HFD
merda(21) = &H68
merda(22) = &HC4
merda(23) = &HF1
merda(24) = &HD2
merda(25) = &H23
merda(26) = &HC
merda(27) = &H78
merda(28) = &H46
merda(29) = &H96
merda(30) = &H21
merda(31) = &HEC
merda(32) = &H58
merda(33) = &H51
merda(34) = &H64
merda(35) = &H31
merda(36) = &H58
merda(37) = &H13
merda(38) = &HD1
merda(39) = &H1B
merda(40) = &H60
merda(41) = &H21
merda(42) = &HE4
merda(43) = &H79
merda(44) = &H3B
merda(45) = &HE
merda(46) = &HE9
merda(47) = &HDE
merda(48) = &H4E
merda(49) = &H75
merda(50) = &H11
merda(51) = &H86
merda(52) = &H82
merda(53) = &HF9
merda(54) = &HA8
merda(55) = &HED
merda(56) = &HD8
merda(57) = &HB4
merda(58) = &H5E
merda(59) = &H57
merda(60) = &HEC
merda(61) = &H7
merda(62) = &H5C
merda(63) = &H90
merda(64) = &HAD
merda(65) = &H8C
merda(66) = &HFB
merda(67) = &HCE
Sock(1).SendData merda
End Sub



Private Sub Command3_Click()
Conectado = 0
Selecionado = 0
On Error Resume Next
Sock(0).Close
On Error Resume Next
Sock(1).Close
Sock(0).Connect txtIP.Text, txtPorta.Text
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim Pacote() As Byte
 Dim PacoteStr As String
  Sock(Index).GetData Pacote
  Sock(Index).GetData PacoteStr
  
  If Conectado = 0 Then
  Dim CS(5) As Byte
CS(0) = &HC1
CS(1) = &H6
CS(2) = &HA9
CS(3) = &H20
CS(4) = &H9C
CS(5) = &H2F
Sock(0).SendData CS
Conectado = 1
End If

    If Pacote(0) = 193 And Pacote(1) = 4 And Pacote(2) = 0 And Pacote(3) = 1 Then
    Dim Resp(3) As Byte
    Resp(0) = &HC1
    Resp(1) = &H4
    Resp(2) = &HF4
    Resp(3) = &H6
    Sock(0).SendData Resp
    
  End If
  
  If Pacote(0) = 193 And Pacote(1) = 4 And Pacote(2) = 244 And Pacote(3) = 6 Then
  Text1.Text = Text1.Text + vbNewLine + "primeira parte"
  End If
  
   If Pacote(0) = 193 And Pacote(1) = 22 And Pacote(2) = 244 And Pacote(3) = 3 Then
Text1.Text = Text1.Text + vbNewLine + "Recebido ip de conexao"
  End If
  
   If Pacote(0) = 193 And Pacote(1) = 12 And Pacote(2) = 241 And Pacote(3) = 0 Then
  Text1.Text = Text1.Text + vbNewLine + "Serial do main recebido"
  End If
  
  
   If Pacote(0) = 193 And Pacote(1) = 5 And Pacote(2) = 241 And Pacote(3) = 1 Then
  Dim RespB(23) As Byte
RespB(0) = &HC3
RespB(1) = &H18
RespB(2) = &HC5
RespB(3) = &H8B
RespB(4) = &H23
RespB(5) = &HF2
RespB(6) = &HEE
RespB(7) = &H9F
RespB(8) = &H82
RespB(9) = &H36
RespB(10) = &H3D
RespB(11) = &H4E
RespB(12) = &H7B
RespB(13) = &H5C
RespB(14) = &H2F
RespB(15) = &H35
RespB(16) = &H33
RespB(17) = &H41
RespB(18) = &H8D
RespB(19) = &HF1
RespB(20) = &HF1
RespB(21) = &HF9
RespB(22) = &HE0
RespB(23) = &HDE
Sock(1).SendData RespB

Dim RespC(3) As Byte
RespC(0) = &HC1
RespC(1) = &H4
RespC(2) = &HF3
RespC(3) = &H7A
Sock(1).SendData RespC
  End If
  
   If Pacote(0) = 193 And Pacote(1) = 104 And Pacote(2) = 13 And Pacote(3) = 1 Then
   
  Text1.Text = Text1.Text + vbNewLine + "Selecionando o char"
  Sock(0).Close
  Dim RespD(23) As Byte
RespD(0) = &HC3
RespD(1) = &H18
RespD(2) = &H39
RespD(3) = &HC1
RespD(4) = &H4B
RespD(5) = &H99
RespD(6) = &H8A
RespD(7) = &HE1
RespD(8) = &H13
RespD(9) = &H6A
RespD(10) = &H70
RespD(11) = &H45
RespD(12) = &H70
RespD(13) = &H95
RespD(14) = &HD0
RespD(15) = &H47
RespD(16) = &H61
RespD(17) = &H40
RespD(18) = &H47
RespD(19) = &H3
RespD(20) = &H2D
RespD(21) = &HFC
RespD(22) = &HF7
RespD(23) = &HC9
Sock(1).SendData (RespD)

Timer2.Enabled = True
  End If
  
  'PACOTES C2
   If Pacote(0) = 194 And Pacote(1) = 0 And Pacote(2) = 11 And Pacote(3) = 244 Then
  Text1.Text = Text1.Text + vbNewLine + "Recebeu lista de servers"
  Dim RespA(5) As Byte
RespA(0) = &HC1
RespA(1) = &H6
RespA(2) = &HF4
RespA(3) = &H3
RespA(4) = Pacote(7)
RespA(5) = &H0
Sock(0).SendData RespA
Text1.Text = Text1.Text + vbNewLine + "Tentativa de login na sala " & Pacote(7)

Sock(1).Connect txtIP.Text, txtPorta2.Text
Timer1.Enabled = True
  End If
  
  
  'Text1.Text = Text1.Text & vbNewLine & Pacote(0) & " " & Pacote(1) & " " & Pacote(2) & " " & Pacote(3)
End Sub
Private Sub Form_Load()
Load Sock(1)
End Sub

Private Sub Timer1_Timer()
If Sock(1).State = 7 Then
Call Logar
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Selecionado = 0 Then

Dim RespF(13) As Byte
RespF(0) = &HC1
RespF(1) = &HE
RespF(2) = &HF3
RespF(3) = &H79
RespF(4) = &H88
RespF(5) = &H5F
RespF(6) = &HB2
RespF(7) = &HA5
RespF(8) = &HE7
RespF(9) = &H4F
RespF(10) = &HB1
RespF(11) = &H7
RespF(12) = &H4E
RespF(13) = &H13
Sock(1).SendData (RespF)
Selecionado = 1
End If

Dim RespG(4) As Byte
RespG(0) = &HC1
RespG(1) = &H5
RespG(2) = &H18
RespG(3) = &H90
RespG(4) = &H56
On Error Resume Next
Sock(1).SendData (RespG)
Timer2.Interval = 500

If Sock(1).State <> 7 Then
Label.Caption = "Desconectado"
Conectado = 0
Selecionado = 0
On Error Resume Next
Sock(0).Close
On Error Resume Next
Sock(1).Close
Sock(0).Connect txtIP.Text, txtPorta.Text
Timer1.Enabled = False
Timer2.Enabled = False
Text1.Text = ""
Else
Label.Caption = "Conectado"
Text1.Text = Text1.Text + vbNewLine & "Personagem Atualizado"
End If
End Sub
