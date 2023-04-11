VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulador de Cliente"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Entrar no Mapa"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pedido Lista de char"
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Logar"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "10.10.10.10"
      RemotePort      =   55901
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Winsock1.Close
On Error Resume Next
Winsock1.Connect
End Sub

Private Sub Command2_Click()
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
merda(28) = &H4C
merda(29) = &HF1
merda(30) = &H73
merda(31) = &H75
merda(32) = &H8C
merda(33) = &H77
merda(34) = &H42
merda(35) = &H78
merda(36) = &H99
merda(37) = &H6A
merda(38) = &HC8
merda(39) = &H5B
merda(40) = &HA8
merda(41) = &H33
merda(42) = &H6
merda(43) = &H39
merda(44) = &HC
merda(45) = &H39
merda(46) = &H37
merda(47) = &H91
merda(48) = &H0
merda(49) = &H7F
merda(50) = &HA9
merda(51) = &HDE
merda(52) = &H11
merda(53) = &HD6
merda(54) = &HCC
merda(55) = &HF2
merda(56) = &HC7
merda(57) = &H91
merda(58) = &H1D
merda(59) = &HD
merda(60) = &H26
merda(61) = &HC8
merda(62) = &H4E
merda(63) = &H65
merda(64) = &H97
merda(65) = &H60
merda(66) = &HDE
merda(67) = &HEB
On Error Resume Next
Winsock1.SendData (merda)
End Sub

Private Sub Command3_Click()
Dim pacote(3) As Byte
pacote(0) = &HC1
pacote(1) = &H4
pacote(2) = &HF3
pacote(3) = &H7A
On Error Resume Next
Winsock1.SendData (pacote)
End Sub

Private Sub Command4_Click()
Dim merda(13) As Byte
merda(0) = &HC1
merda(1) = &HE
merda(2) = &HF3
merda(3) = &H79
merda(4) = &HA4
merda(5) = &H77
merda(6) = &H89
merda(7) = &H9B
merda(8) = &HD9
merda(9) = &H71
merda(10) = &H8F
merda(11) = &H39
merda(12) = &H70
merda(13) = &H2D
On Error Resume Next
Winsock1.SendData (merda)
End Sub
