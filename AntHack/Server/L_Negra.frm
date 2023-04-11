VERSION 5.00
Begin VB.Form L_Negra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista Negra"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   Icon            =   "L_Negra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MarxD.jcbutton jcbutton2 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
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
      Caption         =   "Desbloquear"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MarxD.jcbutton jcbutton1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
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
      Caption         =   "Bloquear"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o IP"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "L_Negra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Text1.SetFocus
End Sub

Private Sub jcbutton1_Click()
    If Len(Text1.Text) < 7 Then Exit Sub
    On Error Resume Next
    Mssql_Rst(1).Close
    Mssql_Rst(1).Open "SELECT ip FROM Lista_Negra WHERE ip='" & IPProcurar & "'", StringDeConexao2, adOpenKeyset, adLockOptimistic, adCmdText
    Mssql_Rst(1).AddNew
    Mssql_Rst(1).Fields("ip").Value = Text1.Text
    Mssql_Rst(1).Update
    Mssql_Rst(1).Close
    
    Text1.Text = vbNullString
    MsgBox "IP Bloqueado!", vbInformation, "Sucesso!"
    Text1.SetFocus
End Sub

Private Sub jcbutton2_Click()
    If Len(Text1.Text) < 7 Then Exit Sub
    On Error Resume Next
    Mssql_Rst(1).Close
    Mssql_Rst(1).Open "SELECT ip FROM Lista_Negra WHERE ip='" & IPProcurar & "'", StringDeConexao2, adOpenKeyset, adLockOptimistic, adCmdText
    Mssql_Rst(1).Delete
    Mssql_Rst(1).Close
    Text1.Text = vbNullString
    MsgBox "IP Desbloqueado!", vbInformation, "Sucesso!"
    Text1.SetFocus
End Sub
