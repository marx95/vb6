VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[MME]"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton jcbutton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "Atualizar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1200
   End
   Begin VB.TextBox Staff 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Info 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin Project1.jcbutton DepositosBT 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Depósitos"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton ContasBT 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Contas"
      CaptionEffects  =   0
   End
   Begin Project1.jcbutton PersonagensBT 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Personagens"
      CaptionEffects  =   0
   End
   Begin Project1.jcbutton WebshopBT 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Webshop"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton Conectar 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Conectar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton Fechar 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Disconectar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton Configurar 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Configurar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Equipe do Server"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atualizações"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Line Line8 
      X1              =   6720
      X2              =   6720
      Y1              =   3960
      Y2              =   4800
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   6720
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   240
      Y1              =   3960
      Y2              =   4800
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   6720
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   6720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Status_LB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desconectado!"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   6495
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Configurar_Click()
    If Status = 1 Then
        MsgBox "Feche a conexão antes de configurar!", vbCritical, "Aviso"
    ElseIf Status = 0 Then
        config.Show
    End If
End Sub

Private Sub Conectar_Click()
    Label1.Visible = True
    Label2.Visible = True
    info.Visible = True
    Staff.Visible = True
    ContasBT.Visible = True
    PersonagensBT.Visible = True
    WebshopBT.Visible = True
    DepositosBT.Visible = True
    Fechar.Visible = True
    Conectar.Visible = False
    Configurar.Visible = False
    jcbutton1.Visible = True
    
    Host = ReadINI(App.Path & "/config.ini", "MME", "host")
    Usuario = ReadINI(App.Path & "/config.ini", "MME", "Usuario")
    Senha = ReadINI(App.Path & "/config.ini", "MME", "Senha")
    Banco = ReadINI(App.Path & "/config.ini", "MME", "Banco")
    LinkAvatar = ReadINI(App.Path & "/config.ini", "MME", "LinkAvatar")
    AnexoLink = ReadINI(App.Path & "/config.ini", "MME", "LinkAnexo")
    Cnn2 = "Provider=SQLOLEDB.1;Password =" & Senha & ";Persist Security Info=False;User ID=" & Usuario & ";Initial Catalog=" & Banco & ";Data Source=" & Host
    cnn.Open Cnn2
    Status = 1
    Status_LB.Caption = "Conectado no " & Host
    
    Call pegaAtu
    Call pegaStaff
End Sub

Private Sub ContasBT_Click()
    contas.Show
End Sub

Private Sub DepositosBT_Click()
    Dep.Show
End Sub

Private Sub Fechar_Click()
    Label1.Visible = False
    Label2.Visible = False
    info.Visible = False
    Staff.Visible = False
    ContasBT.Visible = False
    PersonagensBT.Visible = False
    WebshopBT.Visible = False
    DepositosBT.Visible = False
    Fechar.Visible = False
    Conectar.Visible = False
    Configurar.Visible = True
    Conectar.Visible = True
    On Error Resume Next
    pPBCounter.Enabled = False
    jcbutton1.Visible = False
    
    PBF1.Value = 0
    Status = 0
    On Error Resume Next
    cnn.Close
    Status_LB.Caption = "Desconectado!"
End Sub

Private Sub Form_Paint()
    Status = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell "tskill MME"
    Unload Me
End Sub

Private Sub jcbutton1_Click()
    Call pegaAtu
End Sub

Private Sub PersonagensBT_Click()
    Shell App.Path & "/personagens.exe lollollol", vbNormalFocus
End Sub

Private Sub WebshopBT_Click()
    webshop.Show
End Sub

Private Function pegaAtu()
    info.Text = ""
    On Error Resume Next
    rst.Close

    rst.CursorLocation = adUseClient
    rst.Open "SELECT TOP 1 login FROM mujb.dbo.confirmacoes where aprovado=0", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

    If rst.RecordCount > 0 Then
        info.Text = rst.RecordCount & " depósito(s) á confirmar!"
    Else
        info.Text = "Nenhuma confirmação de pagamento!"
    End If
End Function

Private Function pegaStaff()
    Staff.Text = ""
    On Error Resume Next
    rst.Close

    rst.CursorLocation = adUseClient
    rst.Open "SELECT name, accountid FROM character where ctlcode=32", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

   For i = 0 To rst.RecordCount
        Staff.Text = rst.Fields(0) & " [Login: " & rst.Fields(1) & "]" + vbNewLine + Staff.Text
        rst.MoveNext
   Next i
End Function
