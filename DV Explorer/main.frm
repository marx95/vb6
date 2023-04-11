VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Principal 
   Caption         =   "Divulgador Explorer"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox comus 
      Height          =   5910
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Pausar"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Parar"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Responder Topicos"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox resptitle 
      Height          =   195
      Left            =   9720
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Limpar Historico"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7276
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7276
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7276
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Adicionar Comunidades"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Logar"
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox mensagem 
      Height          =   195
      Left            =   10200
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox title 
      Height          =   195
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10440
      Top             =   120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Criar Tópicos"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar Mensagem"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browser/Source"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox source 
      Height          =   5880
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.TextBox ende 
      Height          =   405
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Command1_Click()
If m_WebControl.Visible = True Then
m_WebControl.Visible = False
source.Visible = True
Else
m_WebControl.Visible = True
source.Visible = False
End If
End Sub

Private Sub Command10_Click()
If Command10.Caption = "Pausar" Then
On Error Resume Next
Unload captcha

Pausado = 1
Command10.Caption = "Despausar"
Else
Pausado = 0
Command10.Caption = "Pausar"
End If
End Sub



Private Sub Command2_Click()
Editor.Show
End Sub

Private Sub Command3_Click()
m_WebControl.object.navigate "http://www.orkut.com.br/GLogin?cmd=logout"
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Divulgar" Then
Criar = 1
ID_CMM = 0
'Call CarregaLink
End If
End Sub
Private Sub Command5_Click()
On Error Resume Next
m_WebControl.object.Document.getelementbyid("Email").Value = Login
m_WebControl.object.Document.getelementbyid("Passwd").Value = Senha
m_WebControl.object.Document.getelementbyid("PersistentCookie").Checked = True
m_WebControl.object.Document.getelementbyid("gaia_loginform").submit
End Sub

Private Sub Command6_Click()
AdicionarCMM.Show
End Sub

Private Sub Command7_Click()
LimparHist.Show
End Sub

Private Sub Command8_Click()
ID_CMM = 0
Respondendo = 1
Call RespondeTopics
End Sub

Private Sub Command9_Click()
On Error Resume Next
Unload captcha
Divulgando = 0
Respondendo = 0
Pausado = 0
m_WebControl.object.stop
m_WebControl.object.navigate "http://orkut.com.br"
End Sub

Private Sub comus_Click()
m_WebControl.object.navigate "http://www.orkut.com/Main#Community?cmm=" & comus.Text
End Sub

Private Sub ende_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
m_WebControl.object.navigate ende.Text
On Error Resume Next
Timer1.Enabled = True
End If
End Sub

Private Sub Form_Load()
Iniciado = 0
Criar = 0
Responder = 0

WindowsTitulo = "Divulgador Explorer"
Login = ReadINI(App.Path & "/conf.ini", "conf", "login")
Senha = ReadINI(App.Path & "/conf.ini", "conf", "senha")

If ReadINI(App.Path & "/links.ini", "links", "ultimo") > 0 Then
PrecisaResp = 1
End If
If ReadINI(App.Path & "/links.ini", "links", "ultimo") > 5 Then
Command4.Enabled = False
End If

ID_Post = 1
Call Load_Mensagem
Call PuxaLista
On Error Resume Next
Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Principal)
  'm_WebControl.Move source.Left, source.Top, source.Width, comus.Height
  m_WebControl.Visible = True
  m_WebControl.object.navigate "http://orkut.com.br"
  m_WebControl.object.Silent = True
End Sub
Private Sub form_resize()
On Error Resume Next
comus.Height = Me.Height - 1900
source.Height = Me.Height - 1950
source.Width = Me.Width - 2400
ende.Width = Me.Width - 5775
Command5.Left = Me.Width - 1835
Command3.Left = Me.Width - 1835
m_WebControl.Move source.Left, source.Top, source.Width, comus.Height
End Sub

Private Sub EnviaMSG()
ID_Post = ID_Post + 1
'ID_CMM = ID_CMM + 1 'isto é a id para o abrelinha
On Error Resume Next
m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("subject").Value = title.Text & " " & ID_Post
m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("messageBody").Value = mensagem.Text
m_WebControl.object.SetFocus
m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("messageBody").Focus
SendKeys "{Tab}"
SendKeys "{Tab}"
SendKeys "{Enter}"
End Sub
Private Sub CarregaLink()
Dim Comunidade As String
On Error GoTo Erro
Comunidade = comus.List(ID_CMM)
comus.Text = Comunidade
m_WebControl.object.navigate "http://www.orkut.com.br/Main#CommMsgPost?cmm=" & Comunidade
Erro:
Status.Panels(3).Text = "O DVExplorer terminou de divulgar"
'Call StartDV
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
source.Text = m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.documentelement.innerhtml
Me.Caption = m_WebControl.object.Document.title & " - " & WindowsTitulo
ende.Text = m_WebControl.object.locationURL
Status.Panels(3).Text = "DVExplorer - MarxD Softwares"

'####################################################################################
If Criar = 1 Then
'Status.Index(1).Text = "Criando Tópicos"

End If
If Responder = 1 Then
'Status.Index(2).Text = "Respondendo Tópicos"

End If
End Sub
Public Sub Load_Mensagem()
title.Text = ReadINI(App.Path & "/conf.ini", "conf", "titulo")
resptitle.Text = ReadINI(App.Path & "/conf.ini", "conf", "resptitulo")
Open App.Path & "/msg.txt" For Input As #1
    mensagem.Text = Input(FileLen(App.Path & "/msg.txt"), #1)
Close #1
End Sub
Public Sub PuxaLista()
On Error Resume Next
comus.Clear
For i = 1 To 100
Dim Linha As Integer
Dim Resultado As String
Linha = i
Resultado = AbreLinha(App.Path & "/comunidades.txt", Linha)
If Resultado = "" Then
Else
comus.AddItem Resultado
End If
Next i
End Sub
Private Sub RespondeTopics()
Dim MaxCMM As Integer
MaxCMM = ReadINI(App.Path & "/links.ini", "links", "ultimo")
LinkResp = Replace(ReadINI(App.Path & "/links.ini", (ID_CMM + 1), "link"), "CommMsgs", "CommMsgPost")

If MaxCMM = ID_CMM Then
'Respondendo = 0
ID_CMM = 0
Call RespondeTopics
End If
Dim Temp() As String
Dim temp2 As String
Temp() = Split(LinkResp, "?cmm=")
Temp() = Split(Temp(1), "&tid=")

comus.Text = Temp(0)
m_WebControl.object.navigate LinkResp
End Sub
Private Sub EnviaResposta()
ID_CMM = ID_CMM + 1
Status.Panels(1).Text = "Respondido com sucesso"
m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("messageBody").Focus
SendKeys "{Tab}"
SendKeys "{Tab}"
SendKeys "{Enter}"
End Sub
