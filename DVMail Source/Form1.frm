VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DVMail - MarxD Software"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer DelayTimer 
      Interval        =   1000
      Left            =   2040
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1560
      Top             =   5040
   End
   Begin VB.TextBox Tmp 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Corpo_txt 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Assunto_txt 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar E-mail's"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label info2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False


Ativado = 1
Call SendEmail
End Sub





Private Sub DelayTimer_Timer()
If Ativado = 1 Then Delay = Delay + 1
End Sub

Private Sub Form_Load()
CarregaTudo
Ultimo = ReadINI(App.Path & "/config.ini", "DVMail", "Ultimo")
End Sub

Public Function CarregaTudo()
Info.Caption = "Conectado com " & ReadINI(App.Path & "/Config.ini", "DVMail", "Email")
Open App.Path & "\Corpo.txt" For Input As #1
   Corpo = Input(FileLen(App.Path & "\Corpo.txt"), #1)
   Corpo_txt.Text = Corpo
   Assunto = ReadINI(App.Path & "/Config.ini", "DVMail", "Assunto")
   Assunto_txt.Text = Assunto
Close #1
Open App.Path & "\Emails.txt" For Input As #2
    Mails = Split(Input(FileLen(App.Path & "\Emails.txt"), #2), vbNewLine)
Close #2

End Function

Public Function SendEmail()
        Dim Msg As CDO.Message
        Dim Cof As CDO.Configuration
        Dim Camp
        Set Msg = New CDO.Message
        Set Cof = New CDO.Configuration
        Set Camp = Cof.Fields

        With Camp
    
             .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
             .Item(cdoSMTPServer) = "mail.munovus.net"  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
             .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
             .Item(cdoSMTPAuthenticate) = 1
             .Item(cdoSendUserName) = "divulgacao@munovus.net" ' informe o usuario de autenticação
             .Item(cdoSendPassword) = "xaubet95"  'Informe a Senha aqui
             .Update
        End With

If Ultimo >= UBound(Mails) Then
    Ultimo = 0
End If

While Verifica(Mails(Ultimo)) = True
    Ultimo = Ultimo + 1
Wend

With Msg
      Set .Configuration = Cof
          .To = Mails(Ultimo)
          .From = "divulgacao@Munovus.net"
          .Subject = Assunto & " - " & i
          .HTMLBody = Corpo & vbNewLine & "<span style='color:#FFFFFF'>" & i & "</span>"
          .Send
End With

info2.Caption = Mails(Ultimo)
Ultimo = Ultimo + 1
Call WriteINI(App.Path & "/config.ini", "DVMail", "ultimo", "" & Ultimo)

End Function

Public Function Verifica(Mail As String) As Boolean
    If InStr(Mail, "@") = 0 Then
        Verifica = True
        Exit Function
    End If
    
    If InStr(Mail, ",") > 0 Then
        Verifica = True
        Exit Function
    End If
    
    If (Mail = vbNullString Or Len(Mail) < 10) Then
        Verifica = True
    Else
        Verifica = False
    End If
End Function


Private Sub Form_Unload(cancel As Integer)
On Error Resume Next
Shell "tskill DVMAIL", vbHide
End
End Sub



Private Sub Timer1_Timer()
Dim Intervalo As Integer

Intervalo = 60 * ReadINI(App.Path & "/config.ini", "DVMail", "Intervalo")

If Delay >= Intervalo Then
    Call SendEmail
    Delay = 0
Else
    Info.Caption = "Aguardando " & ReadINI(App.Path & "/config.ini", "DVMail", "Intervalo") & " minuto(s)"
End If
End Sub
