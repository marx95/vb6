VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form Macro 
   BorderStyle     =   0  'None
   Caption         =   "Macro Divulgador - Equipe MuNovus.net"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Tempo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "3"
      Top             =   2355
      Visible         =   0   'False
      Width           =   975
   End
   Begin MacroMuNovus.jcbutton FecharBT 
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      BackColor       =   15199212
      Caption         =   "Fechar"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":138CA
      CaptionEffects  =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9240
      Top             =   2280
   End
   Begin MacroMuNovus.jcbutton dvBT 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      BackColor       =   15199212
      Caption         =   "Pronto"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":13BE4
      CaptionEffects  =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clique em Pronto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   9
      Top             =   840
      Width           =   2145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Troque o Nick e o Link da Casinha!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   8
      Top             =   360
      Width           =   4380
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   3360
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   450
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Shock 
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      _cx             =   2990
      _cy             =   2355
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Transparent"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clique em Pronto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MousePos"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dvBT_Click()
    If dvBT.Caption = "Pronto" Then
        Macro.Label3.Visible = False
        Macro.Label4.Visible = False
        Info.Caption = "Clique em Divulgar"
        dvBT.Caption = "Divulgar"
        Tempo.Enabled = True
        Call Liberar
        Call Organizar_Text
        Exit Sub
    End If
    
    If dvBT.Caption = "Divulgar" Then
        Call Ligar
        Exit Sub
    End If
    
    If dvBT.Caption = "Parar" Then
        Call Desligar
        Exit Sub
    End If
End Sub

Private Sub FecharBT_Click()
    Call ExitProcess(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "x: " & X / 15 & " - y: " & Y / 15
End Sub

Private Sub Form_Load()
    While Inicio.Visible = True
        Inicio.Visible = False
        Call Sleep(25)
    Wend
    
    Me.BackColor = &HF2F2F2
    Me.Width = 805 * 15
    Me.Height = 355 * 15
    Shock(0).Left = 15
    Shock(0).Top = 15
    Shock(0).Quality = 0
    Shock(0).Width = 200 * 15
    Shock(0).Height = 150 * 15
    Shock(0).Movie = "http://www.xatech.com/web_gear/chat/chat.swf"
    Shock(0).FlashVars = "id=46670049&gn=MuOnlineAM"
    Shock(0).Visible = True
    
    Call Organizar_Botoes(0)
    Call Focalizar_BTControle
End Sub

Private Function Liberar()
    MsgDV(0) = "(DMD) MuNovus - 99Z - CRIE ACC E GANHE 7DIAS VIP, 3.000 PONTOS E 1 SET FULL A ESCOLHA(STAR)"
    MsgDV(1) = "(DMD) MuNovus - 99Z NOVINHO GOGOGOGO(STAR)"
    MsgMax = 1
    XatMax = 7
    
    For i = 0 To XatMax
        If i > 0 Then
            Load Shock(i)
            Shock(i).Movie = Shock(0).Movie
            Shock(i).WMode = Shock(0).WMode
            
            
            If i <= 3 Then
                Shock(i).Left = (Shock(i - 1).Left + (Shock(0).Width)) + 15
                Shock(i).Top = Shock(0).Top
            Else
                If i = 4 Then
                    Shock(i).Left = 15
                Else
                    Shock(i).Left = (Shock(i - 5).Left + Shock(i).Width) + 15
                End If
                Shock(i).Top = Shock(0).Top + Shock(0).Height + 15
            End If
            
            Shock(i).Width = Shock(0).Width
            Shock(i).Height = Shock(0).Height
            Shock(i).Quality = Shock(0).Quality
        End If
        
        If i = 1 Then Shock(i).FlashVars = "id=158331371&rl=Brazilian"
        If i = 2 Then Shock(i).FlashVars = "id=158552062&gn=powerdivulgaca"
        If i = 3 Then Shock(i).FlashVars = "id=172063259&gn=ThefoxGame"
        If i = 4 Then Shock(i).FlashVars = "id=190576800&gn=CreatWebRadioXD"
        If i = 5 Then Shock(i).FlashVars = "id=137264914&gn=fazendoradio"
        If i = 6 Then Shock(i).FlashVars = "id=186447069&gn=webha111"
        If i = 7 Then Shock(i).FlashVars = "id=191950113&xc=2336&cn=515921820&gb=i7&gn=LotaFacil"
        Shock(i).Visible = True
    Next i
    
    Call Organizar_Botoes(XatMax)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProcess(0)
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = CInt(Tempo.Text) * 1000
    Call Enviar_Msg
End Sub
