VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form Macro 
   BorderStyle     =   0  'None
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Tempo 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "Form1.frx":138CA
      Left            =   8280
      List            =   "Form1.frx":138CC
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MacroMuNovus.jcbutton FecharBT 
      Height          =   495
      Left            =   8640
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Fechar"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":138CE
      CaptionEffects  =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9600
      Top             =   5160
   End
   Begin VB.ListBox Xats 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MacroMuNovus.jcbutton dvBT 
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Pronto"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":13BE8
      CaptionEffects  =   0
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   10575
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Shock 
      Height          =   1560
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1755
      _cx             =   3096
      _cy             =   2752
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
   Begin VB.Label Aviso3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - Clique em ""Pronto"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Width           =   3585
   End
   Begin VB.Label Aviso2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Coloque o Link na casinha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   4710
   End
   Begin VB.Label Aviso1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - Troque o Nick"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   2790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo (MS)"
      Height          =   195
      Left            =   8280
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   8280
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MousePos"
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   5400
      Width           =   2535
   End
End
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dvBT_Click()
    If dvBT.Caption = "Pronto" Then
        Info.Caption = "Esc para Divulgar"
        dvBT.Caption = "Divulgar"
        Aviso1.Visible = False
        Aviso2.Visible = False
        Aviso3.Visible = False
        Label2.Visible = True
        Tempo.Visible = True
        Xats.Visible = True
        Call Liberar
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Configurado = False Then Exit Sub
    If KeyAscii = 27 Then Call dvBT_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "x: " & X / 15 & " - y: " & Y / 15
End Sub

Private Sub Form_Load()
    Configurado = False
    Macro.Caption = "Macro Divulgador - Equipe MuOver.net"
    
    While Inicio.Visible = True
        Inicio.Visible = False
        Call Sleep(25)
    Wend
    
    Me.BackColor = &HFFFFFF
    Shock(0).Left = 0
    Shock(0).Top = 0
    Shock(0).Quality = 0
    Shock(0).Width = 8240
    Shock(0).Height = 6240
    Shock(0).Movie = "http://www.xatech.com/web_gear/chat/chat.swf"
    Shock(0).FlashVars = "id=196651562"
    Shock(0).Visible = True
End Sub

Private Function Liberar()
    XatMax = 15
    
    Dim i As Integer
    For i = 3 To 70
        Tempo.AddItem ((i * 100))
    Next i
    Tempo.Text = "5000"
    
    For i = 0 To XatMax
        DoEvents
        Status.Caption = "Carregando xat " & i & "/" & XatMax & "..."
        
        If i > 0 Then
            Load Shock(i)
            Shock(i).Movie = Shock(0).Movie
            Shock(i).WMode = Shock(0).WMode
            Shock(i).Left = Shock(0).Left
            Shock(i).Top = Shock(0).Top
            Shock(i).Width = Shock(0).Width
            Shock(i).Height = Shock(0).Height
            Shock(i).Quality = Shock(0).Quality
        End If
        
        If i = 1 Then Shock(i).FlashVars = "id=192791833&gn=DvSeemLeg"
        If i = 2 Then Shock(i).FlashVars = "id=195658646&gn=DvSemVicio"
        If i = 3 Then Shock(i).FlashVars = "id=137264914&gn=fazendoradio"
        If i = 4 Then Shock(i).FlashVars = "id=88716389&gn=CiadosMuBr2010"
        If i = 5 Then Shock(i).FlashVars = "id=172063259&gn=ThefoxGame"
        If i = 6 Then Shock(i).FlashVars = "id=158331371&gn=PortalVCMuOnline"
        If i = 7 Then Shock(i).FlashVars = "id=186447069&gn=webha111"
        If i = 8 Then Shock(i).FlashVars = "id=203722615&gn=reggaevcm"
        If i = 9 Then Shock(i).FlashVars = "id=158552062&gn=powerdivulgacao"
        If i = 10 Then Shock(i).FlashVars = "id=190576800&gn=CreatWebRadioXD"
        If i = 11 Then Shock(i).FlashVars = "id=99596027&gn=MuEvolution_Oficial"
        If i = 12 Then Shock(i).FlashVars = "id=199746347&gn=malvadadivulgacao"
        If i = 13 Then Shock(i).FlashVars = "id=185701172&gn=dvgold"
        If i = 14 Then Shock(i).FlashVars = "id=153195761&gn=NovaWebCheats"
        If i = 15 Then Shock(i).FlashVars = "id=191065073&gn=MuOnlineCD"
        Shock(i).Visible = False
        Call Sleep(100)
    Next i
    
    Xats.List(0) = "MuOver"
    Xats.List(1) = "MegaDV"
    Xats.List(2) = "Central Mu News"
    Xats.List(3) = "Fazendo WebRadio"
    Xats.List(4) = "CiadosMU"
    Xats.List(5) = "CinadosMU"
    Xats.List(6) = "ViciadosMU"
    Xats.List(7) = "WebHacke"
    Xats.List(8) = "SuperDV"
    Xats.List(9) = "FireDV"
    Xats.List(10) = "CentralMuBR"
    Xats.List(11) = "Criando Portais"
    Xats.List(12) = "EuroDV"
    Xats.List(13) = "DvGold"
    Xats.List(14) = "DvAway"
    Xats.List(15) = "CentralDV"
    
    Call Mostrar_Xat(0)
    On Error Resume Next
    Shock(0).SetFocus
    
    Status.Caption = "Carregado com sucesso!"
    Configurado = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProcess(0)
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = (CInt(Tempo.Text) / XatMax)
    Call Enviar_Msg
End Sub

Private Sub Xats_Click()
    Call Mostrar_Xat(Xats.ListIndex)
End Sub
