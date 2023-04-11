VERSION 5.00
Begin VB.Form SpotMaker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spot Maker do MuNovus.net - [Carregando...]"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   15135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Anuncio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      MouseIcon       =   "Form1.frx":1CCA
      MousePointer    =   99  'Custom
      ScaleHeight     =   735
      ScaleWidth      =   15135
      TabIndex        =   22
      Top             =   0
      Width           =   15135
   End
   Begin VB.TextBox RadioValor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "4"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   9000
      Top             =   5880
   End
   Begin VB.ListBox Lista 
      Height          =   4935
      Left            =   4680
      TabIndex        =   16
      Top             =   810
      Width           =   10335
   End
   Begin Project1.jcbutton Guardar 
      Height          =   375
      Left            =   13680
      TabIndex        =   15
      Top             =   5880
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
      Caption         =   "Guardar"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":1FD4
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Text            =   "MonsterSetBase.txt"
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "www.MuNovus.net"
      Height          =   5040
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.TextBox MonsterID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin Project1.jcbutton Ver 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   435
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Ver Lista"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":22EE
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton Criar 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   4320
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Criar Spot"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":2608
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.ComboBox mapa 
         Height          =   315
         ItemData        =   "Form1.frx":2922
         Left            =   1920
         List            =   "Form1.frx":299E
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Quantidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox posY 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox posX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Direcao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "1"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox ratio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "0"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vador da Área"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monstro ID"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label7 
         Caption         =   "Quantidade"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Posição Y"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Posição X"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Direção"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tempo Respawn"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Mapa ID"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Info 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Width           =   9015
   End
End
Attribute VB_Name = "SpotMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Anuncio_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://muover.net/", vbNullString, vbNullString, 1)
End Sub

Private Sub Criar_Click()
    Monstro_ID_Selecionado = MonsterID.Text
    Dim Erro As Integer

    If Ver.Caption = "Ver Lista" Or Monstro_ID_Selecionado = "" Then Erro = 1
    If posX.Text = "" Then Erro = 1
    If posY.Text = "" Then Erro = 1
    If ratio.Text = "" Then Erro = 1
    If mapa.Text = "" Then Erro = 1
    If Direcao.Text = "" Then Erro = 1
    If Quantidade.Text = "" Then Erro = 1
    If CInt(RadioValor.Text) = 0 Then
        MsgBox "Valor do Radio é inválido!", vbCritical, "ERRO!"
        Exit Sub
    End If

    If Erro = 1 Then
        MsgBox "Faltou preencher algo!", vbCritical, "Preencha tudo"
        Exit Sub
    End If

    Dim Monstro_Linha As String
    Dim Monstro_Linha_Spot As String
    Dim MapaSlit() As String
    Dim MostrarInfo As Integer
    
    MapaSlit = Split(mapa.Text, " - ")
    
    If CInt(MapaSlit(0)) <> Ultimo_Mapa Then
        Ultimo_Mapa = CInt(MapaSlit(0))
        MostrarInfo = 1
        
        If Total_Monstros > 0 Then
            Lista.AddItem "// - Total de " & Total_Monstros
            Lista.AddItem ""
        End If
        
        Lista.AddItem ("// - Mapa: " & MapaSlit(1))
        Lista.AddItem (Replace("// MobID Mapa Tempo CordX CordY CordX2 CordY2 Dir Quant Desc", " ", vbTab))
        Total_Monstros = 0
    End If
    
    Ultimo_Mapa = CInt(MapaSlit(0))
    
    If CInt(Quantidade.Text) > 1 Then
        Monstro_Linha = " " & CInt(Monstro_ID_Selecionado) & " " & CInt(MapaSlit(0)) & " " & CInt(ratio.Text) & " " & CInt(posX.Text) - CInt(RadioValor.Text) & " " & CInt(posY.Text) - CInt(RadioValor.Text) & " " & CInt(posX.Text) + CInt(RadioValor.Text) & " " & CInt(posY.Text) + CInt(RadioValor.Text) & " " & CInt(Direcao.Text) & " " & CInt(Quantidade.Text) & " "
        Monstro_Linha = Replace(Monstro_Linha, " ", vbTab) & "// - " & Monstro_Nome_Selecionado & " - Espalhado"
        Lista.AddItem Monstro_Linha

        Total_Monstros = Total_Monstros + CInt(Quantidade.Text)
    End If
    
    Monstro_Linha_Spot = " " & CInt(Monstro_ID_Selecionado) & " " & CInt(MapaSlit(0)) & " " & CInt(ratio.Text) & " " & CInt(posX.Text) & " " & CInt(posY.Text) & " " & CInt(posX.Text) & " " & CInt(posY.Text) & " " & CInt(Direcao.Text) & " " & CInt(CInt(Quantidade.Text) / 2) & " "
    Monstro_Linha_Spot = Replace(Monstro_Linha_Spot, " ", vbTab) & "// - " & Monstro_Nome_Selecionado & " - Pontual"
    Lista.AddItem Monstro_Linha_Spot
    
    Salvo = 0
    Total_Monstros = Total_Monstros + CInt(CInt(Quantidade.Text) / 2)
    Info.Caption = "Monstros neste grupo: " & Total_Monstros
End Sub

Private Sub Form_Load()
    Ultimo_Mapa = -1
    Text1.Text = "Monstersetbase_" & Replace(Date, "/", "-") & ".txt"
    Lista.AddItem "// ####################### - Spot's feitos pelo software SpotMaker do MuOver.net - ####################### \\"
    Lista.AddItem "1"
    Frame1.Caption = "MuOver.net"
    Me.Caption = "Spot Maker do MuOver.net - [Carregando...]"
    Call Travar
End Sub

Private Sub Form_Paint()
    If Paint = 1 Then Exit Sub
    DoEvents
    Paint = 1
    
    Info.Caption = "Baixando Lista de Monstros..."
    Call DownloadAFile("Monstros.db", "http://pgcontrol.com.br/muover/Monstros.db", False)
    DoEvents
    
    Call DownloadAFile("anuncio.jpg", "http://pgcontrol.com.br/muover/spotmaker.jpg", False)
    On Error Resume Next
    Anuncio.Picture = LoadPicture(App.Path & "\anuncio.jpg")
    
    Dim i As Integer
    For i = 0 To 512
        Dim Monstros() As String
        Monstros = Split(AbreLinha(App.Path & "/Monstros.db", i), ":")
        Monstro_ID(i) = CInt(Monstros(0))
        Monstro_Nome(i) = Monstros(1)
    Next i
    
    Me.Caption = "Spot Maker do MuOver.net"
    Info.Caption = "Carregado com sucesso!"
    Call Liberar
End Sub

Private Sub Form_Unload(cancel As Integer)
    If Salvo = 1 Then
        On Error Resume Next
        Call Kill("Monstros.db")
        End
    End If
    
    If MsgBox("Tem certeza que quer fechar?", vbYesNo, "Fechar?") = vbNo Then
        cancel = 1
    Else
        On Error Resume Next
        Call Kill("Monstros.db")
        On Error Resume Next
        Call Kill("anuncio.jpg")
        End
    End If
End Sub

Private Sub Guardar_Click()
    Dim ArquivoSaida As String
    Dim Spots As String
    Dim i As Integer
    Lista.AddItem "end"
    
    For i = 0 To Lista.ListCount
        If Len(Spots) = 0 Then
            Spots = Lista.List(i)
        Else
            Spots = Spots + vbNewLine + Lista.List(i)
        End If
    Next i

    ArquivoSaida = App.Path & "\" & Text1.Text
    Open ArquivoSaida For Output As #1
        Print #1, Spots
    Close #1
    
    Salvo = 1
    MsgBox "Salvo com sucesso!", vbInformation, "Salvo!"
End Sub

Private Sub MonsterID_Change()
    If Len(MonsterID.Text) > 0 Then Call PegarNomeMonstro(CInt(MonsterID.Text))
End Sub

Private Sub Timer1_Timer()
    If ListaMonstros.Visible = False Then Call Liberar
    If ListaMonstros.Visible = True Then Call Travar
End Sub

Private Sub Ver_Click()
    ListaMonstros.Show
End Sub
