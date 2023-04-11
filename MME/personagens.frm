VERSION 5.00
Begin VB.Form personagensFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[MME] Personagens"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7695
   Icon            =   "personagens.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Personagem Selecionado"
      Height          =   6495
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   48
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   47
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   46
         Top             =   5880
         Width           =   735
      End
      Begin Project1.jcbutton Command7 
         Height          =   375
         Left            =   2760
         TabIndex        =   40
         Top             =   4800
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Colocar ou Tirar Zen"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton Command8 
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Zerar Status"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton Command6 
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   4320
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Status Full"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "personagens.frx":1CCA
         Left            =   120
         List            =   "personagens.frx":1CDB
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2880
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "personagens.frx":1D0B
         Left            =   120
         List            =   "personagens.frx":1D4E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   3240
         TabIndex        =   45
         Top             =   5880
         Width           =   105
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         TabIndex        =   44
         Top             =   5880
         Width           =   105
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Mapa"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   5880
         Width           =   405
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   4680
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label17 
         Caption         =   "Zen"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3480
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   480
         Y2              =   5280
      End
      Begin VB.Label Login 
         Caption         =   "Login"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "R. Mensal"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "R. Diario"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "R. Semanal"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Resets"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Pontos"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Comando"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Energia"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Vitalidade"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Agilidade"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Força"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jogador / Banido / GameMaster"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Level / Experiencia"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o Char"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox listachars 
         Appearance      =   0  'Flat
         Height          =   6075
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Buscar pelo nome do char"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
   End
   Begin Project1.jcbutton Command2 
      Height          =   375
      Left            =   5520
      TabIndex        =   41
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin Project1.jcbutton Command3 
      Height          =   375
      Left            =   5520
      TabIndex        =   42
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Fechar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
End
Attribute VB_Name = "personagensFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Carrega_Perso_Infos()
    Text2.Text = rst.Fields(0)
    Text3.Text = rst.Fields(1)
    Text16.Text = rst.Fields(15)
    Text17.Text = rst.Fields(16)
    Text18.Text = rst.Fields(17)
    
    Text14.Text = rst.Fields(18)
    Text15.Text = rst.Fields(19)
    Text19.Text = rst.Fields(20)

    'Status, for agi...
    Text4.Text = rst.Fields(4)
    Text5.Text = rst.Fields(5)
    Text6.Text = rst.Fields(6)
    Text7.Text = rst.Fields(7)
    Text8.Text = rst.Fields(8)
    Text9.Text = rst.Fields(9)
    Text10.Text = rst.Fields(10)
    
    'resets
    Text11.Text = rst.Fields(11)
    Text12.Text = rst.Fields(12)
    Text13.Text = rst.Fields(13)
    
    'Jogador ou staff
    Select Case rst.Fields(3)
    Case 0
    Combo2.Text = "Jogador"
    Case 1
    Combo2.Text = "Bloqueado"
    Case 32
    Combo2.Text = "GameMaster"
    Case 34
    Combo2.Text = "Administrador"
    Case Else
    Combo2.Text = "Jogador"
    End Select
    
    'Classes
    Select Case rst.Fields(2)
    Case 0
    Combo1.Text = "Dark Wizard"
    Case 1
    Combo1.Text = "Soul Master"
    Case 2
    Combo1.Text = "Grand Master"
    
    Case 16
    Combo1.Text = "Dark Knight"
    Case 17
    Combo1.Text = "Blade Knight"
    Case 18
    Combo1.Text = "Blade Master"
    
    Case 32
    Combo1.Text = "Elf"
    Case 33
    Combo1.Text = "Muse Elf"
    Case 34
    Combo1.Text = "Hight Elf"
    
    Case 48
    Combo1.Text = "Magic Gladiator"
    Case 49
    Combo1.Text = "Duel Master"
    
    Case 64
    Combo1.Text = "Dark Lord"
    Case 65
    Combo1.Text = "Lord Emperor"
    
    Case 80
    Combo1.Text = "Summoner"
    Case 81
    Combo1.Text = "Blody Summoner"
    Case 82
    Combo1.Text = "Dimension Master"
    End Select
End Sub
Private Sub Salva_Perso_Infos()
    rst.Fields(0) = Text2.Text
    rst.Fields(1) = Text3.Text
    rst.Fields(15) = (Text16.Text)
    rst.Fields(17) = Text18.Text
    
    rst.Fields(18) = Text14.Text
    rst.Fields(19) = Text15.Text
    rst.Fields(20) = Text19.Text
    
    'Status, for agi...
    rst.Fields(4) = Text4.Text
    rst.Fields(5) = Text5.Text
    rst.Fields(6) = Text6.Text
    rst.Fields(7) = Text7.Text
    rst.Fields(8) = Text8.Text
    rst.Fields(9) = Text9.Text
    rst.Fields(10) = Text10.Text
    
    'Resets
    rst.Fields(11) = Text11.Text
    rst.Fields(12) = Text12.Text
    rst.Fields(13) = Text13.Text
    
    'Staff ou player
    Select Case Combo2.Text
    Case "Jogador"
    rst.Fields(3) = 0
    Case "Bloqueado"
    rst.Fields(3) = 1
    Case "GameMaster"
    rst.Fields(3) = 32
    Case "Administrador"
    rst.Fields(3) = 34
    Case Else
    rst.Fields(3) = 0
    End Select
    
    
    'Classes
    Select Case Combo1.Text
    Case "-"
    MsgBox "Selecione uma classe!", vbCritical, "Erro!"
    Exit Sub
    
    Case "Dark Wizard"
    rst.Fields(2) = 0
    Case "Soul Master"
    rst.Fields(2) = 1
    Case "Grand Master"
    rst.Fields(2) = 2
    Case "Dark Knight"
    rst.Fields(2) = 16
    Case "Blade Knight"
    rst.Fields(2) = 17
    Case "Blade Master"
    rst.Fields(2) = 18
    Case "Elf"
    rst.Fields(2) = 32
    Case "Muse Elf"
    rst.Fields(2) = 33
    Case "Hight Elf"
    rst.Fields(2) = 34
    Case "Magic Gladiator"
    rst.Fields(2) = 48
    Case "Duel Master"
    rst.Fields(2) = 49
    Case "Dark Lord"
    rst.Fields(2) = 64
    Case "Lord Emperor"
    rst.Fields(2) = 65
    Case "Summoner"
    rst.Fields(2) = 80
    Case "Blody Summoner"
    rst.Fields(2) = 81
    Case "Dimension Master"
    rst.Fields(2) = 82
    End Select
    
    rst.Update
    Call Carrega_Perso_Infos
    MsgBox "Dados Atualizados!", vbInformation, "Sucesso"
End Sub

Private Sub Command2_Click()
Call Salva_Perso_Infos
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command6_Click()
Dim tmp As String
tmp = ReadINI(App.Path & "/Config.ini", "MME", "PontosMaximos")
Text4.Text = tmp
Text5.Text = tmp
Text6.Text = tmp
Text7.Text = tmp
Text8.Text = tmp
End Sub

Private Sub Command7_Click()
If Text18.Text = 50000000 Then
Text18.Text = 0
Else
Text18.Text = 50000000
End If
End Sub

Private Sub Command8_Click()
Text4.Text = 50
Text5.Text = 50
Text6.Text = 50
Text7.Text = 50
Text8.Text = 50
End Sub


Public Function PegaChars()
    listachars.Clear

    On Error Resume Next
    rst.Close
    rst.CursorLocation = adUseClient
    On Error Resume Next
    rst.Open "SELECT TOP 50 Name FROM Character", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

    For i = 0 To 50
        listachars.AddItem rst.Fields(0)
        rst.MoveNext
    Next i
End Function

Public Function PegaCharsPeloNome()
    listachars.Clear

    On Error Resume Next
    rst.Close
    rst.CursorLocation = adUseClient
    On Error Resume Next
    rst.Open "SELECT Name FROM Character where name='" & Text1.Text & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

    Dim Limite As Integer
    If rst.RecordCount > 50 Then
        Limite = 50
    Else
        Limite = rst.RecordCount
    End If
    
    For i = 0 To Limite
        listachars.AddItem rst.Fields(0)
        rst.MoveNext
    Next i
End Function

Private Sub Form_Load()
    If App.PrevInstance = True Then End
    If Command$ <> "lollollol" Then End
    
    Host = ReadINI(App.Path & "/config.ini", "MME", "host")
    Usuario = ReadINI(App.Path & "/config.ini", "MME", "Usuario")
    Senha = ReadINI(App.Path & "/config.ini", "MME", "Senha")
    Banco = ReadINI(App.Path & "/config.ini", "MME", "Banco")
    LinkAvatar = ReadINI(App.Path & "/config.ini", "MME", "LinkAvatar")
    AnexoLink = ReadINI(App.Path & "/config.ini", "MME", "LinkAnexo")
    Cnn2 = "Provider=SQLOLEDB.1;Password =" & Senha & ";Persist Security Info=False;User ID=" & Usuario & ";Initial Catalog=" & Banco & ";Data Source=" & Host
    cnn.Open Cnn2
    Me.Show
End Sub

Private Sub Form_Paint()
    Call PegaChars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst.Close
End Sub
Public Function BuscaChar(charTmp As String)
On Error Resume Next
rst.Close
rst.CursorLocation = adUseClient
    rst.Open "Select name, clevel, class, CtlCode, Strength, Dexterity, Vitality, Energy, Leadership, LevelUpPoint, resets, resetsd, resetss, resetsm, avatarlink, Experience, AccountID, Money, MapNumber, MapPosX, MapPosY  from Character where Name='" & charTmp & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText
    Call Carrega_Perso_Infos
    Frame2.Visible = True
    Frame3.Visible = True
    Command2.Visible = True
    personagens.Caption = "[MME] Personagens - " & charTmp
End Function

Private Sub listachars_Click()
    Call BuscaChar(listachars.Text)
End Sub

Private Sub Text1_Change()
    Call PegaCharsPeloNome
    Me.Caption = "[MME] Personagens - Pressione Enter"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaChar(Text1.Text)
End Sub
