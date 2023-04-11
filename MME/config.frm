VERSION 5.00
Begin VB.Form config 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar o MME"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3615
   Icon            =   "config.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conexão"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin Project1.jcbutton Command2 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   4080
         Width           =   1455
         _ExtentX        =   2566
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
      Begin Project1.jcbutton Command1 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Link Path HTTP de confirmações"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   2370
      End
      Begin VB.Label Label5 
         Caption         =   "Link Path HTTP do Avatar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Banco de Dados"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Senha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "IP Host"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    WriteINI App.Path & "/config.ini", "MME", "host", " " & Text1.Text
    WriteINI App.Path & "/config.ini", "MME", "Usuario", " " & Text2.Text
    WriteINI App.Path & "/config.ini", "MME", "Senha", " " & Text3.Text
    WriteINI App.Path & "/config.ini", "MME", "Banco", " " & Text4.Text
    WriteINI App.Path & "/config.ini", "MME", "LinkAvatar", " " & Text5.Text
    WriteINI App.Path & "/config.ini", "MME", "LinkAnexo", " " & Text6.Text
    MainForm.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = ReadINI(App.Path & "/config.ini", "MME", "host")
    Text2.Text = ReadINI(App.Path & "/config.ini", "MME", "Usuario")
    Text3.Text = ReadINI(App.Path & "/config.ini", "MME", "Senha")
    Text4.Text = ReadINI(App.Path & "/config.ini", "MME", "Banco")
    Text5.Text = ReadINI(App.Path & "/config.ini", "MME", "LinkAvatar")
    Text6.Text = ReadINI(App.Path & "/config.ini", "MME", "LinkAnexo")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Show
    Unload Me
End Sub
