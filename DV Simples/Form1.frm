VERSION 5.00
Begin VB.Form Divulgador 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   15405
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton ProximoBT 
      Height          =   375
      Left            =   10920
      TabIndex        =   6
      Top             =   120
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
      Caption         =   "Próximo"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0000
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton AnteriorBT 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   120
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
      Caption         =   "Anterior"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":031A
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton SalvarBT 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      Enabled         =   0   'False
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
      Caption         =   "Salvar Link"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0634
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton LinksBT 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   120
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
      Caption         =   "Gerenciar Links"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":094E
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton SourceBT 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
      Caption         =   "Mostrar Source"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0C68
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox LinkText 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox SourceText 
      Height          =   735
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LinkInfo 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Divulgador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub AnteriorBT_Click()
    If LinkAgora > 1 Then
        LinkAgora = LinkAgora - 1
        Call Navegar(LinksDV(LinkAgora))
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 1200 * 15
    Me.Height = 600 * 15
    
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Divulgador)
    m_WebControl.Move 120, 1200, Me.Width - 500, Me.Height - 1900
    m_WebControl.Visible = True
    'Call Navegar("http://facebook.com")
    
    LinkText.Width = Me.Width - 500
    SourceText.Width = Me.Width - 500
    SourceText.Height = Me.Height - 1900
    
    Call Carregar_Links
End Sub

Private Sub Form_Resize()
    If Me.Width < (800 * 15) Then Me.Width = 800 * 15
    If Me.Height < (600 * 15) Then Me.Height = 600 * 15
    
    On Error Resume Next
    m_WebControl.Move 120, 1200, Me.Width - 500, Me.Height - 1900
    
    LinkText.Width = Me.Width - 500
    SourceText.Width = Me.Width - 500
    SourceText.Height = Me.Height - 1900
End Sub

Private Sub jcbutton1_Click()
'Call Navegar("https://www.facebook.com/MuAwaYServer")
'm_WebControl.object.document.documentelement.getelementsbyname("add_comment_text_text")(0).Value = "Teste"
MsgBox m_WebControl.object.document.getelementsbyname("add_comment_text_text")(1).Value = "Teste"
End Sub

Private Sub LinksBT_Click()
    Links.Show
End Sub

Private Sub LinkText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Navegar(LinkText.Text)
    End If
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    If Info = "DocumentComplete" Then
        LinkNavegador = m_WebControl.object.document.URL
        LinkText.Text = LinkNavegador
        
        SalvarBT.Enabled = True
        
        On Error Resume Next
        Source = m_WebControl.object.document.documentelement.innerhtml
        
        Source = Replace(Source, "<HEAD>", "")
        Source = Replace(Source, "</HEAD>", "")
        Source = Replace(Source, "<BODY>", "")
        Source = Replace(Source, "</BODY>", "")
        Source = Replace(Source, vbNewLine, "")
        
        'If InStr(1, Source, "name=""add_comment_text""") Then
            'MsgBox "Achou"
            'MsgBox m_WebControl.object.document.getelementsbyname("add_comment_text_text").Value = "Teste"
        'End If
        Exit Sub
    End If
End Sub

Private Sub ProximoBT_Click()
    If LinkAgora < TotalLinks Then
        LinkAgora = LinkAgora + 1
        Call Navegar(LinksDV(LinkAgora))
    End If
End Sub

Private Sub SalvarBT_Click()
    Call Adicionar_Link
End Sub

Private Sub SourceBT_Click()
    If SourceBT.Caption = "Mostrar Source" Then
        SourceBT.Caption = "Mostrar Navegador"
        m_WebControl.Visible = False
        SourceText.Text = Source
        SourceText.Visible = True
    Else
        SourceBT.Caption = "Mostrar Source"
        m_WebControl.Visible = True
        SourceText.Visible = False
    End If
End Sub
