VERSION 5.00
Begin VB.Form Links 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Links"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin Project1.jcbutton Fechar 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin Project1.jcbutton Remover 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Caption         =   "Remover"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.ListBox LinksBox 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Carregar_Lista
End Sub

Private Function Carregar_Lista()
    LinksBox.Clear
    
    For i = 1 To TotalLinks
        Dim Link As String
        Link = ReadINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L")
        If Len(Link) > 6 Then
            LinksBox.AddItem Link
        End If
    Next i
End Function
Private Sub Remover_Click()
    Dim LinkRemover As String
    LinkRemover = LinksBox.Text
    
    If Len(LinkRemover) < 6 Then Exit Sub
    
    For i = 1 To TotalLinks
        If ReadINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L") = LinkRemover Then
            Call WriteINI(App.Path & "/Data/LinksDB.ini", CStr(i), "L", vbNullString)
        End If
    Next i
    
    Call Carregar_Links
    Call Carregar_Lista
End Sub
