VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Divulgador - Macro do MuNovus.net"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MacroMuNovus.Progressbar_user PB 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16750899
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Já está aberto", vbExclamation, "Aviso"
        End
    End If
    PB.ShowText = True
    Status.Caption = "Carregando..."
    DoEvents
End Sub

Private Sub Form_Paint()
    If PreloadInicio = 1 Then Exit Sub
    PreloadInicio = 1

    Timer1.Enabled = True
End Sub

Public Function Iniciar_Veri()
    Status.Caption = "Carregando..."
    DoEvents
    
    Dim erro As Integer
    Dim mCrc As New CRC
    erro = 0
    MinhaCRC = mCrc.CRC(App.Path & "/" & App.EXEName & ".exe")
    
    Call Baixar_Arquivos
    DoEvents
    
    Me.Visible = False
    Form1.Show
End Function

Public Function Baixar_Arquivos()
    On Error Resume Next
    Kill "Xats.rar"
    On Error Resume Next
    Kill "Xats.ini"
    On Error Resume Next
    Kill "a1.jpg"
    On Error Resume Next
    Kill "a2.jpg"
    On Error Resume Next
    Kill "a3.jpg"
    
    If mCrc.CRC(App.Path & "\UnRAR.dll") <> "-540628159" Then
        Tamanho = 155136
        DoEvents
        Call DownloadAFile(App.Path & "\UnRAR.dll", "http://munovus.net/arquivos/UnRAR.dll", True)
    End If
    
    Call DownloadAFile(App.Path & "/Xats.rar", "http://munovus.net/macro/Xats.rar", False)
    Call RARExtract("Xats.rar", "")
    On Error Resume Next
    Kill "Xats.rar"
    
    Call DownloadAFile(App.Path & "/a1.jpg", "http://munovus.net/macro/anuncio1.jpg", False)
    Call DownloadAFile(App.Path & "/a2.jpg", "http://munovus.net/macro/anuncio2.jpg", False)
    Call DownloadAFile(App.Path & "/a3.jpg", "http://munovus.net/macro/anuncio3.jpg", False)
    
    If mCrc.CRC(App.Path & "\Flash10c.ocx") <> "-3295212959" Then
        Tamanho = 1771722
        DoEvents
        Call DownloadAFile(App.Path & "\Flash10c.rar", "http://munovus.net/macro/Flash10c.rar", True)
        Call RARExtract("Flash10c.rar", "")
        On Error Resume Next
        Kill "Flash10c.rar"
    End If
    
    If Len(Dir(App.Path & "\Config.ini")) = 0 Then
        Call DownloadAFile(App.Path & "\Config.rar", "http://munovus.net/macro/Config.rar", False)
        Call RARExtract("Config.rar", "")
        On Error Resume Next
        Kill "Config.rar"
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProcess(0)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call Iniciar_Veri
End Sub
