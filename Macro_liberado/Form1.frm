VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Divulgador - Macro do MuNovus.net"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12255
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Shock 
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   9975
      _cx             =   17595
      _cy             =   9551
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
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
   Begin VB.ListBox Xats 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2370
      Left            =   10200
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Anuncio 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   3
      Left            =   8280
      MouseIcon       =   "Form1.frx":08CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   1095
      ScaleWidth      =   3855
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Anuncio 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   2
      Left            =   4200
      MouseIcon       =   "Form1.frx":0BD4
      MousePointer    =   99  'Custom
      ScaleHeight     =   1095
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   6960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MacroMuNovus.jcbutton DvBt 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Divulgar"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0EDE
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MacroMuNovus.jcbutton LiberarBt 
      Height          =   315
      Left            =   10200
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Pronto... Já coloquei!"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":11F8
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.PictureBox Anuncio 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   1
      Left            =   120
      MouseIcon       =   "Form1.frx":1512
      MousePointer    =   99  'Custom
      ScaleHeight     =   1095
      ScaleWidth      =   3975
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Timer Browser_timer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3600
      Top             =   720
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":181C
      Left            =   4200
      List            =   "Form1.frx":1835
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer EnviarTimer 
      Interval        =   1
      Left            =   3120
      Top             =   720
   End
   Begin VB.Timer StatusTimer 
      Interval        =   60
      Left            =   2640
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O Macro do MuNovus.net utliza o Flash. Caso não esteja aparecendo os Xats, reinstale o Flash Player mais recente possível!"
      Height          =   1275
      Left            =   10200
      TabIndex        =   13
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção! Coloque o Link na Casinha!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   10905
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   8895
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   11895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tempo do Loop de Enviar Msg's"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coordenadas do Mouse"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   465
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Anuncio_Click(Index As Integer)
    Call ShellExecute(Me.hWnd, vbNullString, "http://munovus.net/macro/anuncio.php?id=" & Index, vbNullString, vbNullString, 1)
End Sub

Private Sub Browser_timer_Timer()
    Browser_timer.Interval = 30000
    Call m_WebControl.object.Navigate("http://munovus.net/macro/macro_core.php?crc=" & MinhaCRC & "&verificado=1")
End Sub

Private Sub DvBt_Click()
    If Form1.DvBt.Caption = "Divulgar" Then
        Ativado = 1
        Form1.DvBt.Caption = "Parar"
        Form1.Status.Caption = "Clique para 'Parar'"
        Exit Sub
    Else
        Ativado = 0
        Form1.DvBt.Caption = "Divulgar"
        Form1.Status.Caption = "Clique em 'Divulgar' para começar!"
    End If
End Sub

Private Sub EnviarTimer_Timer()
    If Me.WindowState = 1 Then Ativado = 0
    If Ativado = 0 Then
        Exit Sub
    Else
        Dim TempoSpt() As String
        TempoSpt = Split(Form1.Combo2.Text)
        
        Dim IntervaloEnv As Long
        IntervaloEnv = (TempoSpt(0) * 1000) / TotalXats
        Form1.EnviarTimer.Interval = CInt(IntervaloEnv)
    End If
     
    
    If ChatID = TotalXats Then
        ChatID = 0
    Else
        ChatID = ChatID + 1
    End If
    
    If ChatID = 1 Then Call PegarTextoDV
    
    Call Mostrar_Xat(ChatID)
    Call SetarPosicaoDoMouse(200, 445)
    Call Clicar_Esquerdo
    Call ControlV
    Call Enter
    Call SetarPosicaoDoMouse(67, 45)
    
    On Error Resume Next
    Form1.DvBt.SetFocus
    Total = Total + 1
    Form1.Caption = "Divulgador - Macro do MuNovus.net - [" & Total & " msg's em enviadas]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProcess(0)
End Sub

Private Sub LiberarBt_Click()
    Call LiberarMacro
End Sub

Private Sub Label1_Click()
    Label1.ForeColor = &H80000012
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    If Info = "DocumentComplete" Then
        On Error Resume Next
        Source = m_WebControl.object.Document.documentelement.innerhtml
        Source = Replace(Source, "<HEAD>", vbNullString)
        Source = Replace(Source, "</HEAD>", vbNullString)
        Source = Replace(Source, "<BODY>", vbNullString)
        Source = Replace(Source, "</BODY>", vbNullString)
        Source = Replace(Source, vbNewLine, vbNullString)
        
        If InStr(1, Source, "#VersaoAntiga#") Then
            Ativado = 0
            MsgBox "Por favor, baixe a nova versão do MacroMuNovus em www.macro.munovus.net", vbInformation, "Nova Atualização!"
            End
        End If
        
        If InStr(1, Source, "#ServerOff#") Then
            MsgBox "MasterServer do Macro está offline!", vbCritical, "Macro Indisponivel!"
            End
        End If
        
        If InStr(1, Source, "#LiberarMacro#") Then
            Shock(0).Visible = True
            LiberarBt.Visible = True
            Exit Sub
        End If
        
        If InStr(1, Source, "#INFO#") Then
            Label4.Caption = Replace(Source, "#INFO#", "")
            Exit Sub
        End If
    End If
End Sub

Private Sub StatusTimer_Timer()
    If Me.WindowState = 1 Then Exit Sub
    If Status.Left >= (12000 + Status.Width) Then
            Status.AutoSize = True
            Status.Left = (0 - Status.Width)
    End If
    
    Status.Left = Status.Left + 45
    Status.AutoSize = True
End Sub

Private Sub Form_Load()
    While Inicio.Visible = True
        Inicio.Visible = False
        Call Sleep(25)
    Wend
    
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Form1)
    m_WebControl.Move -15, -15, 15, 15
    m_WebControl.Visible = True
    
    Call DeleteKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\Explorer\Navigating\.Current")
    Call m_WebControl.object.Navigate("http://munovus.net/macro/macro_core.php?crc=" & MinhaCRC)
    Call SetarAnuncios

    TotalXats = 8
    Shock(0).Movie = "http://www.xatech.com/web_gear/chat/chat.swf"
    Shock(0).FlashVars = "id=196913113"
            
    Form1.Visible = True
End Sub

Private Sub Timer1_Timer()
    If Me.WindowState = 1 Then Exit Sub
    Dim mouse As POINTAPI
    Call GetCursorPos(mouse)
    
    Dim posX As Long
    Dim PosY As Long
    posX = mouse.X - (Me.Left / 15)
    PosY = mouse.Y - (Me.Top / 15)
    If posX < 0 Or posX > (Me.Width / 15) Then
        posX = 0
        PosY = 0
    End If
    
    If PosY < 0 Or PosY > (Me.Height / 15) Then
        posX = 0
        PosY = 0
    End If
       
    If posX = 0 Or PosY = 0 Then
    Label1.Caption = "Mouse Cord's: Fora"
    Else
        Label1.Caption = "Mouse Cord's: " & posX & " " & PosY
    End If
End Sub

Private Sub Xats_Click()
    Call Mostrar_Xat(Xats.ListIndex)
End Sub
