VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Divulgador - MuNovus.net"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12135
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Animacao1 
      Interval        =   150
      Left            =   3600
      Top             =   720
   End
   Begin VB.TextBox Proximo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   120
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":08CA
      Left            =   5280
      List            =   "Form1.frx":08FD
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   16
      Top             =   3240
      Width           =   2415
   End
   Begin MacroMuNovus.isButton isButton2 
      Height          =   300
      Left            =   2280
      TabIndex        =   14
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Icon            =   "Form1.frx":0932
      Style           =   6
      Caption         =   "Ver"
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      BackColor       =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":094E
      Left            =   120
      List            =   "Form1.frx":096D
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Timer EnviarTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   720
   End
   Begin MacroMuNovus.isButton isButton1 
      Height          =   615
      Left            =   10560
      TabIndex        =   12
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Icon            =   "Form1.frx":09E4
      Style           =   6
      Caption         =   "Ajuda"
      iNonThemeStyle  =   0
      BackColor       =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin MacroMuNovus.isButton Command1 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Icon            =   "Form1.frx":0A00
      Style           =   6
      Caption         =   "Divulgar"
      IconAlign       =   0
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      BackColor       =   16777215
      Tooltiptitle    =   "s"
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Timer StatusTimer 
      Interval        =   40
      Left            =   2640
      Top             =   720
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   855
      _cx             =   1508
      _cy             =   1296
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1680
      Top             =   720
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
      _cx             =   1508
      _cy             =   1296
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   855
      _cx             =   1508
      _cy             =   1296
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   495
      _cx             =   873
      _cy             =   873
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   495
      _cx             =   873
      _cy             =   1296
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   495
      _cx             =   873
      _cy             =   1296
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   735
      _cx             =   1296
      _cy             =   1296
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   735
      _cx             =   1296
      _cy             =   1296
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
   Begin MacroMuNovus.isButton isButton3 
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Icon            =   "Form1.frx":0A1C
      Style           =   6
      Caption         =   "Pronto"
      iNonThemeStyle  =   0
      BackColor       =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
      Height          =   735
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   855
      _cx             =   1508
      _cy             =   1296
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
   Begin VB.Label Animacao1LB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " - Não se esqueça de colocar www.munovus.net na casinha!"
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
      Left            =   2520
      TabIndex        =   22
      Top             =   5280
      Width           =   7620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2400
      TabIndex        =   21
      Top             =   4920
      Width           =   1260
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
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   1350
      Width           =   8775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primeiro, digite seu Login:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   2760
      Width           =   3690
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
      Top             =   840
      Width           =   11895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tempo ao Enviar Msg's (Em Millisegundos)"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   165
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coordenadas do Mouse"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   465
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Animacao1_Timer()
    If isButton3.Enabled = True Then
     Animacao1LB.Caption = vbNullString
     Exit Sub
    End If
    If Len(Login) >= 4 Then Animacao1.Enabled = False
    If Len(Animacao1LB.Caption) = 27 Then Animacao1LB.Caption = "."
    Animacao1LB.Caption = Animacao1LB.Caption & "."
End Sub

Private Sub Command1_Click()
    Call PegarTextoDV
    Call EnviarMSG
    Ativado = 1
End Sub

Private Sub EnviarTimer_Timer()
    If Me.WindowState = 1 Then Ativado = 0
    If Ativado = 0 Then
        EnviarTimer.Enabled = False
        Exit Sub
    End If
    
    Call SumirXats
    
    ChatID = ChatID + 1
    If ChatID > TotalXats Then
        ChatID = 0
        Call SetarPosicaoDoMouse(448, 38)
        Call Clicar_Esquerdo
        SWF(0).Visible = True
        EnviarTimer.Enabled = False
        Exit Sub
    End If
    
    SWF(ChatID).Visible = True
    
    Call SetarPosicaoDoMouse(280, 500)
    Call Clicar_Esquerdo
    Call ControlV
    Call SetarPosicaoDoMouse(575, 500)
    Call Clicar_Esquerdo

    Total = Total + 1
End Sub

Private Sub isButton1_Click()
    Form2.Show
End Sub

Private Sub isButton2_Click()
If Combo1.ListIndex = -1 Then Exit Sub
    Call SumirXats
    Form1.SWF(Combo1.ListIndex).Visible = True
End Sub

Private Sub isButton3_Click()
    isButton3.Enabled = False
    If Len(Text3.Text) < 4 Then
        MsgBox "Login inválido!", vbCritical, "Erro!"
        isButton3.Enabled = True
        Exit Sub
    Else
        Call m_WebControl.object.navigate("http://munovus.net/site/atualizador_dv.php?f=2&login=" & Text3.Text)
    End If
End Sub

Private Sub Label1_Click()
    Label1.ForeColor = &H80000012
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    If Info = "DocumentComplete" Then
        On Error Resume Next
        Source = m_WebControl.object.Document.documentelement.innerhtml
        
        Source = Replace(Source, "<HEAD>", "")
        Source = Replace(Source, "</HEAD>", "")
        Source = Replace(Source, "<BODY>", "")
        Source = Replace(Source, "</BODY>", "")
        Source = Replace(Source, vbNewLine, "")

        If InStr(1, Source, "#Premiado#") And Len(Login) >= 4 Then
            Label4.Caption = "Sua conta foi premiada!!!"
            Exit Sub
        End If
        
        If InStr(1, Source, "#VersaoErro#") And Len(Login) = 0 Then
            MsgBox "Por favor, baixe a nova versão do MacroMuNovus em www.munovus.net/dv", vbInformation, "Nova Atualização!"
            End
        End If
        
        If InStr(1, Source, "#LoginInvalido#") And Len(Login) = 0 Then
            MsgBox "Login Inválido!!!", vbCritical, "Login invalido!"
            isButton3.Enabled = True
            Exit Sub
        End If
        
        If InStr(1, Source, "#LoginSucesso#") And Len(Login) = 0 Then
            Call SetarLogin
            Exit Sub
        End If
        
        If InStr(1, Source, "#INFO#") Then
            Label4.Caption = Replace(Source, "#INFO#", "")
            Exit Sub
        End If
    End If
End Sub

Private Sub Proximo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Ativado = 0
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
    If App.EXEName <> "MacroMuNovus" Then
        MsgBox "Macro renomeado!", vbCritical, "ERRO!"
        End
    End If
    
    If App.PrevInstance Then
        MsgBox "Já está aberto", vbExclamation, "Aviso"
        End
    End If
    
    Combo1.Text = "MuNovus"
    Combo2.Text = "500"
    TotalXats = 8
    For i = 0 To TotalXats
        SWF(i).Movie = "http://www.xatech.com/web_gear/chat/chat.swf"
        SWF(i).Width = 11895
        SWF(i).Height = 6255
    Next i
    Call SumirXats
    
    SWF(0).FlashVars = "id=191206233&rl=Brazilian"
    SWF(1).FlashVars = "id=158331371&xc=2336&cn=58078659&gb=i2&gn=PortalVCMuOnline&rl=Brazilian"
    SWF(2).FlashVars = "id=158552062&gn=powerdivulgaca"
    SWF(3).FlashVars = "id=172063259&gn=ThefoxGame"
    SWF(4).FlashVars = "id=190576800&amp;gn=CreatWebRadioXD"
    SWF(5).FlashVars = "id=137264914&gn=fazendoradio"
    SWF(6).FlashVars = "id=186447069&gn=webha111"
    SWF(7).FlashVars = "id=178565843&rl=Brazilian"
    SWF(8).FlashVars = "id=46670049&gn=MuOnlineAM"
    
    MsgDV(1) = "[MuNovus] Ganhe 10Golds por Reset"
    MsgDV(2) = "[MuNovus] Acumulativo - 64K"
    MsgDV(3) = "[MuNovus] SERVIDOR FULLPVP - 10GOLDS POR RESET"
    MsgDV(4) = "[MuNovus] COM VAGAS NA EQUIPE GOGOGOGOOG"
    MsgDV(5) = "[MuNovus] GANHE GOLDS POR RESET"
    MsgDV(6) = "[MuNovus] EVENTO REI DO PVP"
    MsgDV(7) = "[MuNovus] SEASON 3 ORIGINAL SEM FRESCURA!!!"
    MsgDV(8) = "[MuNovuS] - PVP EQUILIBRADO"
    
    MsgDV(10) = "[MuNovuS] - Spots Lotadao -"
    MsgDV(11) = "[MuNovuS] - CASINHA GERAL!!!"
    MsgDV(12) = "[MuNovuS] - Com ótima administração"
    MsgDV(13) = "[MuNovuS] - Todos sabados e domingos Eventos especiais"
    MsgDV(14) = "[MuNovuS] - Haha, Ta esperando oque? CASINHA!!"
    MsgDV(15) = "[MuNovuS] - Venha confirir"
    
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Form1)
    m_WebControl.Move Me.Left, Me.Top, 15, 15
    m_WebControl.Visible = True
    
    Dim MinhaCRC As String
    Dim MCRC As New CRC
    MinhaCRC = MCRC.CRC(App.Path & "/MacroMuNovus.exe")
    Call m_WebControl.object.navigate("http://munovus.net/site/atualizador_dv.php?f=1&versao=" & MinhaCRC)
    
    On Error Resume Next
    Text3.Text = GetSettingString(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ID")
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    If CInt(Text2.Text) < 30 Then Text2.Text = "150"
    
    On Error Resume Next
    If CInt(Text2.Text) > 1000 Then Text2.Text = "500"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And isButton3.Enabled = True Then
        isButton3.Enabled = False
        If Len(Text3.Text) < 4 Then
            MsgBox "Login inválido!", vbCritical, "Erro!"
            isButton3.Enabled = True
            Exit Sub
        Else
            Call m_WebControl.object.navigate("http://munovus.net/site/atualizador_dv.php?f=2&login=" & Text3.Text)
        End If
    End If
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

Private Sub Timer2_Timer()
    If EnviarTimer.Enabled = True Then Exit Sub
    
    Select Case Ativado
        Case 0:
            Command1.Caption = "Divulgar"
            Command1.Enabled = True
            Status.Caption = "Parado!"
            Combo1.Enabled = True
            isButton2.Enabled = True
        Case 1:
            Command1.Caption = "Divulgando..."
            Command1.Enabled = False
            Status.Caption = "Aperte 'Enter' para Parar!"
            Combo1.Enabled = False
            isButton2.Enabled = False
    End Select
    
    If Ativado = 0 Then Exit Sub
    Proximo.Text = "Reenviando em " & (CInt(Combo2.Text) - Intervalo)
    Intervalo = Intervalo + 1
    If Intervalo > CInt(Combo2.Text) Then
        Intervalo = 0
        Call m_WebControl.object.navigate("http://munovus.net/site/atualizador_dv.php?f=3&login=" & Login & "&totalmsg=" & Total & "&intervalo=" & Combo2.Text)
        Call PegarTextoDV
        Call EnviarMSG
    End If
End Sub
