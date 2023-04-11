VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "_mx_MarxD_System_1.2"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MarxD.jcbutton Fechar 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "X"
      ForeColorHover  =   255
      MousePointer    =   99
      MouseIcon       =   "Main.frx":78B5
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      MaskColor       =   255
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin MarxD.Progressbar_user PB 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   238
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
      Color           =   255
      Scrolling       =   3
   End
   Begin VB.Timer Timer_Hacks 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   2880
   End
   Begin VB.Timer Loading 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2400
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   1920
      Width           =   3555
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fechar_Click()
    Call PararTudo
    End
End Sub

Private Sub Form_Load()
If Command$ <> "MxHostAntiHack" Or App.EXEName <> "AntiHack" Or App.PrevInstance Then End

Login = GetSettingString(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ID")
AutoClick = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "AutoClick")

Me.Width = 3720
Me.Height = 2295

PB.Color = &HFF&
PB.Width = 3720
PB.Scrolling = 3
Info.ForeColor = &HFFFFFF

ExecutavelDoMain = "Main.exe"
Janela = "MuNovus.net"      ' - Nome da janela do main
IP = "mxsv.sytes.net"
Porta = 90
MainCRC = "4837947665"
JogoAberto = 0

Call CarregaHacks
Call Otimizar    ' - Otimizar processo
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoverFormSemCaption(Button, hWnd)
End Sub

Private Sub Form_Paint()
    If MainCRC = mcrc.CRC(App.Path & "\" & ExecutavelDoMain) Then
        MainVerificado = 1
    Else
        Call Reportar_MainModificado
    End If
    
    Call PegaHacks
    Loading.Enabled = True
End Sub

Private Sub form_unload(Calcel As Integer)
    Call ForaDaTray(hWnd)
    End
End Sub

Private Sub Loading_Timer()
Loading.Interval = 50

' - ################################################################################################
If PB.Value <= 40 Then
    PB.Value = PB.Value + 2
    Exit Sub ' - Nao deixa executar a continuação caso seja menor ou igual q 40
End If

If PB.Value >= 40 And PB.Value < 100 Then
    On Error GoTo erro
    PB.Value = PB.Value + 5
    Exit Sub ' - Nao deixa executar a continuação caso seja menor q 100
End If

' - ################################################################################################
If MainVerificado = 1 And JogoAberto = 0 Then
    Dim ShellInfoMain As Integer
    Loading.Enabled = False
    JogoAberto = 1
    
    Call VaiPraTray(hWnd, Me.Icon, "MuNovus AntiHack 1.2")
    ShellInfoMain = ShellExecute(Me.hWnd, vbNullString, ExecutavelDoMain, vbNullString, vbNullString, 1)
    
    If AutoClick = 1 Then
        Call ShellExecute(Me.hWnd, vbNullString, "AutoClick.exe", Janela, vbNullString, 0)
    End If
    
    If ShellInfoMain = 2 Then
        Call PararTudo
        Info.Caption = "Falha ao executar o Main.exe"
        Exit Sub
    End If
    Main.Visible = False
    Timer_Hacks.Enabled = True
End If

Exit Sub
erro:
PB.Value = 100
End Sub

Private Sub Timer_Hacks_Timer()
If JogoAberto = 1 Then
    If ProcessoExiste(ExecutavelDoMain) = False Then ' - este if deixa o antihack aberto caso o main esteja aberto
         Call ForaDaTray(hWnd)
         Call KillProcess("AutoClick.exe")
         End
    Else ' - este é o else caso o findwindow Janela seja true
        If GetCaption(GetForegroundWindow) = Janela Then ' - Se a janela do main estiver focada, ele aumenta o interval para nao travar
            Timer_Hacks.Interval = 300
        Else ' - Se nao estiver focada, diminui o interval
            Timer_Hacks.Interval = 50
            Call PegaHacks
        End If
    End If
End If
End Sub

Public Function PegaHacks()
Dim CrcErrado As Integer
CrcErrado = 0

On Error Resume Next
If mcrc.CRC(App.Path & "/Data/Player/Player.bmd") <> "-1053799656" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World1/EncTerrain1.att") <> "1655529064" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World1/Terrain1.att") <> "-2000414755" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World3/EncTerrain3.att") <> "-152324027" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World3/Terrain3.att") <> "-1290381860" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World4/EncTerrain4.att") <> "2092968316" Then CrcErrado = 1
On Error Resume Next
If mcrc.CRC(App.Path & "/Data/World4/Terrain4.att") <> "-1401063478" Then CrcErrado = 1

If CrcErrado = 1 Then
    Call PararTudo
    Call ForaDaTray(hWnd)
    Call KillProcess("Main.exe")
    Call KillProcess("AutoClick.exe")
    Call KillProcess("Launcher.exe")
    Call HackEncontrado
    Exit Function
End If

For i = 0 To UBound(Hacks)

If Hacks(i) = vbNullString Then
    Exit Function
Else
    Dim TmpNomeJanela As String
    TmpNomeJanela = GetCaption(GetForegroundWindow)
    If InStr(1, LCase(TmpNomeJanela), Hacks(i)) Then
        Call PararTudo
        Call ForaDaTray(hWnd)
        Call KillProcess("Main.exe")
        Call KillProcess("AutoClick.exe")
        Call KillProcess("Launcher.exe")
        Call HackEncontrado
        Exit Function
    End If
End If
Next i
End Function
