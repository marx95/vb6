VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macro Divulgador - Equipe MuOver.net"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If App.EXEName <> "MacroMuOver" Then
        MsgBox "O Executavel foi renomeado!", vbCritical, "Ocorreu um Erro!"
        End
    End If
    
    If Command = "reiniciar" Then
        Call Sleep(250)
    Else
        If App.PrevInstance Then End
    End If
    
    Dim TmpPathMacro As String
    Dim TmpRegMacro As String
    Dim Reiniciar As Integer
    
    TmpPathMacro = App.Path & "\MacroMuOver.exe"
    TmpRegMacro = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", TmpPathMacro)
 
    If Len(TmpRegMacro) < Len("RUNASADMIN DisableNXShowUI") Then
        Call SaveSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", TmpPathMacro, "RUNASADMIN DisableNXShowUI")
        Reiniciar = 1
    End If
    
    If Reiniciar = 1 Then
        Call ShellExecute(Macro.hWnd, vbNullString, App.EXEName & ".exe", "reiniciar", vbNullString, 0)
        Call Shell("MacroMuOver.exe reiniciar", vbNormalFocus)
        Call ExitProcess(0)
    End If
End Sub

Private Sub Form_Paint()
    If PreLoad = 1 Then Exit Sub
    PreLoad = 1
    
    Dim mcrc As New CRC
    If mcrc.CRC(App.Path & "\Flash10c.ocx") <> "-3295212959" Then Call Baixar_OCX
    
    Call Baixar_TXTDV
    
    Me.Visible = False
    Macro.Show
End Sub

Private Function Baixar_OCX()
    Status.Caption = "Baixando Flash10c.ocx - 4MB"
    DoEvents
    Call DownloadAFile(App.Path & "\Flash10c.ocx", "http://pgcontrol.com.br/muover/Flash10c.ocx", False)
    DoEvents
End Function

Private Function Baixar_TXTDV()
    Status.Caption = "Baixando Informações de Divulgação..."
    DoEvents
    
    On Error Resume Next
    Call Kill(App.Path & "\msgdv.txt")
    Call DownloadAFile(App.Path & "\msgdv.txt", "http://pgcontrol.com.br/muover/msg_macro_staff.txt", False)
    If Len(Dir(App.Path & "\msgdv.txt")) = 0 Then
        MsgBox "Falha ao baixar informações!", vbCritical, "Erro!"
        Call ExitProcess(0)
    Else
        Dim TmpMsgDv() As String
        Open App.Path & "\msgdv.txt" For Input As #1
            Dim TmpMsgs As String
            TmpMsgs = Replace(Input(FileLen(App.Path & "\msgdv.txt"), #1), vbNewLine, vbNullString)
            TmpMsgDv() = Split(TmpMsgs, "#$#")
        Close #1
        
        For i = 0 To UBound(TmpMsgDv)
            If Len(TmpMsgDv(i)) > 4 Then
                MsgDV(MsgMax) = TmpMsgDv(i)
                MsgMax = MsgMax + 1
            End If
        Next i
        MsgMax = MsgMax - 1
    End If
    
    On Error Resume Next
    Call Kill(App.Path & "\msgdv.txt")
End Function
