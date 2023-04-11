VERSION 5.00
Begin VB.Form MuMsg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MUMSG-MXD"
   ClientHeight    =   1425
   ClientLeft      =   17130
   ClientTop       =   9255
   ClientWidth     =   3000
   Icon            =   "MuMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer AppRun 
      Interval        =   5000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer Deslizar 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer Atualizar 
      Interval        =   100
      Left            =   2400
      Top             =   120
   End
   Begin VB.Label Msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2760
   End
End
Attribute VB_Name = "MuMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Atualizar_Timer()
    Atualizar.Interval = 3000
    
    On Error GoTo erro
    If CInt(ReadINI(App.Path & "/Config.ini", "Config", "MuMsg")) = 0 Then End
    Call Navegar(MuMsgLink)
    
    Exit Sub
erro:
    Call WriteINI(App.Path & "/Config.ini", "Config", "MuMsg", "0")
End Sub

Private Sub Deslizar_Timer()
    Deslizar.Interval = 50
    
    If Me.Visible = False Then
        Deslizar.Enabled = False
        Exit Sub
    End If

    If Me.Left < Screen.Width Then
        
        Dim Resto As Long
        Resto = Screen.Width - Me.Left
        If Resto > 2000 Then
            Me.Left = Me.Left + 45
        Else
            Me.Left = Me.Left + 90
        End If
    Else
        Deslizar.Enabled = False
        Exit Sub
    End If
End Sub

Public Function Existe(Arq As String) As Boolean
    If Len(Dir(Arq)) > 4 Then Existe = True
    Existe = False
End Function
Private Sub Form_Load()
    If App.PrevInstance Then End
    
    If Existe(App.Path & "/tip.wav") = False Then
        Call DownloadAFile(App.Path & "/tip.wav", "http://munovus.net/arquivos/tip.wav", False, False)
    Else
        If FileLen(Dir(App.Path & "/tip.wav")) <> 104626 Then
            Call DownloadAFile(App.Path & "/tip.wav", "http://munovus.net/arquivos/tip.wav", False, False)
        End If
    End If
    
    
    Call Tranparencia
    Call DeleteKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\Explorer\Navigating\.Current")
    
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", MuMsg)
    m_WebControl.Move -15, -15, 15, 15
    m_WebControl.Visible = True
    
    MuMsgLink = "http://munovus.net/site/cache/mumsg.html"
    Delay_Desaparecer = 3000
    
    Call Navegar(MuMsgLink)
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    If Info = "DocumentComplete" Then
        
        On Error Resume Next
        Source = m_WebControl.object.document.documentelement.innerhtml
        
        Source = Replace(Source, "<TITLE>", "")
        Source = Replace(Source, "</TITLE>", "")
        Source = Replace(Source, "<HEAD>", "")
        Source = Replace(Source, "</HEAD>", "")
        Source = Replace(Source, "<BODY>", "")
        Source = Replace(Source, "</BODY>", "")
        Source = Replace(Source, vbNewLine, "")
        
        Call VerificarSource
    End If
End Sub

Private Sub Msg_Click()
    Call Sumir
End Sub

Private Sub Titulo_Click()
    Call Sumir
End Sub

Private Sub Form_click()
    Call Sumir
End Sub
