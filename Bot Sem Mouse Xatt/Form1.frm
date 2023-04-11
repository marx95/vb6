VERSION 5.00
Begin VB.Form Bot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bot"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Tempo 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   1680
   End
   Begin VB.ComboBox Xats 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Ligar 
      Caption         =   "Ligar"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Bot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Form_Load()
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Bot)
    m_WebControl.Move 1000, 1000, 6522, 6522
    m_WebControl.Visible = True
    
    Xats.AddItem "Selecione..."
    Xats.AddItem "PortalVCMuOnline"
    Xats.AddItem "fazendoradio"
    Xats.AddItem "CiadosMuBr2010"
    Xats.AddItem "ServidorMuNovus"
    Xats.Text = "Selecione..."
    
    For i = 4 To 6
        Tempo.AddItem i & "000"
    Next i
    Tempo.Text = "4000"
    
    Usuario = "DivulgadorMuNovus"
    Senha = "xaubet95"
    
    MsgDV(0) = "MU 99Z NOVINHO CASINHA GOGOGOGO www.munovus.net"
    MsgDV(1) = "GOGOGOG www.munovus.net BORA GENTE NOVINHO EXTREIO HOJE"
    MsgMax = 1
End Sub

Private Sub Ligar_Click()
    If Ligar.Caption = "Ligar" Then
        Ligar.Caption = "Deligar"
        Timer.Enabled = True
        Xats.Enabled = False
    Else
        Ligar.Caption = "Ligar"
        Timer.Enabled = False
        Xats.Enabled = True
    End If
End Sub

Private Sub m_WebControl_ObjectEvent(Info As EventInfo)
    If Info = "DocumentComplete" Then
        On Error Resume Next
        Source = m_WebControl.object.Document.documentelement.innerhtml
        
        If InStr(1, Source, "xat mobile login") Then
            Timer.Enabled = False
            Call Entrar_Xat
            Exit Sub
        End If
        
        If InStr(1, Source, "DoMessage") Then
            Call Enviar_Msg
            Timer.Enabled = True
            Exit Sub
        End If
        
    End If
End Sub

Private Sub Timer_Timer()
    If Xats.Text = "Selecione..." Then
        Call Ligar_Click
        Exit Sub
    End If
    Timer.Interval = CInt(Tempo.Text)
    Call Navegar("http://m.xat.com")
End Sub
