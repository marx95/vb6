VERSION 5.00
Begin VB.Form Macro 
   BorderStyle     =   0  'None
   Caption         =   "Macro Divulgador - Equipe MuNovus.net"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox Tempo 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   6600
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Timer DVT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   6600
   End
   Begin MacroMuNovus.jcbutton Fechar 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
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
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":138CA
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MacroMuNovus.jcbutton Divulgar 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      Caption         =   "Divulgar"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":13BE4
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.ListBox Xats 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "Form1.frx":13EFE
      Left            =   6600
      List            =   "Form1.frx":13F00
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Info 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Divulgar_Click()
    If Divulgar.Caption = "Divulgar" Then
        Divulgar.Caption = "Parar"
        DVT.Interval = 1
        DVT.Enabled = True
    Else
        Divulgar.Caption = "Divulgar"
        DVT.Enabled = False
    End If
End Sub

Private Sub DVT_Timer()
    DVT.Interval = (CInt(Tempo.Text) + Rnd(201))
    
    Call Enviar_Msg
    TotalMsgs = TotalMsgs + 1
    Info.Caption = "Total Msg's enviado: " & TotalMsgs
End Sub

Private Sub Fechar_Click()
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = "x: " & X / 15 & " - y: " & Y / 15
End Sub

Private Sub Form_Load()
    
    MsgDV(0) = "SERVIDOR 99Z - GANHE 7DIAS VIP, 3000 PONTOS e 1 SET FULL"
    MsgDV(1) = "EXTREIOU HOJE GOGOGOGO"
    MsgDV(2) = "PRECISAMOS DE 3 GM's - CORRE GALERA"
    MsgDV(3) = "2 GOLDS A CADA 10 LEVEL - CORRE GALERA"
    MsgDV(4) = "99Z PONTUATIVO GOGOGOGO"
    MsgMax = 4
    
    Xats.AddItem "CiadosMU", 0
    Xats.AddItem "ViciadosMU", 1
    'Xats.AddItem "", 2
    'Xats.AddItem "", 3
    'Xats.AddItem "", 4
    'Xats.AddItem "", 5
   ' Xats.AddItem "", 6
    
    Xat(0) = "88716389"
    Xat(1) = "158331371"
    Xat(2) = ""
    Xat(3) = ""
    Xat(4) = ""
    Xat(5) = ""
    Xat(6) = ""
    
    For i = 3 To 7
        Tempo.AddItem CStr(CInt(i) * 1000)
    Next i
    Tempo.Text = "5000"
        
    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Macro)
    m_WebControl.Move 120, 120, 7000, 5000
    m_WebControl.Visible = True
    
    Xats.Top = m_WebControl.Top
    Xats.Left = m_WebControl.Width + m_WebControl.Left + 120
    Xats.Width = 120 * 15
    Xats.Height = m_WebControl.Height / 2
    
    Tempo.Top = Xats.Top + Xats.Height + 120
    Tempo.Left = m_WebControl.Width + m_WebControl.Left + 120
    Tempo.Width = Xats.Width
    Tempo.Height = Xats.Height
    
    Divulgar.Top = m_WebControl.Top + m_WebControl.Height + 120
    Info.Left = Divulgar.Left + Divulgar.Width + 120
    Info.Top = Divulgar.Top + 30
    Info.AutoSize = True
    Fechar.Top = m_WebControl.Top + m_WebControl.Height + 120
    
    Me.Width = Xats.Width + Xats.Left + 120
    Me.Height = Fechar.Top + Fechar.Height + 120
    Fechar.Left = Me.Width - Fechar.Width - 120
    
    Me.Visible = True
End Sub

Private Sub Xats_Click()
    If Aviso = 0 Then
        Aviso = 1
       ' MsgBox "Não esqueça de Trocar o Nick e Link da casinha!", vbInformation, "Aviso:"
        Divulgar.Enabled = True
    End If
    
    Call Navegar("http://www.xatech.com/web_gear/chat/chat.swf?id=" & Xat(Xats.ListIndex))
End Sub
