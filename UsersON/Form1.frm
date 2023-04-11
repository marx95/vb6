VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver total online - [MarxD]"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11310
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7590
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   6840
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   120
      List            =   "Form1.frx":0326
      TabIndex        =   6
      Text            =   "Selecione o IP"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Versão antiga (97D)"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer RefreshTM 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   6840
   End
   Begin Project1.Socket Sock 
      Left            =   600
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ListBox Servers 
      Height          =   5715
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   11055
   End
   Begin Project1.jcbutton Cnn 
      Height          =   300
      Left            =   9840
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Caption         =   "Conectar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox porta 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Text            =   "44405"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox ip 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   11160
      X2              =   11160
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "http://mxsv.sytes.net:81"
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
      Height          =   195
      Left            =   8880
      MouseIcon       =   "Form1.frx":03BA
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7320
      Width           =   2145
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   11055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ou"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   180
   End
   Begin VB.Label totalonCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cnn_Click()
Form1.Caption = "Ver total online - [MarxD] - Conectando..."
If Cnn.Caption = "Conectar" Then
    RefreshTM.Enabled = True
    Cnn.Caption = "Parar"
    Call Conectar
    Exit Sub
End If

If Cnn.Caption = "Parar" Then
    RefreshTM.Enabled = False
    Cnn.Caption = "Conectar"
    On Error Resume Next
    Sock.CloseSck
    Form1.Caption = "Ver total online - [MarxD] - Parado!"
    Label2.Caption = ""
    Call destravar
    Exit Sub
End If
End Sub
Private Sub destravar()
    Combo1.Enabled = True
    ip.Enabled = True
    porta.Enabled = True
    Check1.Enabled = True
End Sub
Private Sub travar()
    Combo1.Enabled = False
    ip.Enabled = False
    porta.Enabled = False
    Check1.Enabled = False
End Sub
Private Sub Conectar()
Dim ipCnn As String
If Combo1.Text <> "Selecione o IP" Then

    Select Case Combo1.Text
    Case "jogar.muaway.net"
        Check1.Value = 1
        ipCnn = "jogar.muaway.net"
        
    Case "MuDareDevils"
        ipCnn = "74.222.4.44"
    
    Case "MuC.A Brasil"
        ipCnn = "200.155.20.212"
    Case "ViperMU"
        ipCnn = "201.46.55.19"
    Case Else
        ipCnn = Combo1.Text
    End Select
    
Else

    If ip.Text = "" Then
        Form1.Caption = "Ver total online - [MarxD] - Digite um IP válido!"
        Cnn.Caption = "Conectar"
        Exit Sub
    Else
        Form1.Caption = "Ver total online - [MarxD]"
        ipCnn = ip.Text
    End If
    
End If

    On Error Resume Next
    Sock.CloseSck
    Sock.Connect ipCnn, porta.Text
    Label2.Caption = "[" & ipCnn & "]"
    Call travar
End Sub

Private Sub PegarListaSalas()
Dim pack(0 To 3) As Byte

pack(0) = &HC1
pack(1) = &H4
pack(2) = &HF4
pack(3) = &H6
Sock.SendData pack
End Sub

Private Sub AddSubSalas(Pacote As String)

Dim versao As Integer
Dim tmplol As String
Dim TotalOn As Integer
Dim PackProcess() As String
Dim Cu As String
TotalOn = 0

If Check1.Value = 0 Then
    versao = 3
End If
If Check1.Value = 1 Then
    versao = 2
End If

Servers.Clear


Cu = StringToHex(Pacote)
Cu = Right$(Cu, Len(Cu) - 18)



tmpolol = Replace(Cu, " ", "")
For i = 0 To ((Len(tmpolol) / 2) / 4)
    Dim bbosta As Integer
    Dim vbst As Integer
    
    bbosta = 1500 / 100
    Dim tmp() As String
    tmp = Split(Cu, " ")
    vbst = 1500 - (bbosta * CInt(Val("&H" & tmp(versao))))
    On Error Resume Next
    Servers.AddItem "Sala: " & i & " - Usuarios Online: " & vbst & " - " & 100 - CInt(Val("&H" & tmp(versao))) & "%"
    On Error Resume Next
    Cu = Right$(Cu, Len(Cu) - 12)
    
    
    On Error Resume Next
    TotalOn = TotalOn + (bbosta * CInt(Val("&H" & tmp(versao))))
Next i


totalonCap.Caption = "Total online: " & TotalOn
Form1.Caption = "Ver total online - [MarxD] - Atualizado!"
End Sub

Private Sub ip_Change()
    Combo1.Text = "Selecione o IP"
End Sub

Private Sub Label3_Click()
Call ShellExecute(Me.hWnd, vbNullString, Label3.Caption, vbNullString, vbNullString, 1)
End Sub

Private Sub RefreshTM_Timer()
    Call Conectar
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
Dim Pacote As String
Dim pacoteprocessado() As String
Sock.GetData Pacote
Form1.Caption = "Ver total online - [MarxD] - Atualizando..."
pacoteprocessado = Split(StringToHex(Pacote), " ")
Select Case pacoteprocessado(0)
Case "C1"
    If pacoteprocessado(2) = "00" Then
    Call PegarListaSalas
    Exit Sub
    End If
Case Else
    'If pacoteprocessado(3) = "F4" Then
    Call AddSubSalas(Pacote)
    Exit Sub
    'End If
End Select
MsgBox Pacote
End Sub

Private Sub Timer1_Timer()
If Label3.ForeColor = &H0& Then
    Label3.ForeColor = &HFF&
    Exit Sub
End If
If Label3.ForeColor = &HFF& Then
    Label3.ForeColor = &H0&
    Exit Sub
End If
End Sub
