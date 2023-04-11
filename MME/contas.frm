VERSION 5.00
Begin VB.Form contas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[MME] Contas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5175
   Icon            =   "contas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MudarInfo 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   5400
   End
   Begin Project1.jcbutton Command2 
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
   Begin Project1.jcbutton Command1 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Caption         =   "Atualizar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ver conta"
      Height          =   5295
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Senha"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Login"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o Login"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.ListBox listacontas 
         Appearance      =   0  'Flat
         Height          =   6075
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Procurar por login:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "contas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Atualiza_Contas
    rst.Update
    Call Carrega_Conta
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Public Function Busca(tmplogin As String)
On Error Resume Next
rst.Close
rst.CursorLocation = adUseClient
    rst.Open "SELECT memb___id, memb__pwd, mail_addr, migrado, migrado_cash FROM memb_info WHERE memb___id='" & tmplogin & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText
    Call Carrega_Conta
    contas.Caption = "[MME] Contas - " & tmplogin
    Frame2.Visible = True
    Command1.Visible = True
End Function

Private Sub Carrega_Conta()
    Text2.Text = rst.Fields(0)
    Text3.Text = rst.Fields(1)
    Text4.Text = rst.Fields(2)
    
    Select Case rst.Fields(3)
    Case 0
    Label1.Caption = "Conta Não Migrado"
    Case 1
    Label1.Caption = "Conta Já Migrado"
    End Select
    
    
    'banido
    Select Case rst.Fields(4)
    Case 0
    Label5.Caption = "Cash Não Migrado"
    Case 1
    Label5.Caption = "Cash Já Migrado"
    End Select
End Sub

Private Sub Atualiza_Contas()
    rst.Fields(0) = Text2.Text
    rst.Fields(1) = Text3.Text
    rst.Fields(2) = Text4.Text
    rst.Fields(4) = Text6.Text
    rst.Fields(5) = Text7.Text
    rst.Fields(7) = Text5.Text
    
    Select Case Combo1.Text
    Case "Free"
    rst.Fields(3) = 0
    Case "Vip"
    rst.Fields(3) = 1
    End Select
    
    'banido
     Select Case Combo2.Text
    Case "Desbanido"
    rst.Fields(5) = 0
    Case "Banido"
    rst.Fields(5) = 1
    End Select
    
    MudarInfo.Enabled = True
End Sub

Private Sub Limpa_Infos()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Form_Paint()
    Call PegaContas
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst.Close
End Sub

Public Function PegaContas()
listacontas.Clear
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT TOP 50 memb___id FROM memb_info", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

For i = 0 To 50
    listacontas.AddItem rst.Fields(0)
    rst.MoveNext
Next i
End Function

Public Function PegaContasPorLogin()
If Text1.Text = vbNullString Then
Exit Function
End If

listacontas.Clear
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT memb___id FROM memb_info WHERE memb___id='" & Text1.Text & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

For i = 0 To rst.RecordCount
listacontas.AddItem rst.Fields(0)
rst.MoveNext
Next i
End Function

Private Sub listacontas_Click()
    Call Limpa_Infos
    Call Busca(listacontas.Text)
End Sub

Private Sub MudarInfo_Timer()
MudarInfo.Interval = 1200

Select Case Info.Caption
Case ""
    Info.Caption = "Atualizado!"
    Me.Caption = " Contas - Atualizado!"
Case "Atualizado!"
    Info.Caption = ""
    Me.Caption = " Contas"
    MudarInfo.Enabled = False
    MudarInfo.Interval = 1
End Select
End Sub

Private Sub Text1_Change()
    Call PegaContasPorLogin
    contas.Caption = "[MME] Contas - Pressione Enter"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Limpa_Infos
        Call Busca(Text1.Text)
    End If
End Sub
