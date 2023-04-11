VERSION 5.00
Begin VB.Form ListaMonstros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "ListaMonstros.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton Selecionar 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   8040
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "Selecionar"
      MousePointer    =   99
      MouseIcon       =   "ListaMonstros.frx":1CCA
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox busca 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin VB.ListBox Lista 
      Height          =   6885
      ItemData        =   "ListaMonstros.frx":1FE4
      Left            =   240
      List            =   "ListaMonstros.frx":1FE6
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Buscar por: Nome Exato"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "ListaMonstros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub busca_Change()
    On Error Resume Next
    Lista.Text = busca.Text
End Sub

Private Sub busca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SetarMonstro
End Sub
Private Sub Form_Paint()
    busca.Text = ""
    If PreLoad = 0 Then
        PreLoad = 1
        DoEvents
        Dim i As Integer
        For i = 0 To 512
            If Len(Monstro_Nome(i)) > 4 Then Lista.AddItem Monstro_Nome(i)
        Next i
    End If
End Sub

Private Sub Lista_Click()
    Call SetarMonstro
End Sub

Private Sub Selecionar_Click()
    Call SetarMonstro
End Sub

Public Function SetarMonstro()
    Dim i As Integer
    For i = 0 To 512
            If Monstro_Nome(i) = Lista.Text Then
                SpotMaker.MonsterID.Text = Monstro_ID(i)
                Call PegarNomeMonstro(i)
                Me.Hide
                Exit Function
            End If
    Next i
End Function

Private Sub Form_Unload(cancel As Integer)
    cancel = 1
    Me.Hide
End Sub
