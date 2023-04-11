VERSION 5.00
Begin VB.Form Administrar_Jogadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrar jogadores"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "Administrar_Jogadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconectar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Atualizar lista"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ListBox jogadores_on 
      Height          =   3960
      ItemData        =   "Administrar_Jogadores.frx":1BB2
      Left            =   120
      List            =   "Administrar_Jogadores.frx":1BB4
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Administrar_Jogadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call PegaListaDeUsuariosOnline
End Sub

Private Sub Command2_Click()
Dim User As String
User = Mid(jogadores_on.Text, 2, 1)
Server.Sock(User).CloseSck
AddLog "Cliente(" & User & ") - Disconectado"
Call PegaListaDeUsuariosOnline
Call AtualizaTitulo
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_load()
Call PegaListaDeUsuariosOnline
End Sub

Private Sub PegaListaDeUsuariosOnline()
jogadores_on.Clear
For i = 1 To UsuariosMaximos
If Server.Sock(i).State = sckConnected Then

If Personagem(i).Login <> "" And Personagem(i).Nome <> "" Then
jogadores_on.AddItem "(" & i & ") - Login: " & Personagem(i).Login & " - Personagem: " & Personagem(i).Nome
Else
If Personagem(i).Login <> "" Then
jogadores_on.AddItem "(" & i & ") - Login:" & Personagem(i).Login & " - Seleção de personagem"
Else
jogadores_on.AddItem "(" & i & ") - " & Personagem(i).IP & " - Seleção de servidor"
End If
End If
End If
Next i
End Sub
