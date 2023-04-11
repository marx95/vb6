VERSION 5.00
Begin VB.Form Editor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Mensagem"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox title2 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   8520
      Width           =   1695
   End
   Begin VB.TextBox mensagem 
      Height          =   6135
      Left            =   120
      MaxLength       =   2048
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox title 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo da mensagem de resposta (Máximos 50 dígitos)"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   8520
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "__"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo da Mensagem (Máximos 2048 caracteres)"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Titulo da mensagem do post (Máximo 50 dígitos)"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3420
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WriteINI App.Path & "/conf.ini", "conf", "titulo", title.Text
WriteINI App.Path & "/conf.ini", "conf", "resptitulo", title2.Text
Open App.Path & "/msg.txt" For Output As #1
    Print #1, mensagem.Text
Close #1
Principal.Load_Mensagem
Unload Me
End Sub

Private Sub Form_Load()
title.Text = ReadINI(App.Path & "/conf.ini", "conf", "titulo")
title2.Text = ReadINI(App.Path & "/conf.ini", "conf", "resptitulo")
Open App.Path & "/msg.txt" For Input As #1
    mensagem.Text = Input(FileLen(App.Path & "/msg.txt"), #1)
Close #1
Label11.Caption = "Restam " & (2048 - Len(mensagem.Text)) & " caracteres"
End Sub
Private Sub Label10_Click()
'SUBLINHE
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[u]" & mensagem.SelText & "[/u]")
End Sub

Private Sub Label3_Click()
'negrito
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[b]" & mensagem.SelText & "[/b]")
End Sub

Private Sub Label4_Click()
'vermelho
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[red]" & mensagem.SelText & "[/red]")
End Sub

Private Sub Label5_Click()
'Verde
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[green]" & mensagem.SelText & "[/green]")
End Sub

Private Sub Label6_Click()
'Azul
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[blue]" & mensagem.SelText & "[/blue]")
End Sub

Private Sub Label7_Click()
'amarelo
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[yellow]" & mensagem.SelText & "[/yellow]")
End Sub

Private Sub Label8_Click()
'pink
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[pink]" & mensagem.SelText & "[/pink]")
End Sub

Private Sub Label9_Click()
'Italico
mensagem.Text = Replace(mensagem.Text, mensagem.SelText, "[i]" & mensagem.SelText & "[/i]")
End Sub

Private Sub mensagem_Change()
Label11.Caption = "Restam " & (2048 - Len(mensagem.Text)) & " caracteres"
End Sub
