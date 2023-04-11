VERSION 5.00
Begin VB.Form AdicionarCMM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar Comunidades"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "AdicionarCMM.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4800
      Top             =   720
   End
   Begin VB.TextBox cmmTemp 
      Height          =   375
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox link 
      Height          =   405
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Ex: http://www.orkut.com.br/Main#Community?rl=cpn&cmm=559494"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cole o Link da comunidade a seguir:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2580
   End
End
Attribute VB_Name = "AdicionarCMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call PegaTempLista
End Sub
Private Sub PegaTempLista()
Open App.Path & "/comunidades.txt" For Input As #1
    cmmTemp.Text = Input(FileLen(App.Path & "/comunidades.txt"), #1)
Close #1
End Sub
Private Sub link_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call Verifica
End If
End Sub
Private Sub Verifica()
Dim TempLink As String
If Link.Text = "" Then
Status.Caption = "Digite coloque um link!"
Exit Sub
End If

TempLink = Replace(Link.Text, "http://www.orkut.com.br/Main#Community?rl=cpn&cmm=", "")
TempLink = Replace(Link.Text, "http://www.orkut.com.br/Main#Community?cmm=", "")
For i = 1 To 100
Dim Linha As Integer
Linha = i

If AbreLinha(App.Path & "/comunidades.txt", Linha) = TempLink Then
Status.Caption = "Link Existente!"
Link.Text = ""
Exit For
End If

If Linha = 100 Then
Call AddLinkCmm(TempLink)
Exit For
End If
Next i
End Sub

Private Sub AddLinkCmm(IDL As String)
cmmTemp.Text = cmmTemp.Text + vbNewLine + IDL
Open App.Path & "/comunidades.txt" For Output As #1
    Print #1, cmmTemp.Text
Close #1
Status.Caption = "Adicionado com Sucesso"
Call Principal.Load_Mensagem
Call Principal.PuxaLista
Call PegaTempLista
Link.Text = ""
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Link.SetFocus
End Sub
