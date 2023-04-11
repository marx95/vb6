VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macro Divulgador - Equipe MuNovus.net"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4560
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
      Width           =   4335
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If App.PrevInstance Then End
End Sub

Private Sub Form_Paint()
    If PreLoad = 1 Then Exit Sub
    PreLoad = 1
    
    If Len(Dir(App.Path & "\Flash10c.ocx")) = 0 Then
        Call Baixar_OCX
    Else
        If FileLen(App.Path & "\Flash10c.ocx") < 3979680 Then Call Baixar_OCX
    End If
    
    Me.Visible = False
    Macro.Show
End Sub

Private Function Baixar_OCX()
    Status.Caption = "Baixando Flash10c.ocx - 4MB"
        DoEvents
        Call DownloadAFile(App.Path & "\Flash10c.ocx", "http://munovus.net/arquivos/Flash10c.ocx", False)
End Function

