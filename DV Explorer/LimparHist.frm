VERSION 5.00
Begin VB.Form LimparHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limpando Historico de navegação"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "LimparHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clique aqui para limpar o historico do navegador"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"LimparHist.frx":1CCA
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "LimparHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", vbNormalFocus
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351", vbNormalFocus
Principal.m_WebControl.object.navigate "http://orkut.com.br"
Unload Me
End Sub
