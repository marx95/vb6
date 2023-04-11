VERSION 5.00
Begin VB.Form JoinCmm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deseja entrar na Comunidade?"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "JoinCmm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4560
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Não"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sim"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "JoinCmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()                                                                                                            'table/tbody/tr[2]/td/form/div[4]/span/a
Principal.m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("mbox").childnodes.Item(0).childnodes.Item(0).childnodes.Item(1).childnodes.Item(0).childnodes.Item(0).childnodes.Item(5).childnodes.Item(0).childnodes.Item(0).Click
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Me.SetFocus
End Sub
