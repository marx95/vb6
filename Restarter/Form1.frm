VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restarter"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Log 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   720
      Top             =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If InStr(1, Command, " ") Then
        Dim spt As String
        splt = Split(Command, " ")
        Executavel = splt(0)
    Else
        Executavel = Command
    End If
    Me.Caption = "Restarter - " & Executavel
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If ProcessoExiste(Executavel) = False Then
        On Error Resume Next
        Shell Command, vbHide
        Log.AddItem Executavel & " reiniciado!!!"
    End If
End Sub
