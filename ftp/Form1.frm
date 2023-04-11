VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ftp DNS System"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox IPCache 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Salvar(Arq As String, Texto As String)
    Open Arq For Output As #1
        Print #1, Texto
    Close #1
End Sub

Private Sub Command1_Click()
    If MsgBox("Tem certeza que deseja fechar?", vbYesNo, "Fechar?") = vbYes Then Call ExitProcess(0)
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call ReloadFtpIP
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then Call ExitProcess(0)
    IPCache.Clear
    Call ReloadFtpIP
End Sub

Private Sub ReloadFtpIP()
    Dim TmpIP As String
    Dim UltimoIP As String
    Dim Data As String
    
    TmpIP = PegarIP(ReadINI(App.Path & "/Config.ini", "Config", "URL_IP"))
    UltimoIP = ReadINI(App.Path & "/Config.ini", "Cache", "Ultimo_IP")
    Data = Now
    
    If Len(TmpIP) < 8 Then
        Exit Sub
    Else
        Dim Pontos As Integer
        For i = 1 To Len(TmpIP)
            If Mid(TmpIP, i, 1) = "." Then Pontos = Pontos + 1
        Next i
        If Pontos < 3 Then Exit Sub
    End If
    
    If TmpIP <> UltimoIP Then
        
        Dim ftpDir As String
        If Len(ReadINI(App.Path & "/Config.ini", "Config", "Ftp_Dir")) > 3 Then
            ftpDir = "cd " & ReadINI(App.Path & "/Config.ini", "Config", "Ftp_Dir")
        End If
        
        IPCache.AddItem "[" & Data & "] " & UltimoIP & " -> " & TmpIP
        Call WriteINI(App.Path & "/Config.ini", "Cache", "Ultimo_IP", " " & TmpIP)
        Call WriteINI(App.Path & "/Config.ini", "Cache", "Data", " " & Data)
        Call Salvar(App.Path & "\ip.txt", "0#" & TmpIP & "#2#" & Data & "#4")
        Call Salvar(App.Path & "\Cmd.txt", "open " & (ReadINI(App.Path & "/Config.ini", "Config", "Ftp_Addr")) & vbNewLine & "user " & (ReadINI(App.Path & "/Config.ini", "Config", "Ftp_User")) & " " & (ReadINI(App.Path & "/Config.ini", "Config", "Ftp_Pass")) & vbNewLine & ftpDir & vbNewLine & "binary" & vbNewLine & "put ip.txt" & vbNewLine & "quit")
        Call Salvar(App.Path & "\Exec.bat", "ftp -n -i -s:Cmd.txt" & vbNewLine & "del Cmd.txt" & vbNewLine & "del ip.txt" & vbNewLine & "del Exec.bat")
        Shell "Exec.bat", vbHide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Tem certeza que deseja fechar?", vbYesNo, "Fechar?") = vbYes Then ExitProcess (0)
    Cancel = 1
End Sub

Private Sub Timer1_Timer()
    Intervalo = Intervalo + 1
    Command2.Caption = (6 - Intervalo) & " (Atualizar)"
    If Intervalo > 5 Then
        Command2.Enabled = False
        Intervalo = 0
        Call ReloadFtpIP
    Else
        Command2.Enabled = True
    End If
End Sub
