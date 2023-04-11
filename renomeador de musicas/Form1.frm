VERSION 5.00
Begin VB.Form Renomeador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renomeador de Músicas"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   15855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "<- Inverter ->"
      Height          =   255
      Left            =   8400
      TabIndex        =   21
      Top             =   660
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Retirar Numeros dos Nomes das Músicas"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Renomear Tudo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   14040
      TabIndex        =   17
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pré-visualizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12240
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox ListaFinal 
      Height          =   4740
      Left            =   12600
      TabIndex        =   14
      Top             =   960
      Width           =   3135
   End
   Begin VB.ListBox NomesMusicas 
      Height          =   4740
      Left            =   9000
      TabIndex        =   12
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vai..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Trocar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Procurar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.ListBox ArtistaLista 
      Height          =   4740
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Inversor 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   " - "
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox PreLista 
      Height          =   4740
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listar Arquivos"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      TabIndex        =   19
      Top             =   6000
      Width           =   4965
   End
   Begin VB.Label Total 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   18
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final..."
      Height          =   195
      Left            =   12600
      TabIndex        =   15
      Top             =   660
      Width           =   465
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Músicas"
      Height          =   255
      Left            =   9720
      TabIndex        =   13
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trocar"
      Height          =   195
      Left            =   6000
      TabIndex        =   10
      Top             =   5760
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procurar"
      Height          =   195
      Left            =   4560
      TabIndex        =   9
      Top             =   5760
      Width           =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Artistas"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label aaa 
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos Originais"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inverter por string:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Renomeador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Inversor.Text = Replace(Inversor.Text, " ", vbNullString)
    If Len(Inversor.Text) = 0 Then Exit Sub
    Inversor.Enabled = False
    Command1.Enabled = False
    
    DoEvents

    Dim sPath As String
    Dim sFIle As String

    sPath = App.Path & "\"
    sFIle = Dir(sPath & "*.mp3", vbDirectory)

    Do While sFIle <> ""
        If InStr(sFIle, ".mp3") > 0 Then
            Call PreLista.AddItem(sFIle)
        End If
        sFIle = Dir
    Loop
    
    For i = 0 To PreLista.ListCount
        Dim NovoNome As String
        Dim TmpNome As String
        Dim tmpSplit() As String
        
        TmpNome = vbNullString
        TmpNome = PreLista.List(i)
        TmpNome = Replace(TmpNome, "_", " ")
        TmpNome = Replace(TmpNome, ".mp3", vbNullString)
        TmpNome = Replace(TmpNome, "  ", " ")
        
        ' - ############ Corrige numeros na frente do musica
        If Check1.Value = 1 Then
            Dim O As Integer
            O = 99
            While O > 0
                Dim Pre As String
                Pre = vbNullString
                If Len(O) = 1 And O > 1 Then Pre = "0"
                
                TmpNome = Replace(TmpNome, O & ".", vbNullString)
                TmpNome = Replace(TmpNome, Pre & O & ".", vbNullString)
                
                ' - Caso com Traço
                TmpNome = Replace(TmpNome, O & "-", vbNullString)
                TmpNome = Replace(TmpNome, O & " -", vbNullString)
                TmpNome = Replace(TmpNome, " -" & O, vbNullString)
                TmpNome = Replace(TmpNome, " - " & O, vbNullString)
            
                TmpNome = Replace(TmpNome, Pre & O & "-", vbNullString)
                TmpNome = Replace(TmpNome, Pre & O & " -", vbNullString)
                TmpNome = Replace(TmpNome, " -" & O & Pre, vbNullString)
                TmpNome = Replace(TmpNome, " - " & O & Pre, vbNullString)
                O = O - 1
            Wend
        End If
        
        TmpNome = Replace(TmpNome, " 0 ", vbNullString)
        TmpNome = Replace(TmpNome, "   ", " ")
        TmpNome = Replace(TmpNome, "  ", " ")
        
        If Mid(TmpNome, 1, 1) = " " Then TmpNome = Mid(TmpNome, 2, Len(TmpNome))
        
        On Error Resume Next
        tmpSplit = Split(TmpNome, Inversor.Text)
        
        ' - ############ ARTISTA
        Dim Artista As String
        Dim ArtistaSplit() As String
        
        ArtistaSplit = Split(tmpSplit(0), " ")
        Artista = vbNullString
        
        For a = 0 To UBound(ArtistaSplit)

            If Len(ArtistaSplit(a)) > 1 Then
                ArtistaSplit(a) = UCase(Mid(ArtistaSplit(a), 1, 1)) & LCase(Mid(ArtistaSplit(a), 2, Len(ArtistaSplit(a)) - 1))
            Else
                If ArtistaSplit(a) = "e" Then ArtistaSplit(a) = "&"
            End If
        Next a
        
        Artista = Replace(Join(ArtistaSplit, " "), ".mp3", vbNullString)
        
        ' - ############ Nome Musica
        Dim NomeMusica As String
        Dim MusicaSplit() As String
        
        MusicaSplit = Split(tmpSplit(1), " ")
        NomeMusica = vbNullString
        
        For a = 0 To UBound(MusicaSplit)

            If Len(MusicaSplit(a)) > 1 Then
                MusicaSplit(a) = UCase(Mid(MusicaSplit(a), 1, 1)) & LCase(Mid(MusicaSplit(a), 2, Len(MusicaSplit(a)) - 1))
            Else
                'If MusicaSplit(a) = "e" Then MusicaSplit(a) = "&"
            End If
        Next a
        
        NomeMusica = Replace(Join(MusicaSplit, " "), ".mp3", vbNullString)
        ArtistaLista.AddItem Artista
        NomesMusicas.AddItem NomeMusica
    Next i

    ArtistaLista.List(ArtistaLista.ListCount - 1) = vbNullString
    NomesMusicas.List(NomesMusicas.ListCount - 1) = vbNullString
    
    Call Command3_Click
    Total.Caption = PreLista.ListCount & " Música(s)"
End Sub

Private Sub Command2_Click()
    If Travar = 1 Then Exit Sub
    For i = 0 To ArtistaLista.ListCount

        Dim Artista As String
        Dim ArtistaSplit() As String
        
        ArtistaSplit = Split(ArtistaLista.List(i), " ")
        Artista = vbNullString
        
        For B = 0 To UBound(ArtistaSplit)
            If LCase(ArtistaSplit(B)) = LCase(Procurar.Text) Then ArtistaSplit(B) = Trocar.Text
        Next B
        
        ArtistaLista.List(i) = Join(ArtistaSplit, " ")
    Next i
    
    Call Command3_Click
End Sub

Private Sub Command3_Click()
    ListaFinal.Clear
    Dim i As Integer
    For i = 0 To PreLista.ListCount
        Dim TmpNome As String
        TmpNome = vbNullString
        
        Dim NomeDaMusica As String
        Dim NomeDoArtista As String
        
        NomeDaMusica = vbNullString
        NomeDoArtista = vbNullString
        
        If (Len(NomesMusicas.List(i)) > 0) Then NomeDaMusica = NomesMusicas.List(i)
        If (Len(ArtistaLista.List(i)) > 0) Then NomeDoArtista = ArtistaLista.List(i)
        
        TmpNome = NomeDaMusica & " " & divisor & " - " & NomeDoArtista
        TmpNome = Replace(TmpNome, "   ", " ")
        TmpNome = Replace(TmpNome, "  ", " ")
        
        If Mid(TmpNome, Len(TmpNome), 1) = " " Then
            TmpNome = Mid(TmpNome, 1, (Len(TmpNome) - 1))
        End If
        
        TmpNome = Replace(TmpNome, ".mp3", vbNullString)
        
        If Mid(TmpNome, 1, 1) = " " Then TmpNome = Mid(TmpNome, 2, Len(TmpNome))
        ListaFinal.AddItem TmpNome
    Next i
    
    ListaFinal.List(ListaFinal.ListCount - 1) = vbNullString
    
    DoEvents
    Command3.Enabled = True
    Command4.Enabled = True
    Command2.Enabled = True
    Procurar.Enabled = True
    Trocar.Enabled = True
End Sub

Private Sub Command4_Click()
    If Travar = 1 Then Exit Sub
    Travar = 1
    Command3.Enabled = False
    Command4.Enabled = False
    Command2.Enabled = False
    Check1.Enabled = False
    Command6.Enabled = False
    Command2.Enabled = False
    Procurar.Enabled = False
    Trocar.Enabled = False
    DoEvents
    
    Call Command3_Click
    
    DoEvents
    For i = 0 To PreLista.ListCount
        DoEvents
        Dim NNome As String
        NNome = Replace(ListaFinal.List(i), ".mp3", vbNullString) & ".mp3"
        
        On Error Resume Next
        Name App.Path & "\" & PreLista.List(i) As App.Path & "\" & NNome
    
        DoEvents
        Info.Caption = "Renomeando " & i & " de " & PreLista.ListCount
    Next i
    
    Command4.Enabled = False
    Info.Caption = "Sucesso, pode fechar!"
End Sub

Private Sub Command6_Click()
    If Label3.Caption = "Artistas" Then
        Label3.Caption = "Músicas"
        Label6.Caption = "Artistas"
    Else
        Label3.Caption = "Artistas"
        Label6.Caption = "Músicas"
    End If
End Sub

Private Sub Form_Load()
    Travar = 0
End Sub
