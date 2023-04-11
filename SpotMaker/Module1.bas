Attribute VB_Name = "Global"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Paint As Integer
Global Monstro_ID_Selecionado As String
Global Monstro_Nome_Selecionado As String
Global PreLoad As Integer
Global Monstro_ID(0 To 512) As Long
Global Monstro_Nome(0 To 512) As String

Global Salvo As Integer
Global Ultimo_Mapa As Integer
Global Total_Monstros As Integer

Public Function Travar()
    SpotMaker.Ver.Enabled = False
    SpotMaker.mapa.Enabled = False
    SpotMaker.ratio.Enabled = False
    SpotMaker.Direcao.Enabled = False
    SpotMaker.posX.Enabled = False
    SpotMaker.posY.Enabled = False
    SpotMaker.Quantidade.Enabled = False
    SpotMaker.Criar.Enabled = False
    SpotMaker.RadioValor.Enabled = False
    SpotMaker.Text1.Enabled = False
    SpotMaker.Guardar.Enabled = False
End Function

Public Function Liberar()
    SpotMaker.Ver.Enabled = True
    SpotMaker.mapa.Enabled = True
    SpotMaker.ratio.Enabled = True
    SpotMaker.Direcao.Enabled = True
    SpotMaker.posX.Enabled = True
    SpotMaker.posY.Enabled = True
    SpotMaker.Quantidade.Enabled = True
    SpotMaker.Criar.Enabled = True
    SpotMaker.RadioValor.Enabled = True
    SpotMaker.Text1.Enabled = True
    SpotMaker.Guardar.Enabled = True
End Function

Public Function PegarNomeMonstro(ID As Integer)
    If ID = 0 Then
        Monstro_Nome_Selecionado = "Bull Fighter"
        SpotMaker.Ver.Caption = Monstro_Nome_Selecionado
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To 512
            If Monstro_ID(i) = ID Then
                Monstro_Nome_Selecionado = Monstro_Nome(i)
                SpotMaker.Ver.Caption = Monstro_Nome_Selecionado
                Exit Function
            End If
    Next i
End Function
