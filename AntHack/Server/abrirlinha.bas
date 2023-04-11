Attribute VB_Name = "abrirlinha"
Function AbreLinha(MeuArq As String, aLinhaEscolhida As Integer) As String
    Dim iARQ, QtdLinha As Integer
    Dim sLinha, meuResultado As String
    iARQ = FreeFile
    Open MeuArq For Input As iARQ
    Do While Not EOF(iARQ)
        Line Input #iARQ, sLinha
        If QtdLinha = aLinhaEscolhida - 1 Then
             AbreLinha = sLinha
        End If
        QtdLinha = QtdLinha + 1
    Loop
    Close iARQ
    If aLinhaEscolhida > QtdLinha Then
    Limite = 1
    End If
End Function
