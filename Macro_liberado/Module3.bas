Attribute VB_Name = "EnviarMsg_Mod"
Public Function Mostrar_Xat(Index As Integer)
    For i = 0 To TotalXats
        If CInt(i) = Index Then
            Form1.Shock(Index).Left = 120
            Form1.Xats.ListIndex = Index
        Else
            Form1.Shock(i).Left = Form1.Width + 15
        End If
    Next i
End Function

Public Function SumirXats()
    For i = 0 To TotalXats
        Form1.Shock(i).Left = Form1.Width + 15
    Next i
End Function
