Attribute VB_Name = "EnviarMsg_Mod"
Public Function MostraMuNovusXat()
    Form1.SWF(0).Visible = False
    For i = 1 To TotalXats
        Form1.SWF(i).Visible = False
    Next i
    Form1.SWF(0).Visible = True
End Function

Public Function SumirXats()
    For i = 0 To TotalXats
        Form1.SWF(i).Visible = False
    Next i
End Function

Public Function EnviarMSG()
    Form1.EnviarTimer.Enabled = True
    
    On Error Resume Next
    Form1.Caption = "Divulgador - MuNovus.net - [" & Total & " msg's em enviadas]"
    Call MostraMuNovusXat
End Function
