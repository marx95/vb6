Attribute VB_Name = "global"
Global Trafego(1) As Long
Global KeepAlive(32767) As Long

Global C_Min As Long
Global C_Max As Long

Function SetBytes(Bytes As Long) As String

    On Error GoTo hell

    If Bytes >= 1073741824 Then
        SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.0") _
                   & " GB"
    ElseIf Bytes >= 1048576 Then
        SetBytes = Format(Bytes / 1024 / 1024, "#0.0") & " MB"
    ElseIf Bytes >= 1024 Then
        SetBytes = Format(Bytes / 1024, "#0.0") & " KB"
    ElseIf Bytes < 1024 Then
        SetBytes = Fix(Bytes) & " B"
    End If

    Exit Function
hell:
    SetBytes = "0 Bytes"
End Function

Function C_IP(IPA As String, IPB As String) As Boolean
    If IPA = IPB Then
        C_IP = True
        Exit Function
    End If
    cip = False
End Function

