Attribute VB_Name = "Globals"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (ipPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long

Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Global Ativado As Long
Global Intervalo As Long
Global MsgDV(1 To 5) As String
Global MsgID As Long
Global Total As Long

Global XatArq As String
Global MinhaCRC As String
Global Setado As Long

Global ChatID As Integer
Global TotalXats As Integer

Global PreloadInicio As Long
Global PreloadForm1 As Long

Global DVTime As Long
Global Source As String

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const Esquerdo_Down = &H2
Public Const Esquerdo_UP = &H4

Public Const Direito_Down = &H8
Public Const Direito_UP = &O10

Public Function LiberarMacro()
    If Setado = 1 Then Exit Function
    Setado = 1
    
    XatArq = "/Xats.ini"
    If CInt(ReadINI(App.Path & "/Config.ini", "Config", "Auto_Xat")) = 0 Then XatArq = "/Config.ini" ' - AQUI SETADO PELO USUARIO
    
    For i = 1 To TotalXats
        Form1.Caption = "Carregando Xat " & i & " - Macro do MuNovus.net"
        Load Form1.Shock(i)
            
        Form1.Shock(i).Top = Form1.Shock(0).Top
        Form1.Shock(i).Left = Form1.Shock(0).Left
        Form1.Shock(i).Width = Form1.Shock(0).Width
        Form1.Shock(i).Height = Form1.Shock(0).Height
        Form1.Shock(i).Movie = "http://www.xatech.com/web_gear/chat/chat.swf"
        Form1.Shock(i).FlashVars = "id=" & ReadINI(App.Path & XatArq, "Config", "Chat_" & i)
        Form1.Shock(i).Visible = True
    Next i
    
    Form1.Combo2.Clear
    For i = 3 To 7
        Form1.Combo2.AddItem i & " Segundos"
    Next i
    Form1.Combo2.Text = "3 Segundos"
    
    Form1.Xats.Clear
    Form1.Xats.AddItem "Portal GigaMU.net"
    For i = 1 To 8
        Form1.Xats.AddItem ReadINI(App.Path & XatArq, "Config", "Chat_" & i & "_Nome")
    Next i
    
    
    MsgDV(1) = ReadINI(App.Path & "/Config.ini", "Config", "MSG_1")
    MsgDV(2) = ReadINI(App.Path & "/Config.ini", "Config", "MSG_2")
    MsgDV(3) = ReadINI(App.Path & "/Config.ini", "Config", "MSG_3")
    MsgDV(4) = ReadINI(App.Path & "/Config.ini", "Config", "MSG_4")
    MsgDV(5) = ReadINI(App.Path & "/Config.ini", "Config", "MSG_5")
    
    Form1.Browser_timer.Enabled = True
    Form1.DvBt.Enabled = True
    Form1.Combo2.Enabled = True
    Form1.Xats.Enabled = True
    Form1.LiberarBt.Enabled = False
    
    Form1.Anuncio(1).Visible = True
    Form1.Anuncio(2).Visible = True
    Form1.Anuncio(3).Visible = True
    Form1.DvBt.Visible = True
    Form1.Combo2.Visible = True
    Form1.Label2.Visible = True
    Form1.Xats.Visible = True
    
    Form1.LiberarBt.Visible = False
    Form1.Label3.Visible = False
    Form1.Xats.Top = Form1.LiberarBt.Top
    
    Form1.Caption = "Divulgador - Macro do MuNovus.net"
    Call Mostrar_Xat(0)
End Function

Public Function SetarPosicaoDoMouse(x As Integer, Y As Integer)
    Call SetCursorPos((Form1.Left / 15) + x, (Form1.Top / 15) + Y)
End Function

Public Function Clicar_Esquerdo()
    mouse_event Esquerdo_Down, 0, 0, 0, 0
    mouse_event Esquerdo_UP, 0, 0, 0, 0
    DoEvents
End Function

Public Function Clicar_Direito()
    mouse_event Direito_Down, 0, 0, 0, 0
    mouse_event Direito_UP, 0, 0, 0, 0
End Function

Public Function ControlC()
    Call keybd_event(VK_CONTROL, 0, 0&, 0&)
    Call keybd_event(VK_C, 0, 0&, 0&)
    Call keybd_event(VK_CONTROL, 0, 2, 0&)
End Function

Public Function ControlV()
    Call keybd_event(VK_CONTROL, 0, 0&, 0&)
    Call keybd_event(VK_V, 0, 0&, 0&)
    Call keybd_event(VK_CONTROL, 0, 2, 0&)
    DoEvents
End Function

Public Function PegarTextoDV()
    MsgID = MsgID + 1
    If MsgID > 5 Then MsgID = 1
    Clipboard.Clear
    Clipboard.SetText MsgDV(MsgID) & "(DMD)"
End Function

Public Function SetarAnuncios()
    On Error Resume Next
    Form1.Anuncio(1).Picture = LoadPicture(App.Path & "/a1.jpg")
    Form1.Anuncio(2).Picture = LoadPicture(App.Path & "/a2.jpg")
    Form1.Anuncio(3).Picture = LoadPicture(App.Path & "/a3.jpg")
    Form1.Anuncio(1).Visible = True
    Form1.Anuncio(2).Visible = True
    Form1.Anuncio(3).Visible = True
End Function
