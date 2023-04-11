Attribute VB_Name = "Globals"
Public Declare Function GetCursorPos Lib "user32" (ipPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Global Ativado As Integer
Global Intervalo As Long
Global MsgDV(1 To 15) As String
Global MsgID As Integer
Global Total As Long
Global ChatID As Integer

Global TotalXats As Integer

Global Login As String
Global DVTime As Integer
Global Source As String

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const Esquerdo_Down = &H2
Public Const Esquerdo_UP = &H4

Public Const Direito_Down = &H8
Public Const Direito_UP = &O10

Public Function SetarLogin()
    Login = Form1.Text3.Text
    Form1.Animacao1LB.Visible = False
    Form1.Animacao1.Enabled = False
    Form1.Text3.Visible = False
    Form1.isButton3.Visible = False
    Form1.Label3.Visible = False
    Form1.Label5.Visible = False
    Form1.Label6.Visible = False
    Form1.Command1.Enabled = True
    Form1.Combo1.Enabled = True
    Form1.Combo2.Enabled = True
    Form1.isButton2.Enabled = True
    Form1.Timer2.Enabled = True
    Form1.Label4.Caption = "Autenticado com sucesso no servidor MuNovus!"
    Call MostraMuNovusXat
End Function

Public Function SetarPosicaoDoMouse(X As Integer, Y As Integer)
    SetCursorPos (Form1.Left / 15) + X, (Form1.Top / 15) + Y
End Function

Public Function Clicar_Esquerdo()
    mouse_event Esquerdo_Down, 0, 0, 0, 0
    mouse_event Esquerdo_UP, 0, 0, 0, 0
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
End Function

Public Function PegarTextoDV()
    MsgID = MsgID + 1
    If MsgID > 8 Then MsgID = 1
    Clipboard.Clear
    Clipboard.SetText MsgDV(MsgID) & "(DMD)"
End Function
