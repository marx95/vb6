Attribute VB_Name = "Cursor"
'Public Declare Function GetCursorPos Lib "user32" (ipPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const Esquerdo_Down = &H2
Public Const Esquerdo_UP = &H4

Public Const Direito_Down = &H8
Public Const Direito_UP = &O10

Public Function Mouse(X As Integer, Y As Integer)
    Call SetCursorPos((Macro.Left / 15) + X, (Macro.Top / 15) + Y)
    DoEvents
End Function

Public Function Clicar_Esquerdo()
    Call mouse_event(Esquerdo_Down, 0, 0, 0, 0)
    Call mouse_event(Esquerdo_UP, 0, 0, 0, 0)
    DoEvents
End Function

Public Function Clicar_Esquerdo_Duplo(Quantidade As Long)
    For i = 1 To Quantidade
        Call mouse_event(Esquerdo_Down, 0, 0, 0, 0)
        Call mouse_event(Esquerdo_UP, 0, 0, 0, 0)
        DoEvents
    Next i
End Function

Public Function Clicar_Direito()
    Call mouse_event(Direito_Down, 0, 0, 0, 0)
    Call mouse_event(Direito_UP, 0, 0, 0, 0)
    DoEvents
End Function

Public Function Focalizar_BTControle()
    Dim X As Integer
    Dim Y As Integer
    X = (Macro.dvBT.Left / 15) + (Macro.dvBT.Width / 15) / 2
    Y = (Macro.dvBT.Top / 15) + (Macro.dvBT.Height / 15) / 2
    Call Mouse(X, Y)
    On Error Resume Next
    Macro.dvBT.SetFocus
End Function
