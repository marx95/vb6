Attribute VB_Name = "Module1"
Global WindowsTitulo As String
Global Logado As Integer
Global ID_Post As Integer
Global ID_CMM As Integer
Global Login As String
Global Senha As String
Global Divulgando As Integer
Global TotalMsgSend As Integer
Global TotalCaptchaSend As Integer
Global PrecisaResp As Integer
Global Respondendo As Integer
Global LinkResp As String
Global DimMensagemResp As String
Global LastID As Integer
Global Pausado As Integer

Global Iniciado As Integer
Global Criar As Integer
Global Responder As Integer

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
Global contagem As Integer
Global tempo As String
Global killer As String
Global chera As String
Global dir As String

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function
Public Sub Main()
   InitCommonControlsVB
  Principal.Show
End Sub


