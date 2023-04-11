Attribute VB_Name = "global"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Host As String
Global Usuario As String
Global AnexoLink As String
Global Senha As String
Global Banco As String
Global cnn As New ADODB.Connection
Global rst As New ADODB.Recordset
Global Cnn2 As String
Global LinkAvatar As String
Global Inicio As Integer
Global Fim As Integer
Global Status As Integer
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
Global Intervalo As Integer
Global Request As Integer

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
    MainForm.Show
End Sub
