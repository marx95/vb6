VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   7
      Left            =   3840
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   6
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   5
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command2"
      Height          =   615
      Index           =   4
      Left            =   3840
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command2"
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const SW_NORMAL = 1
Private Const SW_MINIMIZE = 6
Private Const SW_MAXIMIZE = 3
Private Const WM_CLOSE = &H10

Dim CL As New Collection
Dim AppHwnd As Long

Private Sub TrataJanela(FN As Integer)

If AppHwnd > 0 Then
Select Case FN
Case 0 ' No Topo
SetWindowPos AppHwnd, -1, 0, 0, 0, 0, &H2 Or &H1
Case 1 ' Tras para Frente
SetForegroundWindow AppHwnd
Case 2 ' Minimiza
SetWindowPos AppHwnd, -2, 0, 0, 0, 0, &H2 Or &H1
ShowWindow AppHwnd, SW_MINIMIZE
Case 3 ' Normaliza
ShowWindow AppHwnd, SW_NORMAL
Case 4 ' Maximiza
SetWindowPos AppHwnd, -2, 0, 0, 0, 0, &H2 Or &H1
ShowWindow AppHwnd, SW_MAXIMIZE
Case 5 ' Esconde
SetWindowPos AppHwnd, -2, 0, 0, 0, 0, &H2 Or &H1
ShowWindow AppHwnd, SW_HIDE
Case 6 ' Mostra
ShowWindow AppHwnd, SW_SHOW
Case 7 ' Fecha
SetWindowPos AppHwnd, -2, 0, 0, 0, 0, &H2 Or &H1
PostMessage AppHwnd, WM_CLOSE, 0&, 0&
End Select
End If

End Sub

Private Sub Command1_Click(Index As Integer)
TrataJanela Index

End Sub

Private Sub Form_Load()

Dim I As Integer

Me.Caption = "Form1"

Command1(0).Caption = "Poe no topo"
Command1(1).Caption = "Tras p/ Frente"
Command1(2).Caption = "Minimiza"
Command1(3).Caption = "Normaliza"
Command1(4).Caption = "Maximiza"
Command1(5).Caption = "Esconde"
Command1(6).Caption = "Mostra"
Command1(7).Caption = "Fecha"

' Desabilita os botões
For I = 0 To Command1.UBound
Command1(I).Enabled = False
Next

' Guarda as classes numa collection
CL.Add "wndclass_desked_gsk", "Form1"
CL.Add "MuNovus.net", "MuNovus.net"
CL.Add "OpusApp", "MS Word"
CL.Add "OMain", "MS Access"
CL.Add "XLMAIN", "MS Excel"
CL.Add "wndclass_desked_gsk", "MS Visual Basic"
CL.Add "Notepad", "Notepad"
CL.Add "IEFrame", "Internet Explorer"
CL.Add "SciCalc", "Calculadora"
CL.Add "ConsoleWindowClass", "Prompt do DOS"
CL.Add "CabinetWClass", "Meu Computador"

' Enche o Listbox com as aplicações
With List1
.AddItem "Form1"
.AddItem "MuNovus.net"
.AddItem "MS Word"
.AddItem "MS Excel"
.AddItem "MS Access"
.AddItem "MS Visual Basic"
.AddItem "Notepad"
.AddItem "Internet Explorer"
.AddItem "Calculadora"
.AddItem "Prompt do DOS"
.AddItem "Meu Computador"
End With

End Sub

Private Sub List1_Click()

Dim I As Integer, R As Boolean

' Recupera o handle da janela escolhida no listbox
AppHwnd = FindWindow(CL.Item(List1.Text), vbNullString)

' Habilita os botões se o handle for válido
R = (AppHwnd > 0)
For I = 0 To Command1.UBound
Command1(I).Enabled = R
Next

End Sub
