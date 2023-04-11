VERSION 5.00
Begin VB.Form Dep 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[MME] Depósitos"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton jcbutton4 
      Height          =   285
      Left            =   8040
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Ver"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Dep.frx":0000
      Left            =   120
      List            =   "Dep.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   120
      Width           =   3255
   End
   Begin Project1.jcbutton jcbutton3 
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Recusar pagamento"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton jcbutton2 
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Aprovar pagamento"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   2685
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Fechar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin Project1.jcbutton jcbutton5 
      Height          =   285
      Left            =   8040
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Ver"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton jcbutton6 
      Height          =   285
      Left            =   8040
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Ver"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Comentario"
      Height          =   195
      Left            =   6000
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Anexo 3"
      Height          =   195
      Left            =   6000
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Anexo 2"
      Height          =   195
      Left            =   6000
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Anexo 1"
      Height          =   195
      Left            =   6000
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Numero"
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hora"
      Height          =   195
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   3840
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login"
      Height          =   195
      Left            =   3840
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   5880
   End
End
Attribute VB_Name = "Dep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function PegarDepositos()
List1.Clear
On Error Resume Next
rst.Close
rst.CursorLocation = adUseClient
rst.Open "SELECT id, login FROM MuJB.dbo.Confirmacoes WHERE aprovado='" & Combo1.ListIndex & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

For i = 0 To rst.RecordCount
    List1.AddItem rst.Fields(0) & " - " & rst.Fields(1)
    rst.MoveNext
Next i

End Function
Private Function pegarDepInfo(idTMP As String)
On Error Resume Next
rst.Close
rst.CursorLocation = adUseClient
    rst.Open "SELECT id, login, valor, data, hora, banco, numero, anexo1, anexo2, anexo3, comentario, aprovado FROM MuJB.dbo.Confirmacoes WHERE id='" & idTMP & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText
    Text1.Text = rst.Fields(1)
    Text2.Text = rst.Fields(2)
    Text3.Text = rst.Fields(3)
    Text4.Text = rst.Fields(4)
    Text5.Text = rst.Fields(5)
    Text6.Text = rst.Fields(6)
    Text7.Text = rst.Fields(7)
    Text8.Text = rst.Fields(8)
    Text9.Text = rst.Fields(9)
    Text10.Text = rst.Fields(10)
    
    Call Mostrar
    
    Select Case rst.Fields(11)
    Case 0
        jcbutton3.Visible = False
        jcbutton2.Visible = True
    Case 1
        jcbutton3.Visible = True
        jcbutton2.Visible = False
    Case 2
        jcbutton3.Visible = False
        jcbutton2.Visible = True
    End Select
    
End Function

Private Sub Combo1_Click()
Call Sumir
    Call PegarDepositos
End Sub

Private Sub Form_Paint()
    Call PegarDepositos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst.Close
End Sub

Private Sub jcbutton1_Click()
    Unload Me
End Sub

Private Sub jcbutton2_Click()
    Dim Login As String
    Dim valor As String
    valor = rst.Fields(2)
    Login = rst.Fields(1)
    rst.Fields(11) = 1
    rst.Update
    Call adicionarbonus(valor, Login)
    Call PegarDepositos
    Call Sumir
End Sub

Private Function adicionarbonus(valor As String, tmplogin As String)
    valor = Replace(valor, ".", "")
    valor = Replace(valor, ",", "")
    valor = Replace(valor, "R", "")
    valor = Replace(valor, "$", "")
    Dim bonus As Integer
    bonus = CInt(valor)
    On Error Resume Next
    rst.Close
    MsgBox "Pagamento aprovado!" & vbNewLine & "Adicionado " & bonus & " bonus a conta!", vbInformation, "Sucesso!"
    
    rst.CursorLocation = adUseClient
    rst.Open "SELECT bonus FROM memb_info WHERE memb___id='" & tmplogin & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText
    rst.Fields(0) = rst.Fields(0) + bonus
    rst.Update
End Function

Private Sub jcbutton3_Click()
    rst.Fields(11) = 2
    rst.Update
    Call PegarDepositos
    Call Sumir
End Sub

Private Sub jcbutton4_Click()
If Text7.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text7.Text, vbNullString, vbNullString, 1
End Sub

Private Sub jcbutton5_Click()
If Text8.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text8.Text, vbNullString, vbNullString, 1
End Sub

Private Sub jcbutton6_Click()
If Text9.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text9.Text, vbNullString, vbNullString, 1
End Sub

Private Sub List1_Click()
    Dim tmp() As String
    tmp = Split(List1.Text, " - ")
    Call pegarDepInfo(tmp(0))
End Sub

Private Sub Mostrar()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
jcbutton2.Visible = True
jcbutton3.Visible = True
jcbutton4.Visible = True
jcbutton5.Visible = True
jcbutton6.Visible = True
End Sub

Private Sub Sumir()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
Text9.Visible = False
Text10.Visible = False
jcbutton2.Visible = False
jcbutton3.Visible = False
jcbutton4.Visible = False
jcbutton5.Visible = False
jcbutton6.Visible = False
End Sub

Private Sub Text7_click()
If Text7.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text7.Text, vbNullString, vbNullString, 1
End Sub

Private Sub Text8_Change()
If Text8.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text8.Text, vbNullString, vbNullString, 1
End Sub

Private Sub Text9_Change()
If Text9.Text = vbNullString Then Exit Sub
ShellExecute hWnd, vbNullString, AnexoLink & "\" & Text9.Text, vbNullString, vbNullString, 1
End Sub
