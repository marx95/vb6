VERSION 5.00
Begin VB.Form webshop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[MME] Webshop"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   Icon            =   "webshop.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   2055
   End
   Begin VB.Timer MudarInfo 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   6360
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Project1.jcbutton Atualizar 
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Caption         =   "Atualizar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Visivel?"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   7245
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   7245
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Burcas Item"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Burcas Categoria"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label info 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Categoria"
      Height          =   195
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "webshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CarregaCategorias()
List1.Clear
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT DISTINCT tipo FROM [MuJB].[dbo].[Webshop]", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

For i = 0 To rst.RecordCount
    List1.AddItem rst.Fields(0)
    rst.MoveNext
Next i
End Sub

Private Sub CarregaSubCategoria()
List2.Clear
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT nome FROM [MuJB].[dbo].[Webshop] WHERE tipo='" & List1.Text & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

For i = 0 To rst.RecordCount
    On Error Resume Next
    List2.Visible = True
    List2.AddItem rst.Fields(0)
    rst.MoveNext
Next i
End Sub

Private Sub CarregaItem()
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT nome, valor, visivel, tipo FROM [MuJB].[dbo].[Webshop] WHERE nome='" & List2.Text & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText
Text1.Text = rst.Fields(0) ' - nome
Text2.Text = rst.Fields(1) ' - valor
Text3.Text = rst.Fields(3) ' - categoria
Check1.Value = rst.Fields(2) ' - visivel
Me.Caption = "[MME] Webshop - " & Text1.Text
End Sub

Private Sub Atualizar_Click()
On Error Resume Next
rst.Close

rst.CursorLocation = adUseClient
rst.Open "SELECT nome, valor, visivel, tipo FROM [MuJB].[dbo].[Webshop] WHERE nome='" & List2.Text & "'", Cnn2, adOpenKeyset, adLockOptimistic, adCmdText

rst.Fields(0) = Text1.Text ' - nome
rst.Fields(1) = Text2.Text ' - valor
rst.Fields(2) = Check1.Value ' - visivel
rst.Fields(3) = Text3.Text ' - categoria
rst.Update
MudarInfo.Enabled = True
Call CarregaItem
End Sub

Private Sub Form_Paint()
    Call CarregaCategorias
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst.Close
End Sub

Private Sub jcbutton1_Click()
    Unload Me
End Sub

Private Sub List1_Click()
    Call CarregaSubCategoria
End Sub

Private Sub List2_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Check1.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Atualizar.Visible = True
CarregaItem
End Sub

Private Sub MudarInfo_Timer()
MudarInfo.Interval = 1200

Select Case info.Caption
Case ""
    info.Caption = "Item Atualizado!"
Case "Item Atualizado!"
    info.Caption = ""
    MudarInfo.Enabled = False
    MudarInfo.Interval = 1
End Select

End Sub

Private Sub Text4_Change()
    List1.Text = Text4.Text
End Sub
Private Sub Text5_Change()
    List2.Text = Text5.Text
End Sub
