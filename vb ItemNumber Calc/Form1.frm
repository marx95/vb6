VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Number Calc"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "1.02++"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "97D ~ 1.0"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Index"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grupo"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Option1.Value = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Calc
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Calc
End Sub

Public Sub Calc()
    Dim Mult As String
    Mult = 32
    If Option2.Value Then Mult = 512
    On Error Resume Next
    Label1.Caption = (CInt(Text1.Text) * Mult) + CInt(Text2.Text)
End Sub
