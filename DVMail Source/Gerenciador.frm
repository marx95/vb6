VERSION 5.00
Begin VB.Form Gerenciador 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerenciador de email's - MarxD Software"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lista 
      Height          =   6135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox mail_txt 
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Adicionar email:"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Gerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
lista.Text = lista.Text + vbNewLine + mail_txt.Text
Open App.Path & "\Emails.txt" For Output As #3
    Print #3, lista.Text
Close #3
mail_txt.Text = ""
End Sub

Private Sub Form_Load()
Open App.Path & "\Emails.txt" For Input As #1
   lista.Text = Input(FileLen(App.Path & "\Emails.txt"), #1)
Close #1
End Sub
