VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox comunidades 
      Height          =   2205
      ItemData        =   "Form1.frx":0000
      Left            =   11040
      List            =   "Form1.frx":0010
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ja estou logado"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Source 
      Height          =   8055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   13215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

Private Sub Command1_Click()

If m_WebControl.Visible = True Then
m_WebControl.Visible = False
Source.Visible = True
Else
Source.Visible = False
m_WebControl.Visible = True
End If
End Sub

Private Sub Command2_Click()
m_WebControl.object.getelementbyid("navbar_username").Value = "marx951"
m_WebControl.object.getelementbyid("navbar_password").Value = "xaubet951"
m_WebControl.object.getelementbyid("navbar_loginform").submit
End Sub

Private Sub Form_Load()
Logado = 0
QualEstou = 1
Email = "mujb_dv@live.com"
Senha = "mujbserver"
Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Form1)
  m_WebControl.Move 120, 720, 13215, 8055
  m_WebControl.Visible = True
  m_WebControl.object.Navigate "http://www.webcheats.com.br/forum"
  m_WebControl.object.silent = True
End Sub

Private Sub m_WebControl_DocumentComplete(ByVal pDisp As Object, URL As Variant)


End Sub

Private Sub Timer1_Timer()
If Logado = 0 Then
On Error Resume Next
Source.Text = m_WebControl.object.Document.documentelement.innerhtml
End If


If (InStr(1, Source.Text, "navbar_username")) Then
m_WebControl.object.getelementbyid("navbar_username").Value = "marx951"
m_WebControl.object.getelementbyid("navbar_password").Value = "xaubet951"
m_WebControl.object.getelementbyid("navbar_loginform").submit
End If
End Sub
