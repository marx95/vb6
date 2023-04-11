VERSION 5.00
Begin VB.Form captcha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Digite o Captcha"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "captcha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Sfocus 
      Interval        =   500
      Left            =   4200
      Top             =   1560
   End
   Begin VB.TextBox cap_txt 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   400
      TabIndex        =   0
      Top             =   2160
      Width           =   3895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pressione Espaço para atualizar a imagem"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "captcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1

'Private Sub cap_txt_KeyDown(KeyCode As Integer, Shift As Integer)

'End Sub
Private Sub cap_txt_change()
iSenseChange cap_txt
End Sub
Private Sub cap_txt_KeyPress(KeyAscii As Integer)
On Error Resume Next
iSenseKeyPress cap_txt, KeyAscii
If KeyAscii = 13 Then
Call Confirma
End If

If KeyAscii = 32 Then
m_WebControl.object.Refresh
cap_txt.Text = ""
End If
End Sub

Private Sub Form_unload(Cancel As Integer)
Principal.Timer1.Enabled = True
End Sub
Private Sub Form_Load()
Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", captcha)
  m_WebControl.Move 400, 380, 3895, 1635
  m_WebControl.Visible = True
  m_WebControl.object.navigate "http://www.orkut.com.br/CaptchaImage"
  m_WebControl.object.Silent = True
End Sub
Private Sub Form_Paint()
On Error Resume Next
  cap_txt.SetFocus
End Sub

Private Sub Confirma()
Principal.Timer1.Enabled = True
On Error Resume Next
Principal.m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.getelementbyid("captchaTextbox").Value = cap_txt.Text
Principal.m_WebControl.object.Document.getelementbyid("orkutFrame").contentWindow.Document.body.Document.getelementbyid("mboxfull").childnodes.Item(0).childnodes.Item(0).childnodes.Item(1).childnodes.Item(0).childnodes.Item(0).childnodes.Item(17).childnodes.Item(0).childnodes.Item(0).Click
TotalCaptchaSend = TotalCaptchaSend + 1
If TotalCaptchaSend = 1 Then
Principal.Status.Panels(2).Text = TotalCaptchaSend & " captcha digitado"
Else
Principal.Status.Panels(2).Text = TotalCaptchaSend & " captchas digitado"
End If
Unload Me
End Sub

Private Sub Sfocus_Timer()
On Error Resume Next
'cap_txt.SetFocus
End Sub
