VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ultimo As Variant
    ultimo = ReadINI(App.Path & "/config.ini", "face", "ultimo")
    
    For i = ultimo To 100001000000000#
        Call WriteINI(App.Path & "/config.ini", "face", "ultimo", "" & i)
        Call DownloadAFile(App.Path & "/download/" & i & ".ini", "http://graph.facebook.com/" & i, False, False)
        DoEvents
        Command1.Caption = i
    Next i
End Sub

Private Sub Form_Load()
    Dim ultimo As Variant
    ultimo = ReadINI(App.Path & "/config.ini", "face", "ultimo")
    Command1.Caption = ultimo
End Sub
