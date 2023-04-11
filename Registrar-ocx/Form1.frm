VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
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
' Un TextBox para especificar la ruta y nombre del OCX/DLL deseados
' Verifica que el archivo deseado exista y su ruta este correcta
' Command1 para registrar el componente
' Command2 para desregistrar el componente
Private Sub Form_Load()
 Text1.Text = "C:\WINDOWS\SYSTEM32\COMCTL32.OCX"
 Command1.Caption = "Registrar"
 Command2.Caption = "DesRegistrar"
End Sub
Private Sub Command1_Click()
 If RegistrarComponente(Me.hWnd, Text1.Text, True) Then
 MsgBox "El componente se registro exitosamente"
 Else
 MsgBox "Ha fallado el registro del componente"
 End If
End Sub
Private Sub Command2_Click()
 If RegistrarComponente(Me.hWnd, Text1.Text, False) Then
 MsgBox "El componente se des-registro exitosamente"
 Else
 MsgBox "Ha fallado el des-registro del componente"
 End If
End Sub
