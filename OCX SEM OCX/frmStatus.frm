VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStatus 
   Caption         =   "MMM - Make My Manifest"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3060
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5445
      TabIndex        =   1
      Top             =   4725
      Width           =   915
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   225
      Width           =   6375
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'frmStatus
'=========
'
'This form is shown non-modally by modMain as a display
'dialog and a host for a CommonDialog control.  For the
'most part it is operated remotely from modMain via a
'number of public properties.
'

'================= Public procedures =====================

Public Sub Done()
    cmdClose.Enabled = True
End Sub

Public Function GetProjectFile() As String
    With dlgFile
        .CancelError = True
        .DialogTitle = "Select project file"
        .Filter = "VB project files (*.vbp)|*.vbp"
        .Flags = cdlOFNExplorer _
              Or cdlOFNFileMustExist _
              Or cdlOFNHideReadOnly _
              Or cdlOFNLongNames _
              Or cdlOFNPathMustExist _
              Or cdlOFNShareAware
        On Error Resume Next
        .ShowOpen
        If Err.Number <> 0 Then
            GetProjectFile = ""
        Else
            GetProjectFile = .FileName
        End If
    End With
End Function

Public Sub Log(Optional ByVal Text As String = "")
    With txtLog
        .Text = .Text & vbNewLine & Text
        .SelStart = Len(.Text)
    End With
    DoEvents
End Sub

'================== Misc procedures ======================

Private Sub Resizer()
    Static blnLoaded As Boolean
    Static sngLogHeightDelta As Single
    Static sngLogWidthDelta As Single
    Static sngCloseTopDelta As Single
    Static sngCloseLeftDelta As Single
    
    If blnLoaded Then
        If Height < 3500 Then Height = 3500
        If Width < 4000 Then Width = 4000
        txtLog.Height = Height - sngLogHeightDelta
        txtLog.Width = Width - sngLogWidthDelta
        cmdClose.Top = Height - sngCloseTopDelta
        cmdClose.Left = Width - sngCloseLeftDelta
    Else
        blnLoaded = True
        sngLogHeightDelta = Height - txtLog.Height
        sngLogWidthDelta = Width - txtLog.Width
        sngCloseTopDelta = Height - cmdClose.Top
        sngCloseLeftDelta = Width - cmdClose.Left
    End If
End Sub

'=================== Event handlers ======================

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Resizer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not cmdClose.Enabled Then Cancel = 1
End Sub

Private Sub Form_Resize()
    Resizer
End Sub
