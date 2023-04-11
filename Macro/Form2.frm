VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SuperBot do MuNovus - Ajuda"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin MacroMuNovus.isButton isButton1 
      Height          =   735
      Left            =   7800
      TabIndex        =   1
      Top             =   6085
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      Icon            =   "Form2.frx":08CA
      Style           =   6
      Caption         =   "Fechar"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Ajuda 
      Height          =   5850
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      _cx             =   16933
      _cy             =   10319
      FlashVars       =   ""
      Movie           =   "http://www.youtube.com/v/ZA2Kr9vf8k0?version=3&amp;hl=pt_BR"
      Src             =   "http://www.youtube.com/v/ZA2Kr9vf8k0?version=3&amp;hl=pt_BR"
      WMode           =   "Window"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   "LT"
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "NoScale"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   "5850"
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub isButton1_Click()
    Unload Me
End Sub
