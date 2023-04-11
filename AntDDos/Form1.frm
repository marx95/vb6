VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form AntDDoS 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Text            =   "1"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5295
      ItemData        =   "Form1.frx":0000
      Left            =   4920
      List            =   "Form1.frx":0007
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Upload 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Download 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8400
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock SockB 
      Index           =   0
      Left            =   9600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   9120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox LogTxt 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upload:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Download:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "AntDDoS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function LogAdd(Log As String)
    If Len(LogTxt.Text) = 0 Then
        LogTxt.Text = Log
        Exit Function
    End If
    
    LogTxt.Text = LogTxt.Text + vbNewLine + Log
End Function

Private Sub Command1_Click()
    MsgBox Sock(CInt(Text1.Text)).State
End Sub

Private Sub Form_Load()
    List1.AddItem "82"
    Sock(0).LocalPort = 82
    Sock(0).Listen
    C_Min = 1
    C_Max = 1000

    For i = C_Min To C_Max
        Load Sock(CInt(i))
        Load SockB(CInt(i))
    Next i
End Sub

Private Sub Sock_Close(Index As Integer)
    On Error Resume Next
    SockB(Index).Close
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    For i = C_Min To C_Max
        If Sock(CInt(i)).State = 0 Then
            Sock(CInt(i)).Accept requestID
            
            If SockB(CInt(i)).State <> 0 Then
                On Error Resume Next
                SockB(CInt(i)).Close
            End If
            SockB(CInt(i)).Connect "munovus.net", 80
            Exit Sub
        End If
    Next i
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String
    Sock(Index).GetData Data

    On Error GoTo ErroA
    SockB(Index).SendData Data
    Trafego(0) = Trafego(0) + Len(Data)
    Call LogAdd("Sock ID(" & Index & ") Enviado: " & Len(Data))
Exit Sub
ErroA:
End Sub

Private Sub Sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    SockB(Index).Close
End Sub

Private Sub SockB_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String
    SockB(Index).GetData Data
    
    On Error GoTo ErroA
    Sock(Index).SendData Data
    Trafego(1) = Trafego(1) + Len(Data)
    Call LogAdd("SockB ID(" & Index & ") Enviado: " & Len(Data))
Exit Sub
ErroA:
End Sub

Private Sub Timer1_Timer()
    Download.Text = SetBytes(Trafego(0))
    Upload.Text = SetBytes(Trafego(1))
    
    
    For i = C_Min To C_Max
        If Sock(CInt(i)).State <> 7 Then
            If Sock(CInt(i)).State > 0 Then
                On Error Resume Next
                SockB(CInt(i)).Close
            End If
        End If
    Next i
End Sub
