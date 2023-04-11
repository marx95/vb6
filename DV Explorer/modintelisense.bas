Attribute VB_Name = "modintelisense"
'--------------------------------------------------------------------------------
' IntelliSense for VB the TextBox Control.
' Danny Young
' dan@mydan.com
'
' This is, in basic terms, the auto-complete function used in IE4 / IE5
'
' This module doesn't have to be modified in any way.  The only thing you may
' want to edit is the directory you store the keyword file in.  Right now it
' stores it in the application's directory under the subdirectory "IntelliSense"
'
' There is one other thing you may want to change, if you think that the text
' entered into the TextBox is going to be greater than 50 Characters then change
' this amount accordingly in the Type iSense. * remember to delete the original
' file to complete the change *
'
' It creates a new file for every TextBox you have under it's name. So if you
' have a TextBox called "txtName" then it will create a file called "txtname.dat"
'
' It automatically adds new keywords to the data file when the user presses the
' Enter key.
'
' Syntax
' ------
' In the CHANGE event of your TextBox, put in the following line
'     iSenseChange YourTextBoxName
' Where "YourTextBoxName" is the name of the TextBox associated with the event.
'
' In the KEYPRESS event of your TextBox, put in the following line
'     iSenseKeyPress YourTextBoxName, KeyAscii
' Where "YourTextBoxName" is the name of the TextBox associated with the event.
'
' That's it really, included with this package is the example Form to show you
' how it works and to show you the proper syntax.
'
' Any ideas or changes, let me know!
'--------------------------------------------------------------------------------
Option Explicit

Global WasDelete As Boolean

Public Type iSense
    sOut As String * 50
End Type

Public Function IntelliSense(tBox As TextBox, AddRecord As Boolean) As String
    Dim iChannel As Integer, iActive As Integer, iLength As Integer, i As Integer
    Dim iFile As String
    Dim iSense As iSense
    Dim Done As Boolean
    
    iFile = "captcha" & ".DVE"
    iLength = Len(iSense)
    iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
    Close iChannel
    
    iActive = FileLen(iFile) / iLength
    iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
        If AddRecord Then
            iSense.sOut = tBox.Text
            Put iChannel, iActive + 1, iSense
        Else
            Do While Not EOF(iChannel) And Done = False
                i = i + 1
                Get iChannel, i, iSense
                If tBox.Text = Mid(RTrim(iSense.sOut), 1, Len(tBox.Text)) Then
                    IntelliSense = RTrim(iSense.sOut)
                End If
            Loop
        End If
    Close iChannel
End Function

Public Sub iSenseChange(tBox As TextBox)
    Dim iStart As Integer
    Dim iSense As String
    
    iStart = tBox.SelStart
    iSense = IntelliSense(tBox, False)
    If iSense <> "" And Not WasDelete Then
        tBox.Text = iSense
        tBox.SelStart = iStart
        tBox.SelLength = Len(tBox.Text) - iStart
    End If
End Sub

Public Sub iSenseKeyPress(tBox As TextBox, KeyAscii As Integer)
    If KeyAscii = 13 And tBox.Text <> "" Then
        IntelliSense tBox, True
    ElseIf KeyAscii = 8 Then
        WasDelete = True
    Else
        WasDelete = False
    End If
End Sub

