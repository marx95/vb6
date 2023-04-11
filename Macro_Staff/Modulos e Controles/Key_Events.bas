Attribute VB_Name = "Key_Events"
Public Function Control_C()
    SendKeys "^{C}"
    DoEvents
End Function

Public Function Control_V()
    SendKeys "^{V}"
    DoEvents
End Function

Public Function Enter()
    Call SendKeys("{ENTER}")
    DoEvents
End Function
