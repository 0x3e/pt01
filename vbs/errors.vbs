'On Error Resume Next
On Error Goto 0
Sub ShowErr()
    WScript.Echo "Err.Number: " & Err.Number
    WScript.Echo Err.Description
    WScript.Echo("Error Source: " & Err.Source)
    On Error Goto 0
    WScript.Quit 1
End Sub
