' ThisDocument file

Public Sub ShowMessage()
    MsgBox "Hello world from " & Application.Name & ", PID " & CStr(GetCurrentProcessId)
End Sub
