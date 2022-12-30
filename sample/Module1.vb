' Module source code
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Public Sub Auto_Open()
    ' execute code when application starts
    MsgBox "Hello world from application " & Application.Name & " PID " & CStr(GetCurrentProcessId)
End Sub
