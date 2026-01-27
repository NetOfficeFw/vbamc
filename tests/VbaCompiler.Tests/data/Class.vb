Public Property Name As String

Public Sub Initialize()
    Name = "John Doe"
End Sub

Public Function GetGreeting() As String
    GetGreeting = "Hello, " & Name & "!"
End Function
