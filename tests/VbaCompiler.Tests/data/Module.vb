Attribute VB_Name = "Module1"
Option Explicit

' Sample module for VBA compiler testing
Public Sub HelloWorld()
    Debug.Print "Hello, World!"
End Sub

Public Function Add(ByVal a As Integer, ByVal b As Integer) As Integer
    Add = a + b
End Function

Public Function Multiply(ByVal x As Double, ByVal y As Double) As Double
    Multiply = x * y
End Function

Private Sub InternalMethod()
    ' This is a private helper method
    Debug.Print "Internal method called"
End Sub
