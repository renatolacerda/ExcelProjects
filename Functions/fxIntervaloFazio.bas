Attribute VB_Name = "fxIntervaloFazio"
Function �Fazio(intervalo As Range) As String
If WorksheetFunction.CountA(intervalo) = 0 Then
    �Fazio = "Empty"
Else
    �Fazio = "Not Empty"
End If
End Function
