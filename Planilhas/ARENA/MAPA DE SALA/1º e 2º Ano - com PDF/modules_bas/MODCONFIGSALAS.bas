Attribute VB_Name = "MODCONFIGSALAS"
Public Function SALAS(sala)
Dim Q As New Worksheet
Set Q = Sheets("CONFIG-SALAS")
Dim L

For L = 2 To UltimaLinha(Q, 2)
    If sala = Q.Cells(L, 1) Then
        SALAS = Q.Cells(L, 2): Exit For
    End If
Next
End Function
