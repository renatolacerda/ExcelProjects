Attribute VB_Name = "ARENA_QTDSALAS"
Dim P_PRINCIPAL As Worksheet
Dim P_SHEET As Worksheet
Sub CONTA_QTD_SALAS()
Set P_PRINCIPAL = Sheets("BD")
Set P_SHEET = Sheets("QTD POR SALA")
P_SHEET.Range("C3:P65000").ClearContents
For L = 1 To P_PRINCIPAL.Range("A65000").End(xlUp).Row
    sala = P_PRINCIPAL.Range("E" & L)
    TURMA = P_PRINCIPAL.Range("C" & L)
    'ACHA LINHA
    For ll = 3 To P_SHEET.Range("B65000").End(xlUp).Row
        If P_SHEET.Range("B" & ll) = sala Then
            linha = ll
            Exit For
        End If
    Next
    'ACHA COLUNA
    For CC = 3 To 16 'P_SHEET.Range("IV2").End(xlToLeft).Column
        If P_SHEET.Cells(2, CC) = TURMA Then
            COLUNA = CC
            Exit For
        End If
    Next
    'SOMA
    P_SHEET.Cells(linha, COLUNA) = P_SHEET.Cells(linha, COLUNA) + 1
    Next
End Sub
