Attribute VB_Name = "MOD_FORMATAR"
Sub FORMATAR_LISTA_TURMA()
' =-=-=-=-=-=- ((LISTA-TURMA01)) =-=-=-=-=-=-
Range("AE1:AI2").Copy
Range("B13:F" & Range("E65000").End(xlUp).Row).PasteSpecial xlPasteFormats
' =-=-=-=-=-=- ((LISTA-TURMA02)) =-=-=-=-=-=-
Range("AA1:AC2").Copy
Range("H13:J" & Range("J65000").End(xlUp).Row).PasteSpecial xlPasteFormats
Range("A1").Select
End Sub
Sub FORMATAR_LISTA_SALA()
' =-=-=-=-=-=- ((LISTA-TURMA01)) =-=-=-=-=-=-
Range("AA1:AD2").Copy
Range("B13:E" & Range("D65000").End(xlUp).Row).PasteSpecial xlPasteFormats
Range("AF1:AH1").Copy
Range("C12:E12").PasteSpecial xlPasteFormats
End Sub
