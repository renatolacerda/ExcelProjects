Attribute VB_Name = "MOD_EMAILS_IGUAIS_ENTRE2ABAS"
Sub LOCALIZAR_IGUAIS()
Dim COMPRAS As Worksheet
Dim ENVIADOS As Worksheet

Dim LCOMPRAS, LENVIADOS

Set COMPRAS = Sheets("COMPRAS")
Set ENVIADOS = Sheets("ENVIADOS")
For LCOMPRAS = 2 To COMPRAS.Range("C65000").End(xlUp).Row
    For LENVIADOS = 1 To ENVIADOS.Range("A65000").End(xlUp).Row
        If LCase(COMPRAS.Cells(LCOMPRAS, 3)) = LCase(ENVIADOS.Cells(LENVIADOS, 1)) Then
            COMPRAS.Cells(LCOMPRAS, 3).Select
            Call PINTA
        End If
    Next
Next
End Sub
Sub PINTA()
'
' VERDE Macro
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
