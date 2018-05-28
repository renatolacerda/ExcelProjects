Attribute VB_Name = "MOD_PDF"

Sub IMPRESSAO_PDF_SALA()
Dim c As Worksheet
Dim r As Worksheet
Set c = Sheets("CONFIG")
Set r = Sheets("Rel-Turma")

For x = 3 To UltimaLinha(c, 3)

Application.ScreenUpdating = False
    
    CMB_TURMAS = Sheets("CONFIG").Cells(x, 3)
    
    Call FRM_RELATORIO.GERA_REL(ActiveSheet.name, CMB_TURMAS, OPCAO_DE_MAPA)
    
    FRM_RELATORIO.WORDART (CMB_TURMAS)
    ORDENA_SALA_RELATORIO_1
    FORMATAR_LISTA_SALA
    AREA_IMPRESSAO_SALA
    
    PRETO_SALA
 
    ' Impressao turma
    Range("a1").Select
    'MsgBox Sheets("CONFIG").Cells(x, 3)
    myname = ActiveWorkbook.name & "-" & CMB_TURMAS & "-" & Sheets("CONFIG").Range("F2") & "-" & Sheets("CONFIG").Cells(x, 3) & ".pdf"
    myname = Replace(myname, "/", "-")
    myname = Replace(myname, ".xls", "")

Application.ScreenUpdating = True

    ' Impressao sala
    'myname = ActiveWorkbook.name & "- MAPA -" & CMB_TURMAS & "-" & Sheets("CONFIG").Range("F2") & "-" & Sheets("CONFIG").Range("F4") & ".pdf"
    myname = Sheets("CONFIG").Range("F2") & "-" & Sheets("CONFIG").Range("F4") & ".pdf"
    myname = Replace(myname, "/", "-")
    'myname = Replace(myname, ".xls", "")
    
    Application.ScreenUpdating = False
    
    Sheets(CMB_TURMAS).Activate
    
    'PDFPRINT (myname)
    
    'PrintToPDF_Early (myname)
    PDF (myname)
    
    Sheets("Rel-Sala").Activate
    
    Application.ScreenUpdating = True
Next
End Sub
Sub IMPRESSAO_PDF_TURMA()
Dim c As Worksheet
Dim r As Worksheet
Dim TURMAS
Set c = Sheets("CONFIG")
Set r = Sheets("Rel-Turma")

'For x = 2 To UltimaLinha(C, 9)

arrTurmas = Sheets("Config").Cells(3, 1)
TURMAS = Split(arrTurmas, ";")

For Each t In TURMAS

Application.ScreenUpdating = False
    
    CMB_TURMAS = t
    
    Call FRM_RELATORIO.GERA_REL(ActiveSheet.name, CMB_TURMAS)
    FRM_RELATORIO.WORDART (x - 2)
    If Range("A1") = 2 Then
        ORDENA_TURMA_RELATORIO_1
        AREA_IMPRESSAO_TURMA2
    ElseIf Range("A1") = 1 Then
        ORDENA_TURMA_RELATORIO_2
        AREA_IMPRESSAO_TURMA1
    End If
    FORMATAR_LISTA_TURMA
    
    Call PRETO_TURMA
            
    Range("a1").Select
    myname = ActiveWorkbook.name & "-" & CMB_TURMAS & "-" & Sheets("CONFIG").Range("F2") & "-" & Sheets("CONFIG").Range("F4") & ".pdf"
    myname = Replace(myname, "/", "-")
    'myname = Replace(myname, ".xls", "")

Application.ScreenUpdating = True

    PrintToPDF_Early (myname)
    
Next
End Sub
Sub IMPRESSAO_PDF_MAPA()
'
' TESTE Macro
'

'
    Sheets(Array("Sala 12", "Sala 11", "Sala 10", "Sala 9", "Sala 8", "Sala 7", "Sala 6", _
        "Sala 5", "Sala 4", "Sala 3", "Sala 2", "Sala 1")).Select
    Sheets("Sala 1").Activate
    
    PrintToPDF_Early (MAPA)
    
    'Application.ActivePrinter = "PDFCreator em Ne00:"
    'ExecuteExcel4Macro "PRINT(1,,,1,,,,,,,,2,""PDFCreator em Ne00:"",,TRUE,,FALSE)"
End Sub


