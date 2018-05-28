Attribute VB_Name = "TESTE_PDF"
Sub TestPDFThumbnailGeneration()
'Print the document to PDF
    ActiveWorkbook.PrintOut copies:=1, ActivePrinter:="PDFCreator"

    'Wait until the print job has entered the print queue
    Do Until OutputJob.cCountOfPrintjobs = 1
        DoEvents
    Loop
    OutputJob.cPrinterStop = False

    'Wait until PDF creator is finished then release the objects
    Do Until OutputJob.cCountOfPrintjobs = 0
        DoEvents
    Loop

    OutputJob.cClose
    Set OutputJob = Nothing
End Sub


Sub MYTESTE()
PrintToPDF_Early ("TESTE1")
End Sub
