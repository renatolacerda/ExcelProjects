Attribute VB_Name = "ARENA_PDFPRINT"
Sub PDFPRINT(name As String)
'(ByVal sDocumentToConvert As String, ByVal sValue As String, ByVal sNewFolder As String)
Dim pdfjob As PDFCreator_COM.PdfCreatorObj 'PDFCreator.clsPDFCreator
Dim p

p = ActivePrinter

WordApp.ActivePrinter = "PDFCreator"
Set pdfjob = New PDFCreator_COM.PdfCreatorObj 'PDFCreator.clsPDFCreator

' SET VARIABLES
sNewFolder = ThisWorkbook.path
sValue = name
sValue = Replace(".xls", sValue, "")


With pdfjob
.cStart "/NoProcessingAtStartup"
.cOption("UseAutosave") = 1
.cOption("UseAutosaveDirectory") = 1
.cOption("AutosaveDirectory") = sNewFolder
.cOption("AutosaveFilename") = sValue
.cOption("AutosaveFormat") = 0 ' 0 = PDF
.cClearCache
End With

'Print the document to PDF
wrdDoc.PrintOut
'Wait until the print job has entered the print queue
Do Until pdfjob.cCountOfPrintjobs = 1
DoEvents
Loop
pdfjob.cPrinterStop = False
'Wait until PDF creator is finished then release the objects
Do Until pdfjob.cCountOfPrintjobs = 0
DoEvents
Loop
pdfjob.cClose
Set pdfjob = Nothing
ActivePrinter = p
End Sub
