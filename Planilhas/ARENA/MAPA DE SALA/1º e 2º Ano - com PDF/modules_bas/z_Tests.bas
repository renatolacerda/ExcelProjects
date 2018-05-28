Attribute VB_Name = "z_Tests"
' FROM: http://www.vbaexpress.com/forum/showthread.php?44168-PDFCreator-automation-in-VBA-Excel

'Version 1.0
'Andreas Killer
'17.01.16
'Print with PDFCreator 2.0 and above
Option Explicit

Sub PDF(nomeArquivo As String)
  'Initialize PDFCreator
  If Not PDFCreatorPrint() Then
    Debug.Print "Can not initialize PDFCreator"
    Stop
  End If
  
  'Print as usual
  ActiveSheet.PrintOut
  'Another sheet (or whatever as you like)
  'Sheets(2).PrintOut
  
  'Finish all into this file
  Dim path
  path = ThisWorkbook.path
  If Not PDFCreatorPrint(path & "\" & nomeArquivo) Then
    Debug.Print "Can not create PDF file"
    Stop
  End If
End Sub

Function PDFCreatorPrint(Optional ByVal PDFname, _
    Optional ByVal JobCount As Long) As Boolean
  Static pcPrinter As Variant 'PDFCreator.Printers
  Static pcQueue As Object 'PDFCreator.Queue
  Dim pcPrintJob As Object 'PDFCreator.PrintJob
  If IsMissing(PDFname) Then
    'Be sure a PDFCreator printer is the active printer
    If IsEmpty(pcPrinter) Then
      For Each pcPrinter In PDFCreatorPrinters
        If InStr(1, ActivePrinter, pcPrinter, vbTextCompare) > 0 Then Exit For
      Next
      If IsEmpty(pcPrinter) Then
        Err.Raise 68 'PDFCreator is not the active printer
      End If
    End If
    On Error GoTo ExitPoint
    'Initialize the job queue if necessary
    If pcQueue Is Nothing Then
      Set pcQueue = CreateObject("PDFCreator.JobQueue")
      pcQueue.Initialize
    End If
  Else
    'Wait a second, maybe jobs are pending
    If JobCount <= 0 Then JobCount = 9999
    pcQueue.WaitForJobs JobCount, 1
    'Got one?
    If pcQueue.count = 0 Then Exit Function
    'Merge all into one file
    pcQueue.MergeAllJobs
    Set pcPrintJob = pcQueue.GetJobByIndex(0)
    pcPrintJob.ConvertToAsync PDFname
    'Finish
    pcQueue.ReleaseCom
    Set pcQueue = Nothing
    pcPrinter = Empty
  End If
  PDFCreatorPrint = True
ExitPoint:
End Function

Private Function PDFCreatorPrinters() As Collection
  'Returns a collection of all PDFCreator printers
  Dim pcObj As Object 'PDFCreator.PdfCreatorObj
  Dim pcPrinter As Object 'PDFCreator.Printers
  Dim i As Long
  Set PDFCreatorPrinters = New Collection
  On Error GoTo ExitPoint
  Set pcObj = CreateObject("PDFCreator.PdfCreatorObj")
  Set pcPrinter = pcObj.GetPDFCreatorPrinters
  For i = 0 To pcPrinter.count - 1
    PDFCreatorPrinters.Add pcPrinter.GetPrinterByIndex(i)
  Next
ExitPoint:
End Function

Private Sub SetRef()
  On Error Resume Next
  'PDFCreator - Your OpenSource PDF Solution
  Application.VBE.ActiveVBProject.References.AddFromGuid _
    "{8B8D2928-EAAF-492D-8DA5-E06B358D8826}", 2, 0
End Sub

