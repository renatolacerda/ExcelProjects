Attribute VB_Name = "modListaArquivos"
Public Function fxListaArquivos(ByVal Caminho As String) As String()
    'Aten��o: Fa�a refer�ncia � biblioteca Micrsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim result() As String
    Dim Pasta As Folder
    Dim Arquivo As File
    Dim Indice As Long
 
 
    ReDim result(0) As String
    If FSO.FolderExists(Caminho) Then
        Set Pasta = FSO.GetFolder(Caminho)
 
        For Each Arquivo In Pasta.Files
            Indice = IIf(result(0) = "", 0, Indice + 1)
            ReDim Preserve result(Indice) As String
            result(Indice) = Arquivo.Name
        Next
    End If
 
    fxListaArquivos = result
ErrHandler:
    Set FSO = Nothing
    Set Pasta = Nothing
    Set Arquivo = Nothing
End Function
'Importante: Fa�a refer�ncia � biblioteca Micrsoft Scripting Runtime
'para ter acesso aos objetos da File System Object (FSO), necess�rios para execu��o do exemplo.
Private Sub ListaArquivos()
Dim arq As String
    Dim arquivos() As String
    Dim lCtr As Long
    Caminho = ThisWorkbook.Path
    'Debug.Print Caminho
    'MsgBox Caminho
    arquivos = fxListaArquivos(Caminho)
    For lCtr = 0 To UBound(arquivos)
      Debug.Print arquivos(lCtr)
      arq = arquivos(lCtr)
      Sheets("ARQUIVOS DA PASTA").Cells(lCtr + 1, 1) = arq
    Next
End Sub
