Attribute VB_Name = "fxBackup"
' FUNÇÃO PARA CRIAR BACKUP QUANDO FECHAR A PLANILHA
' A QUANTIDADE MÁXIMA DE BACKUPS ESTÁ NO BEFORE_CLOSE DA PLANILHA
Public Function backup(contador) As Integer
    Dim vNome As String, strCaminho As String
    Dim thisFileName As String
    thisFileName = ThisWorkbook.name
    ' SE A PLANILHA FOR UM BACKUP NÃO CRIA UM BACKUP
    If Left(thisFileName, 3) <> "BCK" Then
        vNome = ActiveWorkbook.path & "\Backup" '& Format(Now(), "yyyy.mm.dd")
        strCaminho = Dir(vNome, vbDirectory)
        If (strCaminho = "") Then MkDir (vNome)
        ActiveWorkbook.SaveCopyAs vNome & "\BCK(" & Format(Now(), "yyyy.mm.dd") & ") (versao-" & contador & ")" & ActiveWorkbook.name
        'Application.DisplayAlerts = False
        'ActiveWorkbook.Save    'salva o arquivo ativo
        ActiveWorkbook.Close SaveChanges:=True
        backup = contador + 1
    End If
End Function
