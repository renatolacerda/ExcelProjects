Attribute VB_Name = "fxExtrairCaracteres"
'Origem: http://guiadoexcel.com.br/funcao-para-extrair-somente-caracteres-de-celula-excel-vba
'Fun��o que retira somente o texto da c�lula
Public Function lfExtrairCaracteres(vPesquisa As Range) As String
    Dim lQtde As Long
 
    Application.Volatile
 
    'Recebe o valor da c�lula
    lfExtrairCaracteres = vPesquisa.Text
 
    'Retira os caracteres de 0 a 9, trocando-os por ""
    For lQtde = 0 To 9
        lfExtrairCaracteres = Replace(lfExtrairCaracteres, lQtde, "", 1)
    Next lQtde
 
End Function
