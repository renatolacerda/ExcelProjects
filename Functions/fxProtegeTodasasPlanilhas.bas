Attribute VB_Name = "fxProtegeTodasasPlanilhas"
'Fun��o que protege todas as planilhas de um arquivo
Sub lsProtegerTodasAsPlanilhas()
    'Declara as vari�veis necess�rias
    Dim lPass As String
    Dim lQtdePlan As Integer
    Dim lPlanAtual As Integer
 
    'Solicita a senha
    'O m�todo InputBox � utilizado para solicitar um valor atrav�s de um formul�rio
    lPass = InputBox("Proteger todas as planilhas:", "Senha", ActName)
 
    'Inicia as vari�veis
    'O m�todo Worksheets.Count passa a quantidade de planilhas existentes no arquivo
    lQtdePlan = Worksheets.Count
    lPlanAtual = 1
 
    'Loop pelas planilhas
    'A fun��o While realiza um loop de c�digo enquanto n�o passar por todas as planilhas contadas
    While lPlanAtual <= lQtdePlan
        'O m�todo Worksheets(lPlanAtual).Activate ativa a planilha conforme o �ndice atual 1, 2, 3...
        Worksheets(lPlanAtual).Activate
 
        'O m�todo .Protect proteje a planilha passando os par�metros para proteger
        'objetos de desenho, conte�do, cen�rios e passando o password digitado
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lPass
 
        'Muda o �ndice para passar para a pr�xima planilha
        lPlanAtual = lPlanAtual + 1
    Wend
 
    'O m�todo MsgBox exibe um formul�rio de aviso ao usu�rio.
    MsgBox "Planilhas protegidas!"
 
End Sub
'Fun��o que desprotege todas as planilhas de um arquivo
Sub lsDesprotegerTodasAsPlanilhas()
    'Declara as vari�veis necess�rias
    Dim lPass As String
    Dim lQtdePlan As Integer
    Dim lPlanAtual As Integer
 
    'Solicita a senha
    'O m�todo InputBox � utilizado para solicitar um valor atrav�s de um formul�rio
    lPass = InputBox("Desproteger todas as planilhas:", "Senha", ActName)
 
    'Inicia as vari�veis
    'O m�todo Worksheets.Count passa a quantidade de planilhas existentes no arquivo
    lQtdePlan = Worksheets.Count
    lPlanAtual = 1
 
    'Loop pelas planilhas
    'A fun��o While realiza um loop de c�digo enquanto n�o passar por todas as planilhas contadas
    While lPlanAtual <= lQtdePlan
        'O m�todo Worksheets(lPlanAtual).Activate ativa a planilha conforme o �ndice atual 1, 2, 3...
        Worksheets(lPlanAtual).Activate
 
        'O m�todo .UnProtect desprotege a planilha
        ActiveSheet.Unprotect Password:=lPass
 
        'Muda o �ndice para passar para a pr�xima planilha
        lPlanAtual = lPlanAtual + 1
    Wend
 
    'O m�todo MsgBox exibe um formul�rio de aviso ao usu�rio.
    MsgBox "Planilhas desprotegidas!"
 
End Sub
