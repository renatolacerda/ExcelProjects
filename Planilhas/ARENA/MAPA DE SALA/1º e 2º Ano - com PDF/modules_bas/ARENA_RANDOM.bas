Attribute VB_Name = "ARENA_RANDOM"
Sub RANDOM_ALUNOS()
Application.ScreenUpdating = False
    Worksheets("BD").Select
    Range("D:E") = ""
    valor = Range("B60000").End(xlUp).Row
    ReDim Var(0 To valor) As Integer
    For x = 1 To valor
k:
        Z = Int((valor * Rnd) + 1)
        If Var(Z) = Z Then
            achou = True
        Else
            achou = False
        End If
        If achou = True Then
            GoTo k
        Else
            Range("D" & x) = Z
            Var(Range("D" & x)) = Z
        End If
c:
    Next
    ' ORDENA EM ORDEM CRESCENTE
    Range("A1:D" & Range("A60000").End(xlUp).Row).Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Worksheets(1).Select
Application.ScreenUpdating = True
End Sub
