Attribute VB_Name = "fxArrendodamentos"

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-= ARRENDODAMENTOS =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'  AsymDown      Arredonda assimetricamente os n�meros para menos - semelhante a Int().
'                 Os n�meros negativos ficam mais negativos.
'  SymDown       Arredonda simetricamente os n�meros para menos - semelhante a Fix().
'                 Trunca todos os n�meros para 0.
'                 Igual a AsymDown para n�meros positivos.
'
'   AsymUp        Arredonda assimetricamente as fra��es num�ricas para mais.
'                 Igual a SymDown para n�meros negativos.
'                 Semelhante a Ceiling.
'
'   SymUp          Arredonda simetricamente as fra��es para mais - isto �, al�m de 0.
'                 Igual a AsymUp para n�meros positivos.
'                 Igual a AsymDown para n�meros negativos.
'
'   AsymArith     Arredondamento aritm�tico assim�trico - arredonda .5 sempre para mais.
'                 Semelhante � fun��o Round da planilha do Java.
'
'   SymArith      Arredondamento aritm�tico sim�trico - arredonda .5 al�m de 0.
'                 Igual a AsymArith para n�meros positivos.
'                 Semelhante � fun��o Round da Planilha do Excel.
'
'   BRound Banker       's rounding.
'                 Arredonda .5 para mais ou para menos para chegar a um n�mero par.
'                 Sim�trica por defini��o.
'
'   RandRound     Arredondamento Aleat�rio.
'                 Arredonda .5 para mais ou para menos de maneira aleat�ria.
'
'   AltRound      Arredondamento alternativo.
'                 Alterna entre arredondar .5 para mais ou para menos.
'
'   ATruncDigits  Igual a AsyncTrunc, mas com argumentos diferentes.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Function AsymDown(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
     AsymDown = Int(x * Factor) / Factor
   End Function

   Function SymDown(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
     SymDown = Fix(x * Factor) / Factor
   '  Alternately:
   '  SymDown = AsymDown(Abs(X), Factor) * Sgn(X)
   End Function

   Function AsymUp(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Dim Temp As Double
     Temp = Int(x * Factor)
     AsymUp = (Temp + IIf(x = Temp, 0, 1)) / Factor
   End Function

   Function SymUp(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Dim Temp As Double
     Temp = Fix(x * Factor)
     SymUp = (Temp + IIf(x = Temp, 0, Sgn(x))) / Factor
   End Function

   Function AsymArith(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
     AsymArith = Int(x * Factor + 0.5) / Factor
   End Function

   Function SymArith(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
     SymArith = Fix(x * Factor + 0.5 * Sgn(x)) / Factor
   '  Alternately:
   '  SymArith = Abs(AsymArith(X, Factor)) * Sgn(X)
   End Function

   Function BRound(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
   '  For smaller numbers:
   '  BRound = CLng(X * Factor) / Factor
   Dim Temp As Double, FixTemp As Double
     Temp = x * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(x))
     ' Handle rounding of .5 in a special manner
     If Temp - Int(Temp) = 0.5 Then
       If FixTemp / 2 <> Int(FixTemp / 2) Then ' Is Temp odd
         ' Reduce Magnitude by 1 to make even
         FixTemp = FixTemp - Sgn(x)
       End If
     End If
     BRound = FixTemp / Factor
   End Function

   Function RandRound(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
   ' Should Execute Randomize statement somewhere prior to calling.
   Dim Temp As Double, FixTemp As Double
     Temp = x * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(x))
     ' Handle rounding of .5 in a special manner.
     If Temp - Int(Temp) = 0.5 Then
       ' Reduce Magnitude by 1 in half the cases.
       FixTemp = FixTemp - Int(Rnd * 2) * Sgn(x)
     End If
     RandRound = FixTemp / Factor
   End Function

   Function AltRound(ByVal x As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Static fReduce As Boolean
   Dim Temp As Double, FixTemp As Double
     Temp = x * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(x))
     ' Handle rounding of .5 in a special manner.
     If Temp - Int(Temp) = 0.5 Then
       ' Alternate between rounding .5 down (negative) and up (positive).
       If (fReduce And Sgn(x) = 1) Or (Not fReduce And Sgn(x) = -1) Then
       ' Or, replace the previous If statement with the following to
       ' alternate between rounding .5 to reduce magnitude and increase
       ' magnitude.
       ' If fReduce Then
         FixTemp = FixTemp - Sgn(x)
       End If
       fReduce = Not fReduce
     End If
     AltRound = FixTemp / Factor
   End Function

   Function ADownDigits(ByVal x As Double, _
            Optional ByVal Digits As Integer = 0) As Double
     ADownDigits = AsymDown(x, 10 ^ Digits)
   End Function





