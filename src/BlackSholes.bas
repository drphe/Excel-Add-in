Attribute VB_Name = "BlackSholes"
Private Function dOne(S, X, T, r, v, d)
dOne = (Log(S / X) + (r - d + 0.5 * v ^ 2) * T) / (v * (Sqr(T)))
End Function
Private Function NdOne(S, X, T, r, v, d)
NdOne = Exp(-(dOne(S, X, T, r, v, d) ^ 2) / 2) / (Sqr(2 * Application.WorksheetFunction.Pi()))
End Function
Private Function dTwo(S, X, T, r, v, d)
dTwo = dOne(S, X, T, r, v, d) - v * Sqr(T)
End Function
Private Function NdTwo(S, X, T, r, v, d) As Double
NdTwo = Application.NormSDist(dTwo(S, X, T, r, v, d))
End Function
Function OptionPrice(OptionType, StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, Optional DividendYield = 0) As Double
If OptionType = "C" Then
    OptionPrice = Exp(-DividendYield * TimeToExpire) * StockPrice * Application.NormSDist(dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield)) - StrikePrice * Exp(-RiskFreeRate * TimeToExpire) * Application.NormSDist(dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield) - Volatility * Sqr(TimeToExpire))
ElseIf OptionType = "P" Then
    OptionPrice = StrikePrice * Exp(-RiskFreeRate * TimeToExpire) * Application.NormSDist(-dTwo(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield)) - Exp(-d * TimeToExpire) * StockPrice * Application.NormSDist(-dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield))
End If
End Function
Function OptionPriceCall(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, Optional DividendYield = 0) As Double
    OptionPriceCall = Exp(-DividendYield * TimeToExpire) * StockPrice * Application.NormSDist(dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield)) - StrikePrice * Exp(-RiskFreeRate * TimeToExpire) * Application.NormSDist(dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield) - Volatility * Sqr(TimeToExpire))
End Function
Function OptionPricePut(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, Optional DividendYield = 0) As Double
    OptionPricePut = StrikePrice * Exp(-RiskFreeRate * TimeToExpire) * Application.NormSDist(-dTwo(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield)) - Exp(-d * TimeToExpire) * StockPrice * Application.NormSDist(-dOne(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Volatility, DividendYield))
End Function
Function OptionSigmaCall(OptionPrice As Double, StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Optional DividendYield As Double = 0)
    Dim sigma, OptionPriceTemp As Double
    Dim Tolerance, epsilon As Double
    Dim i, maxIterations As Integer
    Tolerance = CheckNumber(OptionPrice)
    sigma = 0#  ' Gia su khoi dau
    epsilon = 0.001 ' Sai so cho phep
    maxIterations = 1500 ' So lan lap toi da
    For i = 1 To maxIterations
        OptionSigmaCall = sigma + epsilon * i
        OptionPriceTemp = OptionPriceCall(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, OptionSigmaCall, DividenYield)
        If Abs(OptionPrice - OptionPriceTemp) <= Tolerance Then Exit For
    Next i
End Function
Function OptionSigmaPut(OptionPrice As Double, StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, Optional DividendYield As Double = 0)
    Dim sigma, OptionPriceTemp As Double
    Dim Tolerance, epsilon As Double
    Dim i, maxIterations As Integer
    Tolerance = CheckNumber(OptionPrice)
    sigma = 0#  ' Gia su khoi dau
    epsilon = 0.001 ' Sai so cho phep
    maxIterations = 1000 ' So lan lap toi da
    For i = 1 To maxIterations
        OptionSigmaPut = sigma + epsilon * i
        OptionPriceTemp = OptionPricePut(StockPrice, StrikePrice, TimeToExpire, RiskFreeRate, OptionSigmaPut, DividenYield)
        If Abs(OptionPrice - OptionPriceTemp) <= Tolerance Then Exit For
    Next i
End Function
Private Function CheckNumber(So As Double) As Long
  If So < 10 Then
    CheckNumber = 0.1
  ElseIf So < 100 Then
    CheckNumber = 1
  ElseIf So < 1000 Then
    CheckNumber = 10
  Else
    CheckNumber = 100
  End If
End Function
