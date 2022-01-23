Attribute VB_Name = "PortfolioOptimization"
Option Explicit

Const TimeSeriesLength_days = 365

Dim wsTimeSeries As Worksheet
Dim wsEfficientFrontier As Worksheet

Dim totalCalls As Long
Dim outputLiine&, nOptimizations&, nSparsing&

Sub PortfolioOptimization()
    Dim covarianceMatrix#(), inWeights#(), outWeights#()
    Dim i&, j&, nAssets&, rt#, k&, nDays&, maxRT#, maxVariance#, sumWeights#
    Dim expectedReturns_day#()
    Dim return_#, variance#, sharpeRatio#, maxSharpeRatio#, maxWeights#(), maxReturn_year#
    Dim lbd#(), ubd#(), mixe#, mixsd#
    Dim wsPortfolio As Worksheet
    Dim wsCovariance As Worksheet
    Dim wsReturns As Worksheet
    Dim bottomRow&
    
    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsReturns = ThisWorkbook.Worksheets("Returns")
    Set wsCovariance = ThisWorkbook.Worksheets("Covariance")
    Set wsEfficientFrontier = ThisWorkbook.Worksheets("Efficient frontier")
    
    ScreenUpdatingOff
    
    bottomRow = wsEfficientFrontier.Cells(Rows.Count, 1).End(xlUp).Row
    If bottomRow > 1 Then
        wsEfficientFrontier.Range(wsEfficientFrontier.Cells(2, 1), wsEfficientFrontier.Cells(bottomRow, 4)).ClearContents
    End If
    
    bottomRow = wsEfficientFrontier.Cells(Rows.Count, 7).End(xlUp).Row
    If bottomRow > 1 Then
        wsEfficientFrontier.Range(wsEfficientFrontier.Cells(2, 7), wsEfficientFrontier.Cells(bottomRow, 9)).ClearContents
    End If
    
    ' find number of assets
    i = 2
    While wsPortfolio.Cells(i, 1) <> ""
        i = i + 1
    Wend
    nAssets = i - 2
    outputLiine = 2
    nOptimizations = 0
    nSparsing = Round((1567 * Application.WorksheetFunction.Ln(nAssets) - 536) / 300)     ' See comment to "PrintRandomSteps" function.
    
    ReDim covarianceMatrix(nAssets, nAssets): ReDim inWeights(nAssets): ReDim outWeights(nAssets)
    ReDim expectedReturns_day(nAssets)
    ReDim lbd(nAssets)
    ReDim ubd(nAssets)
    ReDim maxWeights(nAssets)
    
    nDays = wsReturns.Cells(2, Columns.Count).End(xlToLeft).Column - 1
    For i = 1 To nAssets
        For j = 1 To nDays
            expectedReturns_day(i) = expectedReturns_day(i) + wsReturns.Cells(i + 1, j + 1)
        Next j
        expectedReturns_day(i) = expectedReturns_day(i) / nDays
        
        lbd(i) = wsPortfolio.Cells(i + 1, 3)
        ubd(i) = wsPortfolio.Cells(i + 1, 4)
        For j = 1 To nAssets
            covarianceMatrix(i, j) = wsCovariance.Cells(i + 1, j + 1)
        Next j
    Next i
    
    ' 1. Calculate Sharpe ratio for all assets (these values are known from other web sites).
    For i = 1 To nAssets
        For j = 1 To nAssets
            If i = j Then
                inWeights(j) = 1#
            Else
                inWeights(j) = 0#
            End If
        Next j
        
        Call CalculatePortfolioOutputs(nAssets, expectedReturns_day, inWeights, covarianceMatrix, nDays, return_, variance, sharpeRatio)
        
        wsPortfolio.Cells(i + 1, 5) = sharpeRatio
    Next i
    
    ' 2. Calculate efficient frontier curve.
    ' a. LBD/UBD-restrictions:
    sumWeights = 0#
    j = 0
    For i = 1 To nAssets
        inWeights(i) = Undefined
        
        If lbd(i) > 1# Then
            MsgBox ("LBD(" & i & ") must be lower than 1.")
            Exit Sub
        End If
        
        If lbd(i) < 0# Then
            MsgBox ("LBD(" & i & ") must be higher than 0.")
            Exit Sub
        End If
        
        If ubd(i) > 1# Then
            MsgBox ("UBD(" & i & ") must be lower than 1.")
            Exit Sub
        End If
        
        If ubd(i) < 0# Then
            MsgBox ("UBD(" & i & ") must be higher than 0.")
            Exit Sub
        End If
        
        If ubd(i) < lbd(i) Then
            MsgBox ("UBD(" & i & ") must be higher than LBD(" & i & ").")
            Exit Sub
        End If
        
        If lbd(i) > 0 Then
            inWeights(i) = lbd(i)
            sumWeights = sumWeights + lbd(i)
            j = j + 1
        End If
    Next i
    
    If sumWeights > 1# Then
        MsgBox ("Sum of all LBD values must be lower than 1.")
        Exit Sub
    End If
    
    sumWeights = 1# - sumWeights
    j = nAssets - j
    For i = 1 To nAssets
        If inWeights(i) = Undefined Then
            inWeights(i) = sumWeights / j
        End If
    Next i
    
    ' b. Optimization.
    k = 2
    maxSharpeRatio = -1000#
    
    For rt = 0.001 To 0.2 Step 0.001
        Call GQP(rt, nAssets, expectedReturns_day, covarianceMatrix, lbd, ubd, inWeights, outWeights, mixe, mixsd, nDays)
        
        Call CalculatePortfolioOutputs(nAssets, expectedReturns_day, outWeights, covarianceMatrix, nDays, return_, variance, sharpeRatio)
        
        wsEfficientFrontier.Cells(k, 1) = (1# + variance) ^ nDays - 1#
        wsEfficientFrontier.Cells(k, 2) = rt
        wsEfficientFrontier.Cells(k, 3) = sharpeRatio
        wsEfficientFrontier.Cells(k, 4) = (1# + return_) ^ nDays - 1#
        
        If sharpeRatio > maxSharpeRatio Then
            For i = 1 To nAssets
                maxWeights(i) = outWeights(i)
            Next i
            maxReturn_year = (1# + return_) ^ nDays - 1#
            maxSharpeRatio = sharpeRatio
            maxRT = rt
            maxVariance = (1# + variance) ^ nDays - 1#
        End If
        
        k = k + 1
    Next rt
    
    For i = 1 To nAssets
        wsPortfolio.Cells(i + 1, 6) = maxWeights(i)
        wsPortfolio.Cells(1, 12) = maxSharpeRatio
        wsPortfolio.Cells(2, 12) = maxReturn_year
        wsPortfolio.Cells(3, 12) = maxRT
        wsPortfolio.Cells(4, 12) = maxVariance
    Next i
    
    ScreenUpdatingOn
End Sub

' Writes out some intermediate steps of optimization to show many possible portfolio outcomes below efficient frontier curve.
' Every (1567*ln(nAssets)-536) / 300 point will be written out.
' For nAssets = 9, every 10th point will be written out.
Sub PrintRandomSteps(nAssets&, expectedReturns_day#(), inWeights#(), covarianceMatrix#(), nDays&, return_#)
    If nOptimizations = nSparsing Then
        Dim variance#, sharpeRatio#
        
        Call CalculatePortfolioOutputs(nAssets, expectedReturns_day, inWeights, covarianceMatrix, nDays, return_, variance, sharpeRatio)
        
        wsEfficientFrontier.Cells(outputLiine, 7) = (1# + variance) ^ nDays - 1#
        wsEfficientFrontier.Cells(outputLiine, 8) = sharpeRatio
        wsEfficientFrontier.Cells(outputLiine, 9) = (1# + return_) ^ nDays - 1#
        
        outputLiine = outputLiine + 1
        nOptimizations = 0
    Else
        nOptimizations = nOptimizations + 1
    End If
End Sub

Sub CalculateCharacteristicsOfUsedWeights()
    Dim covarianceMatrix#(), inWeights#(), expectedReturns_year#()
    Dim i&, j&, nAssets&, rt#, k&, nDays&, sumWeights#
    Dim expectedReturns_day#()
    Dim return_#, variance#, sharpeRatio#
    Dim wsPortfolio As Worksheet
    Dim wsCovariance As Worksheet
    
    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsTimeSeries = ThisWorkbook.Worksheets("Time series")
    Set wsCovariance = ThisWorkbook.Worksheets("Covariance")
    
    ScreenUpdatingOff
    
    ' find number of assets
    i = 2
    While wsPortfolio.Cells(i, 1) <> ""
        i = i + 1
    Wend
    nAssets = i - 2
    
    ReDim covarianceMatrix(nAssets, nAssets): ReDim inWeights(nAssets)
    ReDim expectedReturns_year(nAssets): ReDim expectedReturns_day(nAssets)
    
    nDays = wsTimeSeries.Cells(1, Columns.Count).End(xlToLeft).Column - 1
    sumWeights = 0#
    For i = 1 To nAssets
        expectedReturns_year(i) = (wsTimeSeries.Cells(i + 1, nDays + 1) - wsTimeSeries.Cells(i + 1, 2)) / wsTimeSeries.Cells(i + 1, 2)
        expectedReturns_day(i) = expectedReturns_year(i) / nDays
        inWeights(i) = wsPortfolio.Cells(i + 1, 6)
        sumWeights = sumWeights + inWeights(i)
        For j = 1 To nAssets
            covarianceMatrix(i, j) = wsCovariance.Cells(i + 1, j + 1)
        Next j
    Next i
    
    If sumWeights > 1.0000001 Or sumWeights < 0.999999 Then
        MsgBox ("Sum of all weights must be 1.")
        Exit Sub
    End If
    
    Call CalculatePortfolioOutputs(nAssets, expectedReturns_day, inWeights, covarianceMatrix, nDays, return_, variance, sharpeRatio)
    
    wsPortfolio.Cells(1, 12) = sharpeRatio
    wsPortfolio.Cells(2, 12) = return_ * nDays
    wsPortfolio.Cells(3, 12) = "unknown"
    wsPortfolio.Cells(4, 12) = variance * nDays
    
    ScreenUpdatingOn
End Sub

' Port of Matlab code from web.stanford.edu/~wfsharpe/mat/gqp.txt
' Implements algorithm from https://www.gsb.stanford.edu/faculty-research/working-papers/algorithm-portfolio-improvement
' See also https://quant.stackexchange.com/questions/39594/maximize-sharpe-ratio-in-portfolio-optimization/41632
' Input:
'   rt: risk tolerance
'   nAssets: number of assets
'   e: expected return vector
'   C: covariance matrix
'   lbd: lower bound vector
'   ubd: upper bound vector
'   x0: initial feasible mix vector
' Output:
'   x: optimal mix vector
'   mixe: x'*e
'   mixsd: sqrt(x'*C*x)

'        maximizes:    rt*(x'*e) - x'*C*x
'         subject to: sum(x) = sum(x0)
'                     lbd <= x <= ubd
' algorithm based on:
'    William F. Sharpe, "An Algorithm for Portfolio Improvement,"
'    in Advances in Mathematical Programming and Financial Planning
'    K.D. Lawrence, J.B. Guerard, Jr., and Gary D. Reeves, Editors
'    JAI Press, Inc., 1987, pp. 155-170.
Sub GQP(rt#, nAssets&, e#(), C#(), lbd#(), ubd#(), x0#(), x#(), mixe#, mixsd#, nDays&)
    Dim k#(3)
    Dim maxit&, minMUchg#, n#, iterations&, muAdd#, muSub#, aAdd&, aSub&, t1#, t2#, t3#, kmin#, i&, j&
    Dim mu#(), slack#()     ' d#()
    
    ReDim mu(nAssets): ReDim d(nAssets): ReDim slack(nAssets)
    
' maximum number of iterations
    maxit = 500
    
' set minimum MU change to continue
    minMUchg = 0.000001
    
' initialize
    For i = 1 To nAssets
        x(i) = x0(i)
    Next i
    
    n = nAssets
    
' continue to improve portfolio until further improvement impossible
'   when done, return
    
    iterations = 0
    
    Do While True
        totalCalls = totalCalls + 1
        
        ' compute marginal utilities
        For i = 1 To nAssets
            mu(i) = rt * e(i)
            For j = 1 To nAssets
                mu(i) = mu(i) - 2 * C(i, j) * x(j)
            Next j
        Next i
        
        muAdd = -1E+200: aAdd = 0
        muSub = 1E+200: aSub = 0
        For i = 1 To nAssets
            ' find best variable to add
            ' [MUadd,Aadd] = max(mu - 1E200*(x>=ubd));
            If muAdd < mu(i) And x(i) < ubd(i) Then
                muAdd = mu(i)
                aAdd = i
            End If
            
            ' find best variable to subtract
            ' [MUsub,Asub] =  min(mu + 1E200*(x<=lbd));
            If muSub > mu(i) And x(i) > lbd(i) Then
                muSub = mu(i)
                aSub = i
            End If
        Next i
        
        ' terminate and return if change in mu is less than minimum
        If (muAdd - muSub) <= minMUchg Then
            ' compute mix e and sd
            Call Compute_MixE_MixSD(nAssets, x, e, C, mixe, mixsd)
            
            ' terminate
            Exit Do
        End If
        
'    ' set up delta vector
    ' d = zeros(n,1);
    ' d(Aadd) = 1;
    ' d(Asub) = -1;
        
    ' compute step size
        k(1) = 0#: k(2) = 0#: k(3) = 0#
       
       ' optimal unconstrained step size
       ' k(1) = ((rt*d'*e)-2*(x'*C*d))/(2*(d'*C*d));
        t1 = rt * e(aAdd) - rt * e(aSub)
        t2 = 0#
        For i = 1 To nAssets
            t2 = t2 + 2 * x(i) * C(i, aAdd) - 2 * x(i) * C(i, aSub)
        Next i
        t3 = 2 * (C(aAdd, aAdd) - C(aAdd, aSub) - C(aSub, aAdd) + C(aSub, aSub))
        k(1) = (t1 - t2) / t3
        
       ' maximum step size based on upper bounds
        k(2) = ubd(aAdd) - x(aAdd)
        
       ' maximum step size based on lower bounds
        k(3) = x(aSub) - lbd(aSub)
        
       ' minimum step size
        kmin = Application.WorksheetFunction.Min(k(1), k(2), k(3))
       
       ' terminate and return if minumum step size is zero
        If kmin = 0 Then
            ' compute mix e and sd
            Call Compute_MixE_MixSD(nAssets, x, e, C, mixe, mixsd)
            
            ' terminate
            Exit Do
        End If
        
      ' count and terminate if maximum iterations exceeded
         iterations = iterations + 1
         If iterations > maxit Then
            ' compute mix e and sd
            Call Compute_MixE_MixSD(nAssets, x, e, C, mixe, mixsd)
            
            ' terminate
            Exit Do
         End If
         
        ' change mix
        ' x = x + ( kmin*d) ;
        x(aAdd) = x(aAdd) + kmin
        x(aSub) = x(aSub) - kmin
        
        Call PrintRandomSteps(nAssets, e, x, C, nDays, mixe)
    Loop
End Sub

Sub Compute_MixE_MixSD(nAssets&, x#(), e#(), C#(), mixe#, mixsd#)
    Dim i&, j&
    
    mixe = 0#
    mixsd = 0#
    
    For i = 1 To nAssets
        mixe = mixe + x(i) * e(i)
        
        For j = 1 To nAssets
            mixsd = mixsd + x(i) * x(j) * C(i, j)
        Next j
    Next i
    mixsd = Sqr(mixsd)
End Sub

' calculate return, variance and Sharpe ratio of portfolio
Sub CalculatePortfolioOutputs(nAssets&, expectedReturn_day#(), weights#(), covariance#(), nDays&, return_#, variance#, sharpeRatio#)
    Dim i&, j&
    
    return_ = 0#: variance = 0#: sharpeRatio = 0#
    
    For i = 1 To UBound(weights)
        return_ = return_ + weights(i) * expectedReturn_day(i)
        For j = 1 To UBound(weights)
            variance = variance + weights(i) * weights(j) * covariance(i, j)
        Next j
    Next i
    
    variance = Sqr(variance / nDays)
    
    sharpeRatio = return_ / variance
End Sub

' Reads time series and creates covariance matrix
Sub InitializePortfolioOptimization()
    Dim wsPortfolio As Worksheet
    Dim wsCovariance As Worksheet
    Dim wsReturns As Worksheet
    Dim period1$, period2$, i&, j&, k&, nAssets&, nColumns&
    Dim outDates_reference() As Date    ' reference dates from AAPL (why? see comments below)
    Dim outDates() As Date
    Dim outTimeSeries#(), inputTimeSeries#(), tmpDate As Date
    Dim timeSeries_AllAssets#()
    Dim columnIsUsed() As Boolean
    Dim rightColumn&, bottomRow&
    Dim lastDate As Date
    
    ScreenUpdatingOff
    
    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsTimeSeries = ThisWorkbook.Worksheets("Time series")
    Set wsCovariance = ThisWorkbook.Worksheets("Covariance")
    Set wsReturns = ThisWorkbook.Worksheets("Returns")
    
    ' Cleanup: sheets "TimeSeries", "Covariance" and "Returs" will be completely cleaned.
    wsTimeSeries.Cells.ClearContents
    wsCovariance.Cells.ClearContents
    wsReturns.Cells.ClearContents
    
    ' Get number of assets.
    i = 2
    While wsPortfolio.Cells(i, 1) <> ""
        ' Check the correctness of URL.
        If wsPortfolio.Cells(i, 2) = "" Then
            MsgBox ("URL (column B, row " & i & ") must contain either ""Yahoo"" for assets from finance.yahoo.com or URL from ariva.de")
            Exit Sub
        End If
        
        i = i + 1
    Wend
    nAssets = i - 2
    
    getCookieCrumb
    
    period2 = CStr((Int(DateTime.Now) - CDate("01.01.1970")) * 60 * 60 * 24)
    period1 = CStr((Int(DateTime.Now - TimeSeriesLength_days) - CDate("01.01.1970")) * 60 * 60 * 24)
    
'   problem: different markets/countries - different holidays
'   lazy solution on getting the dates:
'   I would take the apple (AAPL) and update the table with its dates
'   this must be enough: the handle with apple is very frequent
    Call getYahooFinanceData("AAPL", period1, period2, "1d", outDates_reference, inputTimeSeries)
    
    nColumns = UBound(outDates_reference)
    ReDim timeSeries_AllAssets(nAssets, UBound(outDates_reference))
    ReDim columnIsUsed(UBound(outDates_reference))
    For i = 1 To UBound(outDates_reference)
        columnIsUsed(i) = True
    Next i
    
'   Reading of portfolio assets.
    For i = 1 To nAssets
        lastDate = outDates_reference(1)
        
        If wsPortfolio.Cells(i + 1, 2) = "Yahoo" Then
            Call getYahooFinanceData(wsPortfolio.Cells(i + 1, 1), period1, period2, "1d", outDates, inputTimeSeries)
            Call ReadSharesTimeSeries(outDates_reference, outDates, inputTimeSeries, outTimeSeries, wsPortfolio.Cells(i + 1, 1), i + 1)
        End If
        
        If InStr(1, wsPortfolio.Cells(i + 1, 2), "ariva") > 0 Then
            Call DeleteFile(Environ$("USERPROFILE") & "\Downloads\wkn_" + CStr(wsPortfolio.Cells(i + 1, 1)) + "_historic.csv")
            
            Call GetAriva_Fund(wsPortfolio.Cells(i + 1, 2), wsPortfolio.Cells(i + 1, 1))
            
            Call ReadFundsTimeSeries(outDates_reference, outTimeSeries, wsPortfolio.Cells(i + 1, 1))
            
            Call DeleteFile(Environ$("USERPROFILE") & "\Downloads\wkn_" + CStr(wsPortfolio.Cells(i + 1, 1)) + "_historic.csv")
        End If
        
        If InStr(1, wsPortfolio.Cells(i + 1, 2), "moex") > 0 Then
            lastDate = outDates_reference(1)
            ReDim outTimeSeries(UBound(outDates_reference))
            
            While GetMoex(wsPortfolio.Cells(i + 1, 2) & "?iss.json=extended&from=" & Year(lastDate) & "-" & Month(lastDate) & "-" & Day(lastDate), _
                    wsPortfolio.Cells(i + 1, 1), outDates_reference, outTimeSeries, lastDate)
            Wend
        End If
        
        ' Add time series of an asset to the "timeSeries_AllAssets"
        For j = 1 To nColumns
            timeSeries_AllAssets(i, j) = outTimeSeries(j)
            
            If outTimeSeries(j) = Undefined Or outTimeSeries(j) = 0 Then
                ' If the value in time series is undefined (empty), the corresponding date will be not shown in "TimeSeries" sheet.
                columnIsUsed(j) = False
            End If
        Next j
        
        Call ProgressBar(wsPortfolio.Cells(i + 1, 1))
    Next i
    
    CloseSeleniumDriver
    
    ' Update dates.
    k = 1
    For i = 1 To nColumns
        If columnIsUsed(i) Then
            wsTimeSeries.Cells(1, 1 + k) = outDates_reference(i)
            k = k + 1
        End If
    Next i
    
    ' Write time series of all assets to the table. If the column in any asset contains undefined (empty) value, then this column for all assets will be not written.
    ' We have to do this procedure since the covariance matrix calculation do not support empty values in time series.
    For i = 1 To nAssets
        wsTimeSeries.Cells(i + 1, 1) = wsPortfolio.Cells(i + 1, 1)
        
        k = 1
        For j = 1 To nColumns
            If columnIsUsed(j) Then
                wsTimeSeries.Cells(i + 1, k + 1) = timeSeries_AllAssets(i, j)
                k = k + 1
            End If
        Next j
    Next i
    
    '   calculate returns
    rightColumn = wsTimeSeries.Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To nAssets
        wsReturns.Cells(i + 1, 1) = wsPortfolio.Cells(i + 1, 1)
        For j = 2 To rightColumn - 1
            wsReturns.Cells(i + 1, j) = (wsTimeSeries.Cells(i + 1, j + 1) - wsTimeSeries.Cells(i + 1, j)) / wsTimeSeries.Cells(i + 1, j)
        Next j
    Next i
    
'   calculate covariance matrix
    For i = 1 To nAssets
        wsCovariance.Cells(1, i + 1) = wsPortfolio.Cells(i + 1, 1)
        wsCovariance.Cells(i + 1, 1) = wsPortfolio.Cells(i + 1, 1)
        For j = i To nAssets
            wsCovariance.Cells(i + 1, j + 1) = Application.WorksheetFunction.Covar( _
                wsReturns.Range(wsReturns.Cells(i + 1, 2), wsReturns.Cells(i + 1, rightColumn - 1)), _
                wsReturns.Range(wsReturns.Cells(j + 1, 2), wsReturns.Cells(j + 1, rightColumn - 1)))
            wsCovariance.Cells(j + 1, i + 1) = wsCovariance.Cells(i + 1, j + 1)
        Next j
    Next i
    
    Application.StatusBar = ""
    
    ScreenUpdatingOn
End Sub
