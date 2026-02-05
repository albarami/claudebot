'===============================================================================
' PhD Survey Analysis - Excel UDFs
' Statistical functions not natively supported in Excel
' These UDFs supplement Excel's built-in functions for academic analysis
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' SHAPIRO_WILK: Approximation of Shapiro-Wilk normality test
' Returns: Array(W statistic, p-value)
' Note: Uses Royston's algorithm approximation for n <= 5000
'-------------------------------------------------------------------------------
Public Function SHAPIRO_WILK(dataRange As Range) As Variant
    On Error GoTo ErrorHandler
    
    Dim data() As Double
    Dim n As Long, i As Long, j As Long
    Dim cell As Range
    Dim validCount As Long
    
    ' Count valid numeric values
    validCount = 0
    For Each cell In dataRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            validCount = validCount + 1
        End If
    Next cell
    
    If validCount < 3 Then
        SHAPIRO_WILK = CVErr(xlErrValue)
        Exit Function
    End If
    
    n = validCount
    ReDim data(1 To n)
    
    ' Extract numeric values
    i = 1
    For Each cell In dataRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            data(i) = CDbl(cell.Value)
            i = i + 1
        End If
    Next cell
    
    ' Sort data (bubble sort for simplicity)
    Dim temp As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If data(i) > data(j) Then
                temp = data(i)
                data(i) = data(j)
                data(j) = temp
            End If
        Next j
    Next i
    
    ' Calculate mean
    Dim mean As Double
    mean = 0
    For i = 1 To n
        mean = mean + data(i)
    Next i
    mean = mean / n
    
    ' Calculate S^2 (denominator)
    Dim S2 As Double
    S2 = 0
    For i = 1 To n
        S2 = S2 + (data(i) - mean) ^ 2
    Next i
    
    ' Calculate a coefficients and b (numerator)
    Dim m As Long
    Dim b As Double
    m = Int(n / 2)
    b = 0
    
    ' Simplified approximation of a_i coefficients
    Dim a() As Double
    ReDim a(1 To m)
    
    For i = 1 To m
        ' Approximation using normal order statistics
        a(i) = NormalOrderStatistic(n - i + 1, n) - NormalOrderStatistic(i, n)
    Next i
    
    ' Normalize a coefficients
    Dim sumA2 As Double
    sumA2 = 0
    For i = 1 To m
        sumA2 = sumA2 + a(i) ^ 2
    Next i
    sumA2 = Sqr(sumA2)
    
    For i = 1 To m
        a(i) = a(i) / sumA2
    Next i
    
    ' Calculate b
    For i = 1 To m
        b = b + a(i) * (data(n - i + 1) - data(i))
    Next i
    
    ' Calculate W statistic
    Dim W As Double
    If S2 > 0 Then
        W = (b ^ 2) / S2
    Else
        W = 1
    End If
    
    ' Ensure W is in valid range
    If W > 1 Then W = 1
    If W < 0 Then W = 0
    
    ' Approximate p-value using Royston's method
    Dim pValue As Double
    pValue = ShapiroWilkPValue(W, n)
    
    ' Return as 2-element array
    Dim result(1 To 2) As Variant
    result(1) = W
    result(2) = pValue
    SHAPIRO_WILK = result
    Exit Function
    
ErrorHandler:
    SHAPIRO_WILK = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' Helper: Normal order statistic approximation
'-------------------------------------------------------------------------------
Private Function NormalOrderStatistic(i As Long, n As Long) As Double
    Dim p As Double
    p = (i - 0.375) / (n + 0.25)
    NormalOrderStatistic = Application.WorksheetFunction.Norm_S_Inv(p)
End Function

'-------------------------------------------------------------------------------
' Helper: Shapiro-Wilk p-value approximation (Royston 1992)
'-------------------------------------------------------------------------------
Private Function ShapiroWilkPValue(W As Double, n As Long) As Double
    Dim mu As Double, sigma As Double, z As Double
    
    ' Approximation parameters for different sample sizes
    If n <= 11 Then
        mu = 0.0038915 * n ^ 3 - 0.083751 * n ^ 2 - 0.31082 * n - 1.5861
        sigma = Exp(0.0030302 * n ^ 2 - 0.082676 * n - 0.4803)
    Else
        Dim logN As Double
        logN = Log(n)
        mu = -1.5861 - 0.31082 * logN - 0.083751 * logN ^ 2 + 0.0038915 * logN ^ 3
        sigma = Exp(-0.4803 - 0.082676 * logN + 0.0030302 * logN ^ 2)
    End If
    
    ' Transform W to normal
    If W > 0 And W < 1 Then
        z = (Log(1 - W) - mu) / sigma
    Else
        z = 0
    End If
    
    ' Return p-value from standard normal
    ShapiroWilkPValue = 1 - Application.WorksheetFunction.Norm_S_Dist(z, True)
    
    ' Bound p-value
    If ShapiroWilkPValue < 0 Then ShapiroWilkPValue = 0
    If ShapiroWilkPValue > 1 Then ShapiroWilkPValue = 1
End Function

'-------------------------------------------------------------------------------
' LEVENE_TEST: Levene's test for equality of variances
' Accepts 2-10 ranges representing different groups
' Returns: Array(F statistic, p-value)
'-------------------------------------------------------------------------------
Public Function LEVENE_TEST(ParamArray ranges() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim k As Long  ' Number of groups
    Dim N As Long  ' Total sample size
    Dim i As Long, j As Long
    
    k = UBound(ranges) - LBound(ranges) + 1
    If k < 2 Then
        LEVENE_TEST = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Extract data from each group
    Dim groupData() As Variant
    Dim groupN() As Long
    Dim groupMedian() As Double
    ReDim groupData(1 To k)
    ReDim groupN(1 To k)
    ReDim groupMedian(1 To k)
    
    N = 0
    For i = 1 To k
        Dim rng As Range
        Set rng = ranges(i - 1)
        Dim vals() As Double
        Dim cnt As Long
        cnt = 0
        
        Dim cell As Range
        For Each cell In rng
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                cnt = cnt + 1
                ReDim Preserve vals(1 To cnt)
                vals(cnt) = CDbl(cell.Value)
            End If
        Next cell
        
        If cnt < 2 Then
            LEVENE_TEST = CVErr(xlErrValue)
            Exit Function
        End If
        
        groupData(i) = vals
        groupN(i) = cnt
        groupMedian(i) = MedianOfArray(vals, cnt)
        N = N + cnt
    Next i
    
    ' Calculate Z_ij = |Y_ij - median_i|
    Dim Z() As Double
    Dim groupZ() As Variant
    ReDim groupZ(1 To k)
    
    Dim grandMeanZ As Double
    grandMeanZ = 0
    
    For i = 1 To k
        Dim zVals() As Double
        ReDim zVals(1 To groupN(i))
        Dim gd() As Double
        gd = groupData(i)
        
        For j = 1 To groupN(i)
            zVals(j) = Abs(gd(j) - groupMedian(i))
            grandMeanZ = grandMeanZ + zVals(j)
        Next j
        groupZ(i) = zVals
    Next i
    
    grandMeanZ = grandMeanZ / N
    
    ' Calculate group means of Z
    Dim groupMeanZ() As Double
    ReDim groupMeanZ(1 To k)
    
    For i = 1 To k
        Dim zv() As Double
        zv = groupZ(i)
        groupMeanZ(i) = 0
        For j = 1 To groupN(i)
            groupMeanZ(i) = groupMeanZ(i) + zv(j)
        Next j
        groupMeanZ(i) = groupMeanZ(i) / groupN(i)
    Next i
    
    ' Calculate F statistic
    Dim SSbetween As Double, SSwithin As Double
    SSbetween = 0
    SSwithin = 0
    
    For i = 1 To k
        SSbetween = SSbetween + groupN(i) * (groupMeanZ(i) - grandMeanZ) ^ 2
        
        Dim zv2() As Double
        zv2 = groupZ(i)
        For j = 1 To groupN(i)
            SSwithin = SSwithin + (zv2(j) - groupMeanZ(i)) ^ 2
        Next j
    Next i
    
    Dim df1 As Long, df2 As Long
    df1 = k - 1
    df2 = N - k
    
    Dim F As Double
    If SSwithin > 0 Then
        F = (SSbetween / df1) / (SSwithin / df2)
    Else
        F = 0
    End If
    
    ' Calculate p-value
    Dim pValue As Double
    pValue = Application.WorksheetFunction.F_Dist_RT(F, df1, df2)
    
    Dim result(1 To 2) As Variant
    result(1) = F
    result(2) = pValue
    LEVENE_TEST = result
    Exit Function
    
ErrorHandler:
    LEVENE_TEST = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' Helper: Calculate median of array
'-------------------------------------------------------------------------------
Private Function MedianOfArray(arr() As Double, n As Long) As Double
    ' Sort array
    Dim i As Long, j As Long
    Dim temp As Double
    
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Return median
    If n Mod 2 = 1 Then
        MedianOfArray = arr((n + 1) / 2)
    Else
        MedianOfArray = (arr(n / 2) + arr(n / 2 + 1)) / 2
    End If
End Function

'-------------------------------------------------------------------------------
' CRONBACH_ALPHA: Calculate Cronbach's alpha for scale reliability
' dataRange: Range containing items as columns, respondents as rows
' Returns: Alpha coefficient (0 to 1)
'-------------------------------------------------------------------------------
Public Function CRONBACH_ALPHA(dataRange As Range) As Variant
    On Error GoTo ErrorHandler
    
    Dim data() As Double
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim cell As Range
    
    nRows = dataRange.Rows.Count
    nCols = dataRange.Columns.Count
    
    If nCols < 2 Then
        CRONBACH_ALPHA = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Load data into array
    ReDim data(1 To nRows, 1 To nCols)
    For i = 1 To nRows
        For j = 1 To nCols
            If IsNumeric(dataRange.Cells(i, j).Value) Then
                data(i, j) = CDbl(dataRange.Cells(i, j).Value)
            Else
                data(i, j) = 0  ' Handle missing as 0 (could use mean imputation)
            End If
        Next j
    Next i
    
    ' Calculate item variances
    Dim itemMean() As Double
    Dim itemVar() As Double
    ReDim itemMean(1 To nCols)
    ReDim itemVar(1 To nCols)
    
    Dim sumItemVar As Double
    sumItemVar = 0
    
    For j = 1 To nCols
        ' Calculate mean
        itemMean(j) = 0
        For i = 1 To nRows
            itemMean(j) = itemMean(j) + data(i, j)
        Next i
        itemMean(j) = itemMean(j) / nRows
        
        ' Calculate variance
        itemVar(j) = 0
        For i = 1 To nRows
            itemVar(j) = itemVar(j) + (data(i, j) - itemMean(j)) ^ 2
        Next i
        itemVar(j) = itemVar(j) / (nRows - 1)  ' Sample variance
        sumItemVar = sumItemVar + itemVar(j)
    Next j
    
    ' Calculate total score variance
    Dim totalScore() As Double
    ReDim totalScore(1 To nRows)
    Dim totalMean As Double
    totalMean = 0
    
    For i = 1 To nRows
        totalScore(i) = 0
        For j = 1 To nCols
            totalScore(i) = totalScore(i) + data(i, j)
        Next j
        totalMean = totalMean + totalScore(i)
    Next i
    totalMean = totalMean / nRows
    
    Dim totalVar As Double
    totalVar = 0
    For i = 1 To nRows
        totalVar = totalVar + (totalScore(i) - totalMean) ^ 2
    Next i
    totalVar = totalVar / (nRows - 1)
    
    ' Calculate alpha
    Dim k As Long
    k = nCols
    
    If totalVar > 0 Then
        CRONBACH_ALPHA = (k / (k - 1)) * (1 - sumItemVar / totalVar)
    Else
        CRONBACH_ALPHA = 0
    End If
    
    Exit Function
    
ErrorHandler:
    CRONBACH_ALPHA = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' FISHER_Z: Fisher's r-to-z transformation
' r: Pearson correlation coefficient (-1 to 1)
' Returns: Fisher's z value
'-------------------------------------------------------------------------------
Public Function FISHER_Z(r As Double) As Variant
    On Error GoTo ErrorHandler
    
    If r <= -1 Or r >= 1 Then
        FISHER_Z = CVErr(xlErrValue)
        Exit Function
    End If
    
    FISHER_Z = 0.5 * Log((1 + r) / (1 - r))
    Exit Function
    
ErrorHandler:
    FISHER_Z = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' FISHER_Z_INV: Inverse Fisher transformation (z to r)
' z: Fisher's z value
' Returns: Correlation coefficient
'-------------------------------------------------------------------------------
Public Function FISHER_Z_INV(z As Double) As Variant
    On Error GoTo ErrorHandler
    
    FISHER_Z_INV = (Exp(2 * z) - 1) / (Exp(2 * z) + 1)
    Exit Function
    
ErrorHandler:
    FISHER_Z_INV = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' P_VALUE_T: Calculate p-value from t statistic (two-tailed)
' t: t statistic
' df: degrees of freedom
' Returns: Two-tailed p-value
'-------------------------------------------------------------------------------
Public Function P_VALUE_T(t As Double, df As Long) As Variant
    On Error GoTo ErrorHandler
    
    If df < 1 Then
        P_VALUE_T = CVErr(xlErrValue)
        Exit Function
    End If
    
    P_VALUE_T = Application.WorksheetFunction.T_Dist_2T(Abs(t), df)
    Exit Function
    
ErrorHandler:
    P_VALUE_T = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' P_VALUE_F: Calculate p-value from F statistic
' F: F statistic
' df1: numerator degrees of freedom
' df2: denominator degrees of freedom
' Returns: Right-tail p-value
'-------------------------------------------------------------------------------
Public Function P_VALUE_F(F As Double, df1 As Long, df2 As Long) As Variant
    On Error GoTo ErrorHandler
    
    If df1 < 1 Or df2 < 1 Or F < 0 Then
        P_VALUE_F = CVErr(xlErrValue)
        Exit Function
    End If
    
    P_VALUE_F = Application.WorksheetFunction.F_Dist_RT(F, df1, df2)
    Exit Function
    
ErrorHandler:
    P_VALUE_F = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' COHENS_D: Calculate Cohen's d effect size
' mean1, sd1, n1: Group 1 statistics
' mean2, sd2, n2: Group 2 statistics
' Returns: Cohen's d (using pooled SD)
'-------------------------------------------------------------------------------
Public Function COHENS_D(mean1 As Double, sd1 As Double, n1 As Long, _
                         mean2 As Double, sd2 As Double, n2 As Long) As Variant
    On Error GoTo ErrorHandler
    
    If n1 < 2 Or n2 < 2 Then
        COHENS_D = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Calculate pooled standard deviation
    Dim pooledSD As Double
    pooledSD = Sqr(((n1 - 1) * sd1 ^ 2 + (n2 - 1) * sd2 ^ 2) / (n1 + n2 - 2))
    
    If pooledSD = 0 Then
        COHENS_D = 0
    Else
        COHENS_D = (mean1 - mean2) / pooledSD
    End If
    
    Exit Function
    
ErrorHandler:
    COHENS_D = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' ETA_SQUARED: Calculate eta-squared effect size from ANOVA
' SSbetween: Sum of squares between groups
' SStotal: Total sum of squares
' Returns: Eta-squared (proportion of variance explained)
'-------------------------------------------------------------------------------
Public Function ETA_SQUARED(SSbetween As Double, SStotal As Double) As Variant
    On Error GoTo ErrorHandler
    
    If SStotal <= 0 Then
        ETA_SQUARED = CVErr(xlErrValue)
        Exit Function
    End If
    
    ETA_SQUARED = SSbetween / SStotal
    Exit Function
    
ErrorHandler:
    ETA_SQUARED = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' CRAMERS_V: Calculate Cramer's V for chi-square
' chiSq: Chi-square statistic
' n: Sample size
' minDim: Minimum of (rows-1, cols-1) in contingency table
' Returns: Cramer's V (0 to 1)
'-------------------------------------------------------------------------------
Public Function CRAMERS_V(chiSq As Double, n As Long, minDim As Long) As Variant
    On Error GoTo ErrorHandler
    
    If n <= 0 Or minDim < 1 Then
        CRAMERS_V = CVErr(xlErrValue)
        Exit Function
    End If
    
    CRAMERS_V = Sqr(chiSq / (n * minDim))
    Exit Function
    
ErrorHandler:
    CRAMERS_V = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' CI_MEAN: Calculate confidence interval for mean
' mean: Sample mean
' sd: Sample standard deviation
' n: Sample size
' confidenceLevel: Confidence level (e.g., 0.95)
' Returns: Array(lower bound, upper bound)
'-------------------------------------------------------------------------------
Public Function CI_MEAN(mean As Double, sd As Double, n As Long, _
                        Optional confidenceLevel As Double = 0.95) As Variant
    On Error GoTo ErrorHandler
    
    If n < 2 Or confidenceLevel <= 0 Or confidenceLevel >= 1 Then
        CI_MEAN = CVErr(xlErrValue)
        Exit Function
    End If
    
    Dim alpha As Double
    Dim tCrit As Double
    Dim marginError As Double
    
    alpha = 1 - confidenceLevel
    tCrit = Application.WorksheetFunction.T_Inv_2T(alpha, n - 1)
    marginError = tCrit * sd / Sqr(n)
    
    Dim result(1 To 2) As Variant
    result(1) = mean - marginError
    result(2) = mean + marginError
    CI_MEAN = result
    
    Exit Function
    
ErrorHandler:
    CI_MEAN = CVErr(xlErrValue)
End Function
