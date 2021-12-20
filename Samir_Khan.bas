Attribute VB_Name = "Samir_Khan"
Option Explicit

'this code was created by:
'Samir Khan
'simulationconsultant@gmail.com
'The latest version of this spreadsheet can be downloaded from http://investexcel.net/multiple-stock-quote-downloader-for-excel/
'Please link to http://investexcel.net if you like this spreadsheet

Public crumb$
Public cookie$
Private decimalSeparator As String

Public Sub getYahooFinanceData(stockTicker$, startDate$, endDate$, _
        frequency$, outDates() As Date, outTimeSeries#())
    Dim resultFromYahoo$
    Dim objRequest As Variant
    Dim csv_rows$()
    Dim iRows&
    Dim CSV_Fields As Variant
    Dim iCols&
    Dim tickerURL$
    
    decimalSeparator = Application.decimalSeparator
    
    'Construct URL
    '***************************************************
    tickerURL = "https://query1.finance.yahoo.com/v7/finance/download/" & stockTicker & _
        "?period1=" & startDate & _
        "&period2=" & endDate & _
        "&interval=" & frequency & "&events=history" & "&crumb=" & crumb
    '***************************************************
               
    'Get data from Yahoo
    '***************************************************
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .setRequestHeader "Cookie", cookie
        .send
        .waitForResponse
        resultFromYahoo = .responseText
    End With
    '***************************************************
        
    'Parse returned string into an array
    '***************************************************
    csv_rows() = Split(resultFromYahoo, Chr(10))
    ReDim outDates(UBound(csv_rows))
    ReDim outTimeSeries(UBound(csv_rows), 4)
    
    For iRows = LBound(csv_rows) + 1 To UBound(csv_rows) ' ignore first row with index 0
        CSV_Fields = Split(csv_rows(iRows), ",")
        
        outDates(iRows) = CDate(CSV_Fields(0))
        
        For iCols = LBound(CSV_Fields) + 1 To 4
            If IsNumeric(CSV_Fields(iCols)) Then
                outTimeSeries(iRows, iCols) = Replace(CSV_Fields(iCols), ".", decimalSeparator)
            Else
                outTimeSeries(iRows, iCols) = Undefined
            End If
        Next
    Next
End Sub

Public Function getCookieCrumb() As Boolean
    Dim i&
    Dim crumbStartPos&
    Dim crumbEndPos&
    Dim objRequest
    
    If crumb <> "" Then
        Exit Function
    End If
    
    getCookieCrumb = False
    
    For i = 0 To 5  'ask for a valid crumb 5 times
        Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        With objRequest
            .Open "GET", "https://finance.yahoo.com/lookup?s=bananas", False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
            .send
            .waitForResponse (10)
            cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
            crumbStartPos = InStrRev(.responseText, """crumb"":""") + 9
            crumbEndPos = crumbStartPos + 11 'InStr(crumbStartPos, .ResponseText, """", vbBinaryCompare)
            crumb = Mid(.responseText, crumbStartPos, crumbEndPos - crumbStartPos)
        End With
        
        If Len(crumb) = 11 Then 'a valid crumb is 11 characters long
            getCookieCrumb = True
            Exit For
        End If:
    Next i
End Function
