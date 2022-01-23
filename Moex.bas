Attribute VB_Name = "Moex"
Option Explicit

Private decimalSeparator As String

Public Function GetMoex(url$, ticker$, outDates_reference() As Date, outTimeSeries#(), lastDate As Date) As Boolean
    Dim resultFromMoex$
    Dim i%, j%, n%
    Dim tmpRows$(), tmpString$
    Dim objRequest
    Dim tmpDatumSplit$()
    
    decimalSeparator = Application.decimalSeparator
    
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", url, False
        .send
        .waitForResponse
        resultFromMoex = .responseText
    End With
    
    tmpRows = Split(resultFromMoex, Chr(10))
    
    For i = 0 To UBound(tmpRows)
        If InStr(1, tmpRows(i), """WAPRICE"": ") > 0 And InStr(1, tmpRows(i), """TRADEDATE"": ") > 0 Then
            ' Datum:
            tmpString = Split(Split(tmpRows(i), """TRADEDATE"": """)(1), """,")(0)
            tmpDatumSplit = Split(tmpString, "-")
            lastDate = DateSerial(tmpDatumSplit(0), tmpDatumSplit(1), tmpDatumSplit(2))
            
            tmpString = "null"
            
            ' CLOSE price
            If Split(Split(tmpRows(i), """CLOSE"": ")(1), ",")(0) <> "null" And Split(Split(tmpRows(i), """CLOSE"": ")(1), ",")(0) <> "" Then
                tmpString = Split(Split(tmpRows(i), """CLOSE"": ")(1), ",")(0)
            End If
            
            ' LEGALCLOSEPRICE
            If tmpString = "null" Then
                If Split(Split(tmpRows(i), """LEGALCLOSEPRICE"": ")(1), ",")(0) <> "null" And Split(Split(tmpRows(i), """LEGALCLOSEPRICE"": ")(1), ",")(0) <> "" Then
                    tmpString = Split(Split(tmpRows(i), """LEGALCLOSEPRICE"": ")(1), ",")(0)
                End If
            End If
            
            ' WAPRICE
            If tmpString = "null" Then
                If Split(Split(tmpRows(i), """WAPRICE"": ")(1), ",")(0) <> "null" And Split(Split(tmpRows(i), """WAPRICE"": ")(1), ",")(0) <> "" Then
                    tmpString = Split(Split(tmpRows(i), """WAPRICE"": ")(1), ",")(0)
                End If
            End If
            
            For j = 1 To UBound(outDates_reference)
                If lastDate = outDates_reference(j) Then
                    If tmpString = "null" Or tmpString = "" Then
                        outTimeSeries(j) = Undefined
                    Else
                        outTimeSeries(j) = CDbl(Replace(tmpString, ".", decimalSeparator))
                    End If
                    Exit For
                End If
            Next j
            n = n + 1
        End If
    Next i
    
    If n = 100 Then
        GetMoex = True
    Else
        GetMoex = False
    End If
End Function
