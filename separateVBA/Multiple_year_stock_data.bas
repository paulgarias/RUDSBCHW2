Attribute VB_Name = "Module1"
Option Explicit

Function search_all_stock_names(ws):

Dim tickerName() As Variant
Dim lastRow As Long 'This value is large, so we need long
Dim lastColumn As Integer
Dim tickerStr, prvTickerStr, fullTkrName As String
Dim i, iTkrName As Long


ReDim Preserve tickerName(0) As Variant

'Set ws = Worksheets("A")

prvTickerStr = ""
'For Each ws In Worksheets
lastRow = ws.UsedRange.Rows.Count
lastColumn = ws.UsedRange.Columns.Count

iTkrName = 0
For i = 2 To lastRow

        tickerStr = ws.Cells(i, 1).Value
        If (prvTickerStr <> tickerStr) Then
            
            tickerName(iTkrName) = tickerStr
            prvTickerStr = tickerStr
            iTkrName = iTkrName + 1
            ReDim Preserve tickerName(0 To UBound(tickerName) + 1) As Variant
        End If
        
Next i

'Next ws

ReDim Preserve tickerName(0 To UBound(tickerName))
search_all_stock_names = tickerName


End Function

Function search_stock_index(ws, stkStr, tickerNames):
    Dim ii, jj As Long
    
    For ii = 0 To UBound(tickerNames)
        If (ws.Cells(ii + 2, 9).Value = stkStr) Then Exit For
    Next ii
    
    search_stock_index = ii + 2
    
End Function

Function get_stock_volumes(ws, tickerNames):
    Dim i, j, k As Integer
    Dim msgstring As String
    Dim tickerLocRow As Long
    Dim tickerVolLocRow As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim stockName, prvStockName As String
    
    
    lastRow = ws.UsedRange.Rows.Count
    lastColumn = ws.UsedRange.Columns.Count
    
    For i = 2 To lastRow

        stockName = ws.Cells(i, 1).Value
        If (prvStockName <> stockName) Then
            'Get the index corresponding the the new stockName
            'j = search_stock_index()
            
            'Get the total value in the cell and add to the total
            'ws.Cells(j, 10).Value = ws.Cells(j, 10).Value + ws.Cells(i, 7)
            
        Else
            'Get the total value in the cell and add to the total
            ws.Cells(j, 10).Value = ws.Cells(j, 10).Value + ws.Cells(i, 7)
        End If
        
            
    Next i
    
    MsgBox msgstring
    get_stock_volumes = "hi"
End Function

Sub Financial():
    Dim tkrNames() As Variant
    Dim wks As Worksheet
    Dim i, j As Integer
    Dim dummy As String
    Dim lastRow As Long
    Dim valueDbl, startPrice, endPrice As Double
    Dim tickerStr As String
    

    
    'Set wks = Worksheets("A")
    
    For Each wks In Worksheets
    
        lastRow = wks.UsedRange.Rows.Count
        
        'Set the year start and end dates
        dummy = date_format(wks)
        
        'Get all the stocks in the sheet
        tkrNames = search_all_stock_names(wks)
    
        'And write out the stocks into a new section in each sheet
        wks.Cells(1, 9).Value = "Ticker"
        wks.Cells(1, 10).Value = "Ticker Stock Volume"
        wks.Cells(1, 11).Value = "Yearly Change"
        wks.Cells(1, 12).Value = "Percentage Change"
        
        For i = 0 To (UBound(tkrNames) - 1)
            wks.Cells(i + 2, 9).Value = tkrNames(i)
        Next i
        
        'Initialize the initial values to 0
        For i = 0 To (UBound(tkrNames) - 1)
            wks.Cells(i + 2, 10).Value = 0
        Next i
        
        j = search_stock_index(wks, tkrNames(0), tkrNames)
        
        startPrice = wks.Cells(2, 3).Value
        tickerStr = wks.Cells(2, 1).Value
        endPrice = wks.Cells(2, 6).Value
        
        For i = 3 To lastRow
            
            tickerStr = wks.Cells(i, 1).Value
            If (tkrNames(j - 2) <> tickerStr Or (i = lastRow)) Then
                'Update preceding stock data
                wks.Cells(j, 11).Value = endPrice - startPrice
                'Need to format the color in the column 11 (K)
                If (wks.Cells(j, 11).Value >= 0) Then
                    wks.Cells(j, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    wks.Cells(j, 11).Interior.Color = RGB(255, 0, 0)
                    
                End If
                
                If (startPrice > 0) Then
                    wks.Cells(j, 12).Value = (endPrice - startPrice) / startPrice
                    wks.Cells(j, 12).NumberFormat = "0.00%"
                Else
                    wks.Cells(j, 12).Value = 0
                    wks.Cells(j, 12).NumberFormat = "0.00%"
                End If
                
                'Because this is a new symbol, we need to update the start price
                startPrice = wks.Cells(i, 3).Value
                
                j = search_stock_index(wks, tickerStr, tkrNames)
                wks.Cells(j, 10).Value = wks.Cells(j, 10).Value + wks.Cells(i, 7).Value
                
            Else
            
                wks.Cells(j, 10).Value = wks.Cells(j, 10).Value + wks.Cells(i, 7).Value
                
            End If
            endPrice = wks.Cells(i, 6).Value
            
        Next i
        
        wks.Range("O2").Value = "Greatest % increase"
        wks.Range("O3").Value = "Greatest % Decrease"
        wks.Range("O4").Value = "Greatest total volume"
        
        wks.Range("P1").Value = "Ticker"
        wks.Range("Q1").Value = "Value"
        
        wks.Range("Q2").Value = wks.Cells(get_max_percent(wks), 12).Value
        wks.Range("Q2").NumberFormat = "0.00%"
        wks.Range("Q3").Value = wks.Cells(get_min_percent(wks), 12).Value
        wks.Range("Q3").NumberFormat = "0.00%"
        wks.Range("Q4").Value = wks.Cells(get_max_volume(wks), 10).Value
        
        wks.Range("P2").Value = wks.Cells(get_max_percent(wks), 9).Value
        wks.Range("P3").Value = wks.Cells(get_min_percent(wks), 9).Value
        wks.Range("P4").Value = wks.Cells(get_max_volume(wks), 9).Value
        
    Next wks
    
    
End Sub

Function date_format(ws):
'MsgBox Cells(2, 2).Value
Dim lastRow As Long
Dim wks As Worksheet
Dim Cr, Dr, C As Range
Dim strRange As String
Dim maxD, minD As Double

lastRow = ws.UsedRange.Rows.Count

Set Cr = ws.Range("B2:B" & lastRow)
'Set Dr = ws.Range("H2:H" & lastRow)

'With Dr
'.FormulaR1C1 = "=TEXT(RC[-6],""0000-00-00"")+0"
'.NumberFormat = "mm/dd/yyyy"
'End With

maxD = Application.WorksheetFunction.Max(Cr)
minD = Application.WorksheetFunction.Min(Cr)

ws.Range("T1") = maxD
ws.Range("U1") = minD

'Get the maxium and mimum ranges for the year and place them in the cells T1 and U1 in each worksheet
With ws.Range("T1:U1")
    .NumberFormat = "yyyymmdd"
End With

date_format = "Success"

End Function

Sub reset_H()

Dim wksts As Worksheet

For Each wksts In Worksheets
    wksts.Range("H:H").Value = ""

Next wksts



End Sub



Function get_max_volume(ws):
'Get max in column 10
Dim i, j, k, l As Integer
Dim lastRow As Long
Dim maxValdbl As Double


'Dim wks As Worksheet

'Set wks = Worksheets("A")

lastRow = ws.UsedRange.Rows.Count
'Start with the first value
i = 2
j = 2
maxValdbl = ws.Cells(2, 10).Value
For i = 2 To lastRow
    If (ws.Cells(i + 1, 10).Value = "") Then
        Exit For
    Else
        If (ws.Cells(i + 1, 10).Value > maxValdbl) Then
            maxValdbl = ws.Cells(i + 1, 10).Value
            j = i + 1
        End If
        
    End If
    
Next i

get_max_volume = j

End Function

Function get_max_percent(ws):
'Get max in column 10
Dim i, j, k, l As Integer
Dim lastRow As Long
Dim maxValdbl As Double


'Dim wks As Worksheet

'Set wks = Worksheets("A")

lastRow = ws.UsedRange.Rows.Count
'Start with the first value
i = 2
j = 2
k = 12
maxValdbl = ws.Cells(2, k).Value
For i = 2 To lastRow
    If (ws.Cells(i + 1, k).Value = "") Then
        Exit For
    Else
        If (ws.Cells(i + 1, k).Value > maxValdbl) Then
            maxValdbl = ws.Cells(i + 1, k).Value
            j = i + 1
        End If
        
    End If
    
Next i

get_max_percent = j

End Function

Function get_min_percent(ws):
'Get max in column 10
Dim i, j, k, l As Integer
Dim lastRow As Long
Dim minValdbl As Double


'Dim wks As Worksheet

'Set wks = Worksheets("A")

lastRow = ws.UsedRange.Rows.Count
'Start with the first value
i = 2

j = 2
k = 12
minValdbl = ws.Cells(2, k).Value
For i = 2 To lastRow
    If (ws.Cells(i + 1, k).Value = "") Then
        Exit For
    Else
        If (ws.Cells(i + 1, k).Value < minValdbl) Then
            minValdbl = ws.Cells(i + 1, k).Value
            j = i + 1
        End If
        
    End If
    
Next i

get_min_percent = j

End Function



