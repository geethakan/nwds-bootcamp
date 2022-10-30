'Sub procedure to calculate yearly price difference and total volume of tickers
Sub StockReport():

    'Declare variables
    Dim ws As Worksheet
    Dim startWsName As String
    
    Dim r As Long           'row
    Dim rr As Long          'result row
    Dim rPrv As Long        'Previous row
    Dim tikr As String
    Dim oldtikr As String
    Dim tikrChg As Boolean
    
    'Variables for getting final result from all sheets
    Dim highPer As Double
    Dim lowPer As Double
    Dim highVol As Double
    Dim tempDbl As Double
    Dim tempLng As Double
    Dim hpTkr As String
    Dim lpTkr As String
    Dim hvTkr As String
    Dim highvolStr As String
    
    'Set starting worksheet to position to in the very end
    startWsName = Application.ActiveSheet.name
        
    'Loop thru all sheets
     For Each ws In ThisWorkbook.Worksheets
                        
        'Set last row count for first column; initialize; Activate sheet
        ws.Activate
        lRow = Cells(Rows.Count, 1).End(xlUp).Row
        InitFtn
        
        highPer = 0
        lowPer = 0
        tempDbl = 0
        highVol = 0
        tempLng = 0
        
                
        'Loop thru all rows of the current sheet;
        For r = 2 To lRow
        
            tikr = Cells(r, 1).Value
            If r = 2 Then
                rr = r
                oldtikr = Cells(r, 1).Value
                Cells(r, 9).Value = Cells(r, 1).Value
                Cells(r, 23).Value = Cells(r, 2).Value
                Cells(r, 24).Value = Cells(r, 3).Value
            End If
            
            If tikr = oldtikr Then
                tikrChg = False
            Else
                tikrChg = True
            End If
        
            
            If tikrChg = True Then
                'old ticker yearly change and % calculated
                rPrv = r - 1
                Cells(rr, 25).Value = Cells(rPrv, 2).Value
                Cells(rr, 26).Value = Cells(rPrv, 6).Value
                Cells(rr, 10).Value = Cells(rr, 26).Value - Cells(rr, 24).Value
                Cells(rr, 10).NumberFormat = "#0.00"
                If Cells(rr, 10).Value > 0 Then
                    Cells(rr, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(rr, 10).Interior.Color = RGB(255, 0, 0)
                End If
                Cells(rr, 11).Value = Cells(rr, 10).Value / Cells(rr, 24).Value
                Cells(rr, 11).NumberFormat = "0.00%"
                
                '============================================='
                'Capture high low values
                tempDbl = CDbl(Cells(rr, 11).Value)
                If tempDbl > highPer Then
                    highPer = tempDbl
                    hpTkr = Cells(rr, 9).Value
                End If
                
                If tempDbl < lowPer Then
                    lowPer = tempDbl
                    lpTkr = Cells(rr, 9).Value
                End If
                
                tempLng = CDbl(Cells(rr, 12).Value)
                If tempLng > highVol Then
                    highVol = tempLng
                    hvTkr = Cells(rr, 9).Value
                End If
                '============================================='
                
                'Set current value for oldtikr; write new tikr in results
                rr = rr + 1
                oldtikr = Cells(r, 1).Value
                Cells(rr, 9).Value = Cells(r, 1).Value
                Cells(rr, 23).Value = Cells(r, 2).Value
                Cells(rr, 24).Value = Cells(r, 3).Value
                Cells(rr, 12).Value = Cells(r, 7).Value
            Else
                'Add to TotVolume
                Cells(rr, 12).Value = Cells(rr, 12) + Cells(r, 7).Value
            End If
            
            'If last row, complete results before ending loop
            If r = lRow Then
            
                Cells(rr, 25).Value = Cells(r, 2).Value
                Cells(rr, 26).Value = Cells(r, 6).Value

                Cells(rr, 10).Value = Cells(rr, 26).Value - Cells(rr, 24).Value
                Cells(rr, 10).NumberFormat = "#0.00"
                If Cells(rr, 10).Value > 0 Then
                    Cells(rr, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(rr, 10).Interior.Color = RGB(255, 0, 0)
                End If
                Cells(rr, 11).Value = Cells(rr, 10).Value / Cells(rr, 24).Value
                Cells(rr, 11).NumberFormat = "0.00%"
                
                '============================================='
                'Capture high low values
                tempDbl = CDbl(Cells(rr, 11).Value)
                If tempDbl > highPer Then
                    highPer = tempDbl
                    hpTkr = Cells(rr, 9).Value
                End If
                
                If tempDbl < lowPer Then
                    lowPer = tempDbl
                    lpTkr = Cells(rr, 9).Value
                End If
                
                tempLng = CDbl(Cells(rr, 12).Value)
                If tempLng > highVol Then
                    highVol = tempLng
                    hvTkr = Cells(rr, 9).Value
                End If
                '============================================='

            End If
    
        Next r
        
        'Set final values for the sheet here
        Range("P2").Value = hpTkr
        Range("P3").Value = lpTkr
        Range("P4").Value = hvTkr
        Range("Q2").Value = highPer
        Range("Q3").Value = lowPer
        highvolStr = Format(Str(highVol), "0.00E+00")
        Range("Q4").Value = highvolStr
    
        Range("Q1").EntireColumn.AutoFit
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        
        'Clear columns w,x,y and z
        Range("W1", Range("W1").End(xlDown)).Clear
        Range("X1", Range("X1").End(xlDown)).Clear
        Range("Y1", Range("Y1").End(xlDown)).Clear
        Range("Z1", Range("Z1").End(xlDown)).Clear
            
    Next ws
    
    Worksheets(startWsName).Activate
    
    
End Sub
'================================================================================================='
Function InitFtn():

    Range("I1", Range("I1").End(xlDown)).Clear
    Range("J1", Range("J1").End(xlDown)).Clear
    Range("K1", Range("K1").End(xlDown)).Clear
    Range("L1", Range("L1").End(xlDown)).Clear
    Range("W1", Range("W1").End(xlDown)).Clear
    Range("X1", Range("X1").End(xlDown)).Clear
    Range("Y1", Range("Y1").End(xlDown)).Clear
    Range("Z1", Range("Z1").End(xlDown)).Clear
    
    'Set Headers and adjust width
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("J1").EntireColumn.AutoFit
    Range("K1").EntireColumn.AutoFit
    Range("L1").EntireColumn.AutoFit
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("O1").EntireColumn.AutoFit
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Setting these for testing and verification
    Range("W1").Value = "Open Date"
    Range("X1").Value = "Open Value"
    Range("Y1").Value = "Close Date"
    Range("Z1").Value = "Close Value"

End Function
