Attribute VB_Name = "Module1"
Sub stockanalysis()

Dim totalstockvolume As Double
Dim summarytickerindex As Double
Dim tickerstarterindex As Double
Dim yearlychange As Double
Dim percentchange As Double



 
    'use a for each loop to loop through all of the worksheets
    For Each ws In Worksheets
    
        totalstockvolume = 0
        summaryindexticker = 2
        tickerstarterindex = 2
        yearlychange = 0
        percentchange = 0
        
        
        ' get the count of rows
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
        ' Set column names and row names for summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Loop through all stockdatarows
        
        For Index = 2 To LastRow
        
            If ws.Cells(Index, 1).Value <> ws.Cells(Index + 1, 1).Value Then
            
                totalstockvolume = totalstockvolume + ws.Cells(Index, 7)
                yearlychange = ws.Cells(Index, 6) - ws.Cells(tickerstarterindex, 3)
                
                If ws.Cells(tickerstarterindex, 3) = 0 Then
                    percentchange = 0
                Else
                    percentchange = Round((yearlychange / ws.Cells(tickerstarterindex, 3) * 100), 2)
                End If
                
                ws.Range("L" & summaryindexticker).Value = totalstockvolume
                ws.Range("I" & summaryindexticker).Value = ws.Cells(Index, 1).Value
                ws.Range("J" & summaryindexticker).Value = yearlychange
                ws.Range("K" & summaryindexticker).Value = percentchange
                
                totalstockvolume = 0
                summaryindexticker = summaryindexticker + 1
                tickerstarterindex = Index
                
                
                
                
            Else
                totalstockvolume = totalstockvolume + ws.Cells(Index, 7)
            End If
            
            
           
        Next Index
        
        
        Next ws
    
    End Sub
    
