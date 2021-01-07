# VBA_Stock_Market_Analysis
VBA scripting to analyze real stock market data


Sub Wallstreetbets()

    Dim CurrentWS As Worksheet

    For Each CurrentWS In Worksheets

        'Variables
        Dim TickerVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim DeltaPrice As Double
        Dim DeltaPercent As Double
        TickerVolume = 0
        OpenPrice = 0
        ClosePrice = 0
        DeltaPrice = 0
        DeltaPercent = 0
        summarytablerow = 2
        
        
        
        Dim GreatestIncreaseTicker As String
        GreatestIncreaseTicker = " "
        Dim GreatestDecreaseTicker As String
        GreatestIncreaseTicker = " "
        Dim GreatestTickerVolume As String
        GreatestTickerVolume = " "
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestVolume As Double
        GreatestVolume = 0

    
        lastrow = CurrentWS.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        'Headers for the Summary Table Summary Analysis
        CurrentWS.Cells(1, 9).Value = "Ticker"
        CurrentWS.Cells(1, 10).Value = "Yearly Change"
        CurrentWS.Cells(1, 11).Value = "Percent Change"
        CurrentWS.Cells(1, 12).Value = "Total Stock Volume"
        CurrentWS.Cells(1, 16).Value = "Ticker"
        CurrentWS.Cells(1, 17).Value = "Value"
        CurrentWS.Cells(2, 15).Value = "Greatest % Increase"
        CurrentWS.Cells(3, 15).Value = "Greatest % Decrease"
        CurrentWS.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'Set initial value of the OpenPrice for the first ticker
        OpenPrice = CurrentWS.Cells(2, 3).Value
        
        'Loop through each row starting with row 2
        For Row = 2 To lastrow
        
            'If the next ticker doesnt match...
            If CurrentWS.Cells(Row + 1, 1).Value <> CurrentWS.Cells(Row, 1).Value Then
            
                '**********************************************************************************************************************
                'SUMMARY TABLE
            
                'Insert each TickerName to the Summary Table
                TickerName = CurrentWS.Cells(Row, 1).Value
                CurrentWS.Range("I" & summarytablerow).Value = TickerName
            
            
                'Calculate the Yearly Change by getting the delta from OpenPrice(value set outisde the loop) and ClosePrice
                ClosePrice = CurrentWS.Cells(Row, 6).Value
                DeltaPrice = ClosePrice - OpenPrice
                'Then insert into the Summary Table
                CurrentWS.Range("J" & summarytablerow).Value = DeltaPrice
            
            
                'Calculate the Percentage Change by getting the delta from
                If OpenPrice <> 0 Then
                DeltaPercent = (DeltaPrice / OpenPrice)
                Else
                DeltaPercent = 0
                End If
                'Then insert into the Summary Table
                CurrentWS.Range("K" & summarytablerow).Value = Format(DeltaPercent, "0.00%")
                

                'Add all the Total Stock Volume
                TickerVolume = CurrentWS.Cells(Row, 7).Value + TickerVolume  '<< this code will only capture the last TotalStockVolume of the ticker symbol. To add, write this code after the Else statement
                'Then insert to Summary Table
                CurrentWS.Range("L" & summarytablerow).Value = TickerVolume

                                
                'Color format the Percent Change
                If DeltaPrice > 0 Then
                 CurrentWS.Range("J" & summarytablerow).Interior.ColorIndex = 4
                ElseIf DeltaPrice <= 0 Then
                 CurrentWS.Range("J" & summarytablerow).Interior.ColorIndex = 3
                End If
                
                
                'To insert into the next row of the Summary Table
                summarytablerow = summarytablerow + 1
                
                'Reset the ClosingPrice to restart the calculation
                ClosePrice = 0
                'Then calculation starts again by having a new OpeningPrice for the next ticker (value set inside the loop)
                OpenPrice = CurrentWS.Cells(Row + 1, 3).Value
                'Reset the DeltaPercent and TickerVolume after the Summary Analysis
                
                '**********************************************************************************************************************
                'SUMMARY ANALYSIS
                
                'Calculate the highest and lowest Percentage change value
                If DeltaPercent > GreatestIncrease Then
                 GreatestIncrease = DeltaPercent
                 GreatestIncreaseTicker = TickerName
                ElseIf DeltaPercent < GreatestDecrease Then
                 GreatestDecrease = DeltaPercent
                 GreatestDecreaseTicker = TickerName
                End If
                
                'Calculate the greatest volume
                If TickerVolume > GreatestVolume Then
                 GreatestVolume = TickerVolume
                 GreatestTickerVolume = TickerName
                End If
                 
                'Reset the DeltaPercent and TickerVolume to restart the calculation
                DeltaPercent = 0
                TickerVolume = 0
                
                '**********************************************************************************************************************
                
            Else
                'to keep adding the Tickervolume
                TickerVolume = CurrentWS.Cells(Row, 7).Value + TickerVolume
   
            End If
            
        Next Row
                
                'Insert the Summary Analysis result outside the loop since the rows and columns are set
                CurrentWS.Range("P2").Value = GreatestIncreaseTicker
                CurrentWS.Range("P3").Value = GreatestDecreaseTicker
                CurrentWS.Range("P4").Value = GreatestTickerVolume
                CurrentWS.Range("Q2").Value = GreatestIncrease
                CurrentWS.Range("Q3").Value = GreatestDecrease
                CurrentWS.Range("Q4").Value = GreatestVolume
        
        
        'Auto-adjust the column width of the worksheet
        CurrentWS.Columns.AutoFit

    Next CurrentWS

End Sub