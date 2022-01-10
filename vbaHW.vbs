Sub TickerSymbol()
    
    ' column titles
    Range("H1").Value = "Ticker"
    Range("I1").Value = "Yearly Change"
    Range("J1").Value = "Percent Change"
    Range("K1").Value = "Total Stock Volume"
    
    ' creating variables
    Dim i As Double
    Dim stockVolume As Double
    Dim ticker As String
    Dim tickerLocation As Double
    Dim openingPrice As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    
    
    
    ' assigning initial values of variables
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    ticker = Cells(2, 1).Value
    stockVolume = 0
    i = 2
    tickerLocation = 2
    openingPrice = Range("C2").Value
    
    'start a while loop to go over all the table as long as column A is not empty
    Do While Not IsEmpty(Cells(i, 1))
        ' calculating the total stock volume for each ticker
        ' if we have the same ticker then increase the total volume
        If Cells(i, 1).Value = ticker Then
         stockVolume = stockVolume + Cells(i, 7).Value
        ' if found new ticker
        Else
            ' print previous ticker
            Cells(tickerLocation, 8).Value = ticker
            ' calculate and print yearly change
            Cells(tickerLocation, 9).Value = Cells(i - 1, 6).Value - openingPrice
            ' if yearly change is positive, make background green
            If Cells(tickerLocation, 9).Value > 0 Then
                Cells(tickerLocation, 9).Interior.ColorIndex = 4
            ' if yearly change is negative, make background red
            ElseIf Cells(tickerLocation, 9).Value < 0 Then
                Cells(tickerLocation, 9).Interior.ColorIndex = 3
            ' if yearly change is zero then keep background with no fill
            End If
            
            ' calculate and print percent change
            If openingPrice = 0 Then
                Cells(tickerLocation, 10).Value = 0
            Else
                Cells(tickerLocation, 10).Value = Cells(tickerLocation, 9).Value / openingPrice
            End If
            ' format the percentage change cell to 2 decimals and making it a percentage
            Cells(tickerLocation, 10).NumberFormat = "0.00%"
            
            '  if found a larger percent increase, update the stored ticker and value
            If Cells(tickerLocation, 10).Value > greatestPercentIncrease Then
                greatestPercentIncrease = Cells(tickerLocation, 10).Value
                greatestPercentIncreaseTicker = ticker
            End If
            
            '  if found a larger percent decrease, update the stored ticker and value
            If Cells(tickerLocation, 10).Value < greatestPercentDecrease Then
                greatestPercentDecrease = Cells(tickerLocation, 10).Value
                greatestPercentDecreaseTicker = ticker
            End If
            
            Cells(tickerLocation, 11).Value = stockVolume
            '  if found a larger total volume, update the stored ticker and value
            If greatestTotalVolume < Cells(tickerLocation, 11).Value Then
                greatestTotalVolume = Cells(tickerLocation, 11).Value
                greatestTotalVolumeTicker = ticker
            End If
            
            ' set and increment variables as needed
            ticker = Cells(i, 1).Value
            stockVolume = 0
            tickerLocation = tickerLocation + 1
            openingPrice = Cells(i, 3).Value
        End If
    
        i = i + 1

    Loop
    
    ' since we exit the while before printing the last ticker data, this is done here
    Cells(tickerLocation, 8).Value = ticker
    Cells(tickerLocation, 9).Value = Cells(i - 1, 6).Value - openingPrice
    Cells(tickerLocation, 10).Value = Cells(tickerLocation, 9).Value / openingPrice
    If Cells(tickerLocation, 10).Value > 0 Then
                Cells(tickerLocation, 9).Interior.ColorIndex = 4
            ElseIf Cells(tickerLocation, 9).Value < 0 Then
                Cells(tickerLocation, 9).Interior.ColorIndex = 3
            End If
    Cells(tickerLocation, 10).NumberFormat = "0.00%"
    Cells(tickerLocation, 11).Value = stockVolume
    
    ' print new table headers
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    ' print the greatest % increase
    Range("N2").Value = "Greatest % Increase"
    If Cells(tickerLocation, 10).Value > greatestPercentIncrease Then
        Range("O2").Value = ticker
        Range("P2").Value = Cells(tickerLocation, 10).Value
    Else
        Range("O2").Value = greatestPercentIncreaseTicker
        Range("P2").Value = greatestPercentIncrease
   End If
   Range("P2").NumberFormat = "0.00%"
     
    ' print the greatest % decrease
    Range("N3").Value = "Greatest % Decrease"
    If Cells(tickerLocation, 10).Value < greatestPercentDecrease Then
        Range("O3").Value = ticker
        Range("P3").Value = Cells(tickerLocation, 10).Value
    Else
        Range("O3").Value = greatestPercentDecreaseTicker
        Range("P3").Value = greatestPercentDecrease
    End If
    Range("P3").NumberFormat = "0.00%"
    
    ' print the greatest total volume
    Range("N4").Value = "Greatest Total Volume"
    If Cells(tickerLocation, 11).Value > greatestTotalVolume Then
        Range("O4").Value = ticker
        Range("P4").Value = Cells(tickerLocation, 11).Value
    Else
        Range("O4").Value = greatestTotalVolumeTicker
        Range("P4").Value = greatestTotalVolume
    End If

End Sub

Sub RunOnAllSheets()
    Dim sheet As Worksheet
    Application.ScreenUpdating = False
    For Each sheet In Worksheets
        sheet.Select
        Call TickerSymbol
    Next
    Application.ScreenUpdating = True
End Sub

