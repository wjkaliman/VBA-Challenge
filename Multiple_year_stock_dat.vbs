Attribute VB_Name = "Module1"
'1 Create a script that will loop through all the stocks for one year
'and output the following information
'2 The Ticker symbol
'3 Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'4 The percent change from opening price at the beginning of a given year to the closing price at the end of that year
'5 The Total Stock volume of the stock
'6 You should also have conditional formatting that will highlight positive change in green and negative change in red

Sub Multiyear_ticker()

    'Q1 Set all initial varibles
    Dim LastRow As Long
    
    'Q2 create a ticker Symbol
    Dim TickerSymbol As String
    
    ' how many unique tickers we find and also our row counter in our summary table
    Dim count As Integer
    ' Q3 this is a double because of the decimals
    Dim openprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    ' Q4 Set initial varible for percent change
    Dim percentchange As Double
    ' Q5 set initial value for holding the total stock volumne
    Dim TotalStockVolumne As Double
    
    TotalStockVolumne = 0
    
    ' Q5 tracking location of each ticker and add up the total stock volumne
    Dim SummaryTable As Integer
    SummaryTable = 2
    
    Dim Start As Long
    Start = 2
    
    'Q2 this means its starting on row 2
    count = 2

    'Q1 in column A what is the last row? part answer to Q1 before loop
    LastRow = Cells(Rows.count, 1).End(xlUp).Row
    'LastRow = 70926

    ' 1)Loop through all tickers in column A
    For currentRow = 2 To LastRow
    
        TotalStockVolumne = TotalStockVolumne + Cells(currentRow, 7)

        '   Check if we are still within the same ticker
        If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then
        
        
            'Q2 Set TickerSymbol because we know to when it changes
            TickerSymbol = Cells(currentRow, 1).Value
            'location of output of symbol
            Cells(count, 9) = TickerSymbol
           
            
            'Q3 setting location of closingprice
            closingprice = Cells(currentRow, 6).Value
            
            'Q3 need opening price
            openprice = Cells(Start, 3).Value
            ' reset start to be opening price for each ticker
            Start = currentRow + 1
            ' this is the change point for A to AA
            
            
            'Q3 yearly change
            yearlychange = closingprice - openprice
            Cells(count, 10) = yearlychange
            
            
            'Q4 percent change
            percentchange = 0
            
            If openprice > 0 Then
            
                percentchange = (closingprice - openprice) / openprice
            
            End If
            
            Cells(currentRow, 11).Value = percentchange
            Range("K" & SummaryTable).Value = percentchange
            
            'Q5 the total stock volume of a stock
            
            
            Range("L" & SummaryTable).Value = TotalStockVolumne
            TotalStockVolumne = 0
            
        'Where do I put this conditional Formatting
            If Cells(SummaryTable, 10).Value > 0 Then
                Cells(SummaryTable, 10).Interior.ColorIndex = 4 'green
            ElseIf Cells(SummaryTable, 10).Value < 0 Then
                Cells(SummaryTable, 10).Interior.ColorIndex = 3 'red
            End If
        
            count = count + 1
            SummaryTable = SummaryTable + 1
        
        End If
    
    Next currentRow
    
    MsgBox ("code is finished!")
    
    

End Sub


