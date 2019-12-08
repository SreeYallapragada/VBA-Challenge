Attribute VB_Name = "Module1"
Sub vba_challenge2()
   
   'Declare the variables for the dataset
   Dim ticker As String
   Dim openingPrice As Double
   Dim closingPrice As Double
   Dim yearlyChange As Double
   Dim percentChange As Double
   Dim totalStockVolume As Double
   Dim tickerCounter As Integer
   
   'Declare variables indicating the first/last position of a particular ticker symbol
   Dim firstTickerPosition As Double
   Dim lastTickerPosition As Double
   
   'Initialize the total stock volume
   totalStockVolume = 0
   
   'Loop through all of the tabs in the data set
   For Each ws In Worksheets
     
   'Determine the last row of each worksheet
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   LastRow = LastRow + 1
   
   'Determine the last column to place headers for the ticker/total stock volume results
   LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
   
   'Add header titles to the columns in all worksheets
   ws.Cells(1, LastColumn + 2).Value = "Ticker"
   ws.Cells(1, LastColumn + 3).Value = "Opening Price"
   ws.Cells(1, LastColumn + 4).Value = "Yearly Closing Price"
   ws.Cells(1, LastColumn + 5).Value = "Yearly Change"
   ws.Cells(1, LastColumn + 6).Value = "Percent Change"
   ws.Cells(1, LastColumn + 7).Value = "Total Stock Volume"
   
   'Initialize the ticker counter to start analyzing the ticker symbols
   tickerCounter = 2
   
   'Loop through all the rows to identify each ticker symbol and its stock volume
   For i = 2 To LastRow
   
        'Check each row to see if the current ticker value is different than the previous ticker value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Register the value of the last ticker position
            lastTickerPosition = i
        
            'Assign the value of the ticker on the current row to the String variable "ticker"
            ticker = ws.Cells(i, 1).Value
        
            'Write the value of the current total volume for the first ticker symbol
            totalStockVolume = totalStockVolume + ws.Cells(i, LastColumn).Value
    
                  'Initialize the following variables if the total stock volume is > 0
                  If totalStockVolume > 0 Then
         
                  openingPrice = ws.Cells(firstTickerPosition, 3).Value
                  closingPrice = ws.Cells(lastTickerPosition, 6).Value
                  yearlyChange = closingPrice - openingPrice
                  percentChange = yearlyChange / openingPrice
     
                  Else
         
                 'Total stock volume would be equal to zero if the volume was < 0
                  openingPrice = 0
                  closingPrice = 0
                  yearlyChange = 0
                  percentChange = 0
         
                  End If
         
            'Determine the position to place the first ticker and its corresponding stock volume
            ws.Cells(tickerCounter, LastColumn + 2).Value = ticker
        
            'Display the opening price
            ws.Cells(tickerCounter, LastColumn + 3).Value = openingPrice
        
            'Display the closing price
            ws.Cells(tickerCounter, LastColumn + 4).Value = closingPrice
        
            'Display the yearly change
            ws.Cells(tickerCounter, LastColumn + 5).Value = yearlyChange
        
                  'Color the percent chance cells green if there is positive change
                  If yearlyChange > 0 Then
                
                        ws.Cells(tickerCounter, LastColumn + 5).Interior.ColorIndex = 4
                
                  'Color the cells red if the change is negative
                  ElseIf yearlyChange < 0 Then
                
                        ws.Cells(tickerCounter, LastColumn + 5).Interior.ColorIndex = 3
                
                  End If
        
          'Display the percent change and make it a percent
          ws.Cells(tickerCounter, LastColumn + 6).Value = percentChange
        
          ws.Cells(tickerCounter, LastColumn + 6).NumberFormat = "0.00%"
        
          'Put the value of the total stock volume
          ws.Cells(tickerCounter, LastColumn + 7).Value = totalStockVolume
        
          'Reset the total stock volume to 0 when the next ticker symbol cell is different
          totalStockVolume = 0
        
          'Reset the first ticker position value when the next ticker symbol is different
          firstTickerPosition = 0
        
          'Redefine the current ticker counter to the next ticker counter
          tickerCounter = tickerCounter + 1
        
          
          'Now, update the total stock volume for the next ticker symbol
   Else
        
        If ws.Cells(i, 3).Value > 0 Then
       
            If firstTickerPosition = 0 Then
        
            firstTickerPosition = i
    
            End If
        
        End If
        
        'Keep a sum of the total stock volume for each day of trading for each ticker
        totalStockVolume = totalStockVolume + ws.Cells(i, LastColumn).Value
    
   End If

   Next i
   
   
   'Determine the last column to place headers for the ticker/total stock volume results
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'The new column after the results section
    FirstColumn = LastColumn - 5
    
        With ActiveSheet
        
        LastRow = .Cells(.Rows.Count, "N").End(xlUp).Row
          
        End With
    
    
    'Place titles on each of the new columns
    ws.Cells(1, LastColumn + 2).Value = "Ticker"
    ws.Cells(1, LastColumn + 3).Value = "Value"
    ws.Cells(2, LastColumn + 1).Value = "Greatest % Increase"
    ws.Cells(3, LastColumn + 1).Value = "Greatest % Decrease"
    ws.Cells(4, LastColumn + 1).Value = "Greatest Total Volume"
    
    'Declare the variables for the greatest increase, decrease, and total volume
    Dim greatestStockInc As String
    Dim greatestStockDec As String
    Dim greatestTotalVol As String
    
    'Declare the numerical values of stock comparison
    Dim greatIncrease As Double
    Dim greatDecrease As Double
    Dim greatVolume As Double
    
    'Initialize the numerical values to zero
    greatIncrease = 0
    greatDecrease = 0
    greatVolume = 0
        
        'Loop through the entire dataset
        For i = 2 To LastRow
        
            'Look for the value that has the greatest % increase
            If ws.Cells(i, LastColumn - 1).Value > greatIncrease Then
             
                greatIncrease = ws.Cells(i, LastColumn - 1).Value
         
                'Note down its corresponding ticker symbol
                greatestStockInc = ws.Cells(i, LastColumn - 5)
            
            End If
        
            'Look for the value with the greatest % decrease
            If ws.Cells(i, LastColumn - 1) < greatDecrease Then
             
                greatDecrease = ws.Cells(i, LastColumn - 1).Value
             
                'Write its corresponding ticker symbol
                greatestStockDec = ws.Cells(i, LastColumn - 5).Value
        
            End If
        
            'Look for the stock with the greatest total volume
            If ws.Cells(i, LastColumn).Value > greatVolume Then
                
                greatVolume = ws.Cells(i, LastColumn).Value
                
                'Write the ticker symbol that it corresponds to
                greatestTotalVol = ws.Cells(i, LastColumn - 5).Value
            
            End If
       
        Next i
            
    'Write the ticker of greatest % increase in the new cells and format the number to %
     ws.Cells(2, LastColumn + 2).Value = greatestStockInc
     ws.Cells(2, LastColumn + 3).Value = greatIncrease
    
     ws.Cells(2, LastColumn + 3).NumberFormat = "0.00%"
            
 
    'Write the ticker of greatest % decrease and make it a %
     ws.Cells(3, LastColumn + 2).Value = greatestStockDec
     ws.Cells(3, LastColumn + 3).Value = greatDecrease
             
     ws.Cells(3, LastColumn + 3).NumberFormat = "0.00%"
            
            
    'Write the ticker with the greatest total volume
     ws.Cells(4, LastColumn + 2).Value = greatestTotalVol
     ws.Cells(4, LastColumn + 3).Value = greatVolume
     
     ws.Cells(4, LastColumn + 3).NumberFormat = "0"
    
    
    Next ws

End Sub
