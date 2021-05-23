Sub stock_summary()

Dim Last_Row As Long
Dim Iters As Integer
Dim Current_Ticker As String
Dim Next_Ticker As String
Dim Result_Table_Row As Integer
Dim Percent_Change, Price_Diff, Open_Price, Close_Price As Double
Dim Sheet_Count As Integer
Dim Stock_Volume As Variant

Sheet_Count = Sheets.Count

For W = 1 To Sheet_Count
    
Sheets(W).Activate

With Sheets(W)




Last_Row = Cells(Rows.Count, 1).End(xlUp).Row ' Last Row in dataset
Iters = -1 ' iteration counter (Excluding current row)
Results_Table_Row = 2 ' The starting row of our results table
Stock_Volume = 0 ' initializing stock volume counter




        
    
    
For R = 2 To Last_Row
            
Current_Ticker = Cells(R, 1).Value
Next_Ticker = Cells(R + 1, 1).Value
Iters = Iters + 1
            
Stock_Volume = Stock_Volume + Cells(R, 7).Value
            
If (Current_Ticker <> Next_Ticker) Then
            
            
' Adding the ticker symbol to results table
            
Cells(Results_Table_Row, 10).Value = Current_Ticker
            
' Calculating the price difference for each stock
                
Open_Price = Cells(R - Iters, 3).Value
Close_Price = Cells(R, 6).Value
Price_Diff = Close_Price - Open_Price
Cells(Results_Table_Row, 11).Value = Price_Diff
                
If (Price_Diff < 0) Then
                    
Cells(Results_Table_Row, 11).Interior.ColorIndex = 3
                    
Else
                    
Cells(Results_Table_Row, 11).Interior.ColorIndex = 4
                    
End If
                     
' Calculating Percent Change for each stock
            
Percent_Change = (Price_Diff / Open_Price)
Cells(Results_Table_Row, 12).Value = Percent_Change
Cells(Results_Table_Row, 12).NumberFormat = "00.00%"
                
' Setting the stock_volume for each stock
                
Cells(Results_Table_Row, 13).Value = Stock_Volume
                
                
' Resetting important variables
                
Cells(Results_Table_Row, 11).Value = Price_Diff '
Iters = -1
Results_Table_Row = Results_Table_Row + 1
Stock_Volume = 0
            
            
End If
            
Next R
           
End With

Next W

    

End Sub











