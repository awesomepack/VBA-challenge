Sub stock_summary()

Dim Last_Row As Long
Dim Iters As Integer
Dim Current_Ticker As String
Dim Next_Ticker As String
Dim Result_Table_Row As Integer
Dim Percent_Change, Price_Diff, Open_Price, Close_Price As Double
Dim Sheet_Count As Integer
Dim Stock_Volume As Variant



'Code to deal with a multi sheet workbook

Sheet_Count = Sheets.Count 'Determining the the number of sheets present in the workbook

For W = 1 To Sheet_Count  'iterating over worksheets in workbook
    
Sheets(W).Activate  

With Sheets(W) 



'defining variables for future reference
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row ' Last Row in dataset
Iters = -1 ' iteration counter (Excluding current row)
Results_Table_Row = 2 ' The starting row of our results table
Stock_Volume = 0 ' initializing stock volume counter


'Header Values for the results/summary table

Cells(1, 10).Value = "Ticker" 
Cells(1, 11).Value = "Year_Change" 
Cells(1, 12).Value = "Pct_Change" 
Cells(1, 13).Value = "Volume" 



'Code creating results table for data in a given worksheet



For R = 2 To Last_Row 'Looping through the data in a given worksheet


'Variable Assignment
Current_Ticker = Cells(R, 1).Value  'R'th ticker value
Next_Ticker = Cells(R + 1, 1).Value  '(R+1)'th ticker value
Iters = Iters + 1 'Number of iterations since stock change
Stock_Volume = Stock_Volume + Cells(R, 7).Value 'Stock volume accumulator
            

'Noting a change in ticker symbol , code summarizes stock data
If (Current_Ticker <> Next_Ticker) Then 'True when ticker symbols are not identical
            
Cells(Results_Table_Row, 10).Value = Current_Ticker 'Add ticker value to results table
            

'Calculating yearly price difference and inserting into results table      
Open_Price = Cells(R - Iters, 3).Value 'Row with open price is last Row minus the number of iterations before a ticker change
Close_Price = Cells(R, 6).Value
Price_Diff = Close_Price - Open_Price
Cells(Results_Table_Row, 11).Value = Price_Diff

'Conditonal formatting for yearly change             
If (Price_Diff < 0) Then
                    
Cells(Results_Table_Row, 11).Interior.ColorIndex = 3
                    
Else
                    
Cells(Results_Table_Row, 11).Interior.ColorIndex = 4
                    
End If


                     
' Calculating Percent Change for each stock
'conditional present to avoid overflow error induced by dividing by zero
If Open_Price = 0 Then

Percent_Change = 0
Cells(Results_Table_Row, 12).Value = Percent_Change 
Cells(Results_Table_Row, 12).NumberFormat = "00.00%"

Else

Percent_Change = (Close_Price / Open_Price)
Cells(Results_Table_Row, 12).Value = Percent_Change
Cells(Results_Table_Row, 12).NumberFormat = "00.00%"

End If


                
'Inserting accumulated stock volume for current ticker              
Cells(Results_Table_Row, 13).Value = Stock_Volume
                
                
' Reinitializing variables for the next ticker symbol     
Cells(Results_Table_Row, 11).Value = Price_Diff '
Iters = -1
Results_Table_Row = Results_Table_Row + 1
Stock_Volume = 0
            
            
End If
            
Next R
           
End With

Next W

End Sub












