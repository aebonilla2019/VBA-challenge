Sub stock_macro()
'define variable to keep track # of stock ticker
Dim row_num As Single

'define variable to store the row nunmber of the last row in the stock ticker list
Dim last_row As Single

'define variable to keep track of the unique # of stock ticker
Dim result_row_num As Single

'define variable to store the row nunmber of the last row in the unique stock ticker list
Dim results_last_row As Single
    
'define variables to store begining and ending stock prices
Dim bgn_stock_pr As Single
Dim end_stock_pr As Single

'Define variables to store where to begin and end adding volumes
Dim bgn_volum_row As Single
Dim end_volum_row As Single

'Define variable for total volume
Dim total_volum As Single

 'define variable for worksheet count and number
Dim ws_count As Integer
Dim ws_num As Integer
    
'define result array to store results
Dim results_array(1 To 3, 1 To 2) As Variant
    



   'total number of worksheets
   ws_count = ActiveWorkbook.Worksheets.Count
    
    'cycle through each worksheet
    For ws_num = 1 To ws_count
        
    'begin recording results from row 2
    resultRow = 2
           
     With ActiveWorkbook.Worksheets(ws_num)
        
            'write column and row headers for results table
          .Range("I1").Value = "Ticker"
          .Range("J1").Value = "Yearly Change"
          .Range("K1").Value = "Percent Change"
          .Range("L1").Value = "Total Stock Volume"
        
          .Range("P1").Value = "Ticker"
          .Range("Q1").Value = "Value"
        
          .Range("O2").Value = "Greatest % Increase"
          .Range("O3").Value = "Greatest % Decrease"
          .Range("O4").Value = "Greatest Total Volume"
        
           ' Record row number of final row with a ticker symbol
           last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
            
               'Loop through each row with a stock ticker in worksheet
                For row_num = 1 To last_row
            
                'for each row as if the next row and the current row match; if they do not
            
                 If .Cells(row_num, "A").Value <> .Cells(row_num + 1, "A").Value Then
                    
                    'if row doesn't = 1 then ...
                    
                    If row_num <> 1 Then
                        
                        'calculate and write Yearly Change
                        end_stock_pr = .Cells(row_num, "F").Value
                        .Cells(resultRow - 1, "J").Value = Format(end_stock_pr - bgn_stock_pr, "#,###.00")
                        
                        'calculate and write Total Stock Volume
                        end_volum_row = row_num
                        .Cells(resultRow - 1, "L").Value _
                        = .Application.Sum(Range(Cells(bgn_volum_row, "G"), Cells(end_volum_row, "G")))
                        

                 'if beginning stock price is not 0 then...
                        
                        If bgn_stock_pr <> 0 Then
                            
                            'calculate and write percent change of stock value
                          .Cells(resultRow - 1, "K").Value = Format(end_stock_pr / bgn_stock_pr - 1, "Percent")
                            
                            'fotmat cell green if change is positive
                            If .Cells(resultRow - 1, "J").Value > 0 Then
                            
                              .Cells(resultRow - 1, "J").Interior.ColorIndex = 4
                                
                            'format cell red if change is negative
                            Else
                              .Cells(resultRow - 1, "J").Interior.ColorIndex = 3
                            End If
                        
                        End If
                   
                    End If    
                 
            'if next row is blank then...
                    
                    If VarType(.Cells(row_num + 1, "A").Value) <> 0 Then
                        
                        'write ticker value in results column
                        .Cells(resultRow, "I").Value = .Cells(row_num + 1, "A").Value
                        
                        'record beginning stock value for next stock
                        bgn_stock_pr = .Cells(row_num + 1, "C").Value
                        
                        'record bgn row number to be used for adding stock volume for next stock
                        bgn_volum_row = row_num + 1
                        
                        'add one row for next set of results
                        resultRow = resultRow + 1
                    
                    End If
                    
                 End If
            
            Next row_num
            
            results_last_row = .Cells(.Rows.Count, "I").End(xlUp).Row
            
            'record initial values in array to compare to
            results_array(1, 2) = 0
            results_array(2, 2) = 0
            results_array(3, 2) = 0
            
           'loop through results col
            
            For resultRow = 2 To results_last_row
            
           'if values are greater/less than value saved in array,
           'write over them with current value and record corresponding ticker symbol

                'if the value in the result row is greater than value in the result array then...
                If .Cells(resultRow, "K").Value > results_array(1, 2) Then
                    'write over the current amount in the result array
                    results_array(1, 2) = .Cells(resultRow, "K").Value
                    results_array(1, 1) = .Cells(resultRow, "I").Value
                End If
                
                'if the value in the result row is less than value in the result array then...
                If .Cells(resultRow, "K").Value < results_array(2, 2) Then
                    'write over the current amount in the result array
                    results_array(2, 2) = .Cells(resultRow, "K").Value
                    results_array(2, 1) = .Cells(resultRow, "I").Value
                End If
                
                'if the value in the result row is greater than value in the result array then...
                If .Cells(resultRow, "L").Value > results_array(3, 2) Then
                    'write over the current amount in the result array
                    results_array(3, 2) = .Cells(resultRow, "L").Value
                    results_array(3, 1) = .Cells(resultRow, "I").Value
                End If
            
            Next resultRow
        
            ' write values stored in array to final cells
            .Cells(2, "P").Value = results_array(1, 1)
            .Cells(2, "Q").Value = Format(results_array(1, 2), "Percent")
            .Cells(3, "P").Value = results_array(2, 1)
            .Cells(3, "Q").Value = Format(results_array(2, 2), "Percent")
            .Cells(4, "P").Value = results_array(3, 1)
            .Cells(4, "Q").Value = results_array(3, 2)
        
    End With

    Next ws_num
    

End Sub