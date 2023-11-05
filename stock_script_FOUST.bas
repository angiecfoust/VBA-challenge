Attribute VB_Name = "Module1"
Sub stockyear()

Dim ticker_name As String
Dim change As Double
Dim percent As Double
Dim volume As Double
Dim open_price As Double
Dim close_price As Double
Dim counter As Double
Dim stock_date As String
Dim month_day As String



Range("c1").EntireColumn.Insert
Range("c1").Value = "MonthDate"



volume = 0
change = 0


Dim summary_table_row As Integer
summary_table_row = 2

Range("j1").Value = "Ticker"
Range("k1").Value = "Yearly Change"
Range("l1").Value = "Percent Change"
Range("m1").Value = "Total Stock Volume"
Range("n1").Value = "open" 'remove
Range("o1").Value = "close" 'remove


    For i = 2 To 753001
        
        stock_date = Cells(i, 2).Value
        month_day = Right(stock_date, 4)
        Cells(i, 3).Value = month_day
        
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_name = Cells(i, 1).Value 'this is working- DO NOT CHANGE
            If Cells(i, 3).Value = "102" Then
            open_price = Cells(i, 4)
            Else
            open_price = Cells(i + 1, 4).Value 'not working
            End If
            
        close_price = Cells(i, 7).Value 'this is working- DO NOT CHANGE
        volume = volume + Cells(i, 8).Value 'this is working- DO NOT CHANGE
        
        
        
        Cells(summary_table_row, 10).Value = ticker_name
        Cells(summary_table_row, 13).Value = volume
        Cells(summary_table_row, 14).Value = open_price 'remove
        Cells(summary_table_row, 15).Value = close_price 'remove
        'Cells(summary_table_row, 11).Value = (close_price - open_price) 'this works- once the open and close prices are fixed
        
        
        summary_table_row = summary_table_row + 1

        
        ticker = ""
        volume = 0
        'change = 0
        
         
        
        Else
        volume = volume + Cells(i, 8).Value
             

    End If


Next i


End Sub
