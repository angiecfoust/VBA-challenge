Attribute VB_Name = "Module1"
Sub stockyear()

Dim ticker_name As String
Dim change As Double
Dim percent As Double
Dim volume As Double

volume = 0
change = 0

Dim summary_table_row As Integer
summary_table_row = 2
Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"



For i = 2 To 753001
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker_name = Cells(i, 1).Value
    volume = volume + Cells(i, 7).Value
    Cells(summary_table_row, 9).Value = ticker_name
    'Cells(summary_table_row, 10).Value = change 'still need to figure this out
    Cells(summary_table_row, 12).Value = volume
    
    summary_table_row = summary_table_row + 1
    
    volume = 0
    'change= 0
    
    Else
    volume = volume + Cells(i, 7).Value

End If

     

Next i





    'loop for one year and output the following columns
        'ticker symbol
        'yearly change opening price to closing price
        'cond format: negative red cells, positive green cells
        'percent change from opening to closing
        'total stock volume of the stock
    
    'new chart- rows will be below (report ticker and value for each)
        'greatest % increase
        'greatest % decrease
        'greatest total volume
    

End Sub
