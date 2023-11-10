Attribute VB_Name = "Module1"
Sub stockyear()

For Each ws In Worksheets

Dim ticker_name As String
Dim change As Double
Dim percent As Double
Dim volume As Double
Dim open_price As Double
Dim close_price As Double
Dim open_index As Long


volume = 0
change = 0
open_index = 2


Dim summary_table_row As Integer
summary_table_row = 2

ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"




    For i = 2 To 753001
        
        ticker_name = ws.Cells(i, 1).Value 'this is working- DO NOT CHANGE
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
        open_price = ws.Cells(open_index, 3).Value
        close_price = ws.Cells(i, 6).Value 'this is working- DO NOT CHANGE
        volume = volume + ws.Cells(i, 7).Value 'this is working- DO NOT CHANGE
        
        
        ws.Cells(summary_table_row, 9).Value = ticker_name
        ws.Cells(summary_table_row, 12).Value = volume
        ws.Cells(summary_table_row, 10).Value = (close_price - open_price) 'this works- once the open and close prices are fixed
        ws.Cells(summary_table_row, 11).Value = FormatPercent(((close_price - open_price) / open_price), 2, vbFalse, vbFalse, vbFalse)
            If ws.Cells(summary_table_row, 11).Value > 0 Then
                ws.Cells(summary_table_row, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(summary_table_row, 11).Interior.ColorIndex = 3
            
            End If
                  
        
        summary_table_row = summary_table_row + 1
        open_index = (i + 1)

        
        ticker = ""
        volume = 0
        
         
        
        Else
        
        volume = volume + ws.Cells(i, 7).Value
        open_price = ws.Cells(i, 3).Value
             

    End If
   

Next i

    ws.Range("j1").EntireColumn.AutoFit
    ws.Range("k1").EntireColumn.AutoFit
    ws.Range("p1").EntireColumn.AutoFit
    ws.Range("l1").EntireColumn.AutoFit
    


Dim maxincrease As String
Dim maxdecrease As String
Dim maxvolume As String


    ws.Range("q2") = FormatPercent((WorksheetFunction.Max(ws.Range("k2:k753001"))), 2, vbFalse, vbFalse, vbFalse)
    ws.Range("q4") = WorksheetFunction.Max(ws.Range("l2:l753001"))
    ws.Range("q3") = FormatPercent((WorksheetFunction.Min(ws.Range("k2:k753001"))), 2, vbFalse, vbFalse, vbFalse)
    

    maxincrease = ws.Range("q2").Value
    maxdecrease = ws.Range("q3").Value
    maxvolume = ws.Range("q4").Value
    
    If maxincrease = ws.Cells(summary_table_row, 11).Value Then
        ws.Range("p2").Value = ws.Cells(summary_table_row, 11 - 2).Value
        
        End If
        
    If maxdecrease = ws.Cells(summary_table_row, 11).Value Then
        ws.Range("p3").Value = ws.Cells(summary_table_row, 11 - 2).Value
        
        End If
        
    If maxvolume = Cells(summary_table_row, 12).Value Then
        ws.Range("p4").Value = ws.Cells(summary_table_row, 11 - 2).Value
        
        End If
        
    
    
    ws.Range("o1").EntireColumn.AutoFit
    ws.Range("p1").EntireColumn.AutoFit
    ws.Range("q1").EntireColumn.AutoFit


Next ws


End Sub

    
