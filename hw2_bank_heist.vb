Sub BankHeist():
    Dim Ticker_Symbol As String
    Dim percent_change As Long

    Dim Total_volume As Double
    Total_volume = 0

    Dim Ticker_Symbol_Row As Integer
    Ticker_Symbol_Row = 2
    
    Lastrow = Worksheets("A").Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To Lastrow


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Symbol = Cells(i + 1, 1).Value
opening_value = Cells(3, 3).Value
closing_value = Cells(i + 1, 7).Value
percent_change = closing_value - opening_value / opening_value

Total_stock_Volume = Total_stock_Volume + Cells(i, 7).Value

Range("I" & Ticker_Symbol_Row).Value = Ticker_Symbol
Range("J" & Ticker_Symbol_Row).Value = Total_stock_Volume
Ticker_Symbol_Row = Ticker_Symbol_Row + 1
Range("K" & Ticker_Symbol_Row).Value = percent_change
Total_stock_Volume = 0

opening_value = Cells(i + 1, 3).Value

percent_change = closing_value - opening_value / opening_value
Range("k" & ticker_symbole_row).Value = percent_change

Else

Total_stock_Volume = Total_stock_Volume + Cells(i, 7).Value
End If





        


    
    End Sub
    

