Attribute VB_Name = "Module1"
Sub TestData()
'Loop through data
    'Find last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'Declare variables
    Summarytablerow = 2
    Opening_price = Cells(2, 3).Value
    
    'Add headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    
    For i = 2 To lastrow
        'If the current cell doesnt match the next cell then determine when ticker changes
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'List ticker in column I
            Cells(Summarytablerow, 9).Value = Cells(i, 1).Value
            'Declare closing price variable
            Closing_price = Cells(i, 6).Value
            'subtract yc from cp
            Yearly_change = Closing_price - Opening_price
            'Declare yearly change variable
            Cells(Summarytablerow, 10).Value = Yearly_change
            
            Percent_change = Yearly_change / Opening_price
            Cells(Summarytablerow, 11).Value = Percent_change
            
            'Calculate Total Stock Volume
            Total_Stock_Volume = Cells(i, 7).Value
            Cells(Summarytablerow, 12).Value = Total_Stock_Value
            
            
            'Update Summarytablerow
            Summarytablerow = Summarytablerow + 1
            'Update opening price
            Opening_price = Cells(i + 1, 3).Value
            
        End If
    
    Next i
End Sub
