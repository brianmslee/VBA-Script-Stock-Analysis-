Sub Multiple_year_stock_data()

'Setting variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Summary_Table_Row As Integer
    Dim Total_Stock_Volume As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim First_Opening_Price As Double
    Dim Last_Closing_Price As Double
    Dim Max_Ticker As String
    Dim Min_Ticker As String
    Dim Max_Volume_Ticker As String
    Dim Max_Percent As Double
    Dim Min_Percent As Double
    Dim Max_Volume As Double

'For Each Loop to loop all of the sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets

'Creating Column Names
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
    
    'Formula to find last row of each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Giving value to variables
        Summary_Table_Row = 2
        Total_Stock_Volume = 0
        First_Opening_Price = Cells(2, 3).Value
        
    
    'Loop all of the data to the last row but every sheet has a different last row
        For i = 2 To LastRow
            
            
            
        'Loop to differentiate when the ticker names become different
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                
                Cells(Summary_Table_Row, 9).Value = Ticker
                
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                Last_Closing_Price = Cells(i, 6).Value
                
                Yearly_Change = Last_Closing_Price - First_Opening_Price
                
                Percent_Change = (Yearly_Change / First_Opening_Price) * 100
                
                Cells(Summary_Table_Row, 11).Value = Percent_Change & "%"
                
                Cells(Summary_Table_Row, 10).Value = Yearly_Change
                
                First_Opening_Price = Cells(i + 1, 3).Value
        
                Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                
                
                
                
                
                

        
        'Adding to the summary table row after each ticker name
                Summary_Table_Row = Summary_Table_Row + 1
        
        
        'Resetting the count
                Total_Stock_Volume = 0
                Yearly_Change = 0
                
        
        
        
        'If the ticker names are still the same
            Else
        
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        
        
            End If
    
    
        Next i
        
    Last_Summary_Row = ws.Cells(Rows.Count, 10).End(xlUp).Row
          
          
        For i = 2 To Last_Summary_Row
          
    
    
            If Cells(i, 10).Value > 0 Then
    
                Cells(i, 10).Interior.ColorIndex = 10
    
            ElseIf Cells(i, 10).Value <= 0 Then
        
                Cells(i, 10).Interior.ColorIndex = 3
                
            End If
                
      
        Next i
      
            
            
            
        
    
    
Next
    
            
        








End Sub

