Attribute VB_Name = "Module2"
Sub Module2_Script()

    'create loop through all sheets
    For Each ws In Worksheets
    
    'Step 1 - Work with Columns I through L
    'create new headers on each sheet. Use 'ws' before cell
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'verified the above works
    
    'set variables for future use
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Total_vol As Variant

    
    Dim Ticker As String
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
        
    'find the last row of the data
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'initially set volume counter 'total_vol' to zero
    Total_vol = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
       
    'define open price as 3 column 'C', starting in second row
    Open_Price = ws.Cells(2, 3).Value
    
        'start summary starting in second row
        For i = 2 To lastrow
        
            'add the volumes in row G
            Total_vol = Total_vol + ws.Cells(i, 7).Value
            
            'if next cell is not equal, output continue iteration
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                
                'define new values in rows I - L
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_vol
                                
                Close_Price = ws.Cells(i, 6).Value
                                      
                Quarterly_Change = Close_Price - Open_Price
                    'if open price is zero, will cause error. Change percent_change to zero in that case
                    If Open_Price <> 0 Then
                    Percent_Change = Quarterly_Change / Open_Price
                    Else: Percent_Change = 0
                    End If
                    
            Open_Price = ws.Cells(i + 1, 3).Value
            ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            'use conditional formatting to highlight positive changes in green and negative changes in red
                If Quarterly_Change < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
            
            'keep adding new rows for each ticker and reset the ticker total for each line
            Summary_Table_Row = Summary_Table_Row + 1
            Total_vol = 0
                       
            
            End If
            
            'continue the iterations
            Next i
            
            'verify that values have correct formatting
            ws.Range("K2").EntireColumn.NumberFormat = "0.00%"
            ws.Range("L2").EntireColumn.NumberFormat = "#,##0"
            

    'Step 2 - Work with Columns O through Q
    
        'create new headers on each sheet. Use 'ws' before cell
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'caclulate the greatest increase,decrease, and total volume
        'define new variables
        Dim Greatest_Increase As Variant
        Dim Greatest_Decrease As Variant
        Dim Greatest_Total_Vol As Variant
        Dim Greatest_Increase_Ticker As Variant
        Dim Greatest_Decrease_Ticker As Variant
        Dim Greatest_Total_Ticker As Variant
    
        'set as range data to search
        Dim Ticker_Range As Range
        Set Ticker_Range = ws.Range("I2:I1000000")
        Dim Percent_Range As Range
        Set Percent_Range = ws.Range("K2:K1000000")
        Dim Total_Range As Range
        Set Total_Range = ws.Range("L2:L1000000")
    
        Greatest_Increase = Application.Max(Percent_Range)
        Greatest_Decrease = Application.Min(Percent_Range)
        Greatest_Total_Vol = Application.Max(Total_Range)
        Greatest_Increase_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Increase, Percent_Range, 0))
        Greatest_Decrease_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Decrease, Percent_Range, 0))
        Greatest_Total_Ticker = WorksheetFunction.Index(Ticker_Range, WorksheetFunction.Match(Greatest_Total_Vol, Total_Range, 0))
                
        'Add new values to colums P & Q
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(4, 17).Value = Greatest_Total_Vol
        ws.Cells(2, 16).Value = Greatest_Increase_Ticker
        ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
        ws.Cells(4, 16).Value = Greatest_Total_Ticker
        
        'verify that values have correct formatting
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,##0"
        
        'autofit all columns of each worksheet and center contents of row 1
        ws.Columns.AutoFit
        ws.Rows(1).HorizontalAlignment = xlCenter
        
    Next ws
    
End Sub
