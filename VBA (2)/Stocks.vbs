Attribute VB_Name = "Module1"
Sub WallStreet()
    'TODO Hard challenge
    'Select all worksheets (ws)
    For Each ws In Worksheets
    
        'Set column headers for yearly results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set column headers for min & max
        'ws.Cells(1, 15).Value = "Ticker"
        'ws.Cells(1, 16).Value = "Value"
        
        'Set row headers for min & max
        'ws.Cells(2, 14).Value = "Greatest % Increase"
        'ws.Cells(3, 14).Value = "Greatest % Decrease"
        'ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Declare variables
        Dim Ticker As String
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyVolume As LongLong
        Dim YearlyChange As Double
        Dim YearlyPercentage As Double
        Dim ResultsTableIndex As Integer
        YearlyVolume = 0
        ResultsTableIndex = 2
        
        'Declare variables for min & max
        'Dim GreatestIncreaseTicker As String
        'Dim GreatestIncrease As Double
        'GreatestIncrease = 0
        'Dim GreatestDecreaseTicker As String
        'Dim GreatestDecrease As Double
        'GreatestDecrease = 0
        'Dim GreatestVolumeTicker As String
        'Dim GreatestVolume As Double
        'GreatestVolume = 0
        
        'Set last row index
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all tickers
        For RowIndex = 2 To LastRow
        
            'First instance of ticker
            If ws.Cells(RowIndex - 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
                'Pull yearly open
                YearlyOpen = ws.Cells(RowIndex, 3).Value
                
                'Pull ticker name
                Ticker = ws.Cells(RowIndex, 1).Value
            End If
                 
            'Add to yearly volume
            YearlyVolume = YearlyVolume + ws.Cells(RowIndex, 7).Value
            
            'Last instance of ticker
            If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
                'Pull yearly close
                YearlyClose = ws.Cells(RowIndex, 6).Value
                
                'Calculate yearly change
                YearlyChange = YearlyClose - YearlyOpen
                
                If YearlyOpen <> 0 Then
                    'Calculate percentage
                    YearlyPercentage = YearlyChange / YearlyOpen
                    'Set yearly percentage in cell
                    ws.Cells(ResultsTableIndex, 11).Value = YearlyPercentage
                    'Format yearly percentage
                    ws.Cells(ResultsTableIndex, 11).NumberFormat = "0.00%"
                    
                Else
                    ws.Cells(ResultsTableIndex, 11).Value = "n/a"
                End If
                
                
                'Add row to results table
                ws.Cells(ResultsTableIndex, 9).Value = Ticker
                ws.Cells(ResultsTableIndex, 10).Value = YearlyChange
                ws.Cells(ResultsTableIndex, 12).Value = YearlyVolume
                
                'Set color condition
                If YearlyChange > 0 Then
                    ws.Cells(ResultsTableIndex, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Cells(ResultsTableIndex, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Increment results table index
                ResultsTableIndex = ResultsTableIndex + 1
                            
                'Reset volume
                YearlyVolume = 0
                
            End If
        Next RowIndex
    Next ws
End Sub

