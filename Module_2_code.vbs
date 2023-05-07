Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data_Analysis()

For Each Sheet In Worksheets
    'Setting up variables
    Dim WorksheetName As String
        WorksheetName = Sheet.Name
    Dim x As Long
    Dim y As Long
    Dim Ticker_Count As Long
    Dim Last_Row_A As Long
    Dim Last_Row_I As Long
    Dim Percent_Change As Double
    Dim Greatest_Incr As Double
    Dim Greatest_Decr As Double
    Dim Greatest_Vol As Double
          
    'Column headers for 1st table
    Sheet.Cells(1, 9).Value = "Ticker"
    Sheet.Cells(1, 10).Value = "Yearly Change"
    Sheet.Cells(1, 11).Value = "Percent Change"
    Sheet.Cells(1, 12).Value = "Total Stock Volume"
    'Column headers for summary table
    Sheet.Cells(1, 16).Value = "Ticker"
    Sheet.Cells(1, 17).Value = "Value"
    Sheet.Cells(2, 15).Value = "Greatest % Increase"
    Sheet.Cells(3, 15).Value = "Greatest % Decrease"
    Sheet.Cells(4, 15).Value = "Greatest Total Volume"
        
    'Assigning values to variables for looping
    Ticker_Count = 2
    y = 2
    Last_Row_A = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
       
        
    'Looping to fill in 1st table
    For x = 2 To Last_Row_A
        'Checking if proceeding ticker name has changed
        If Sheet.Cells(x + 1, 1).Value <> Sheet.Cells(x, 1).Value Then
            'If true, write ticker name under ticker column
            Sheet.Cells(Ticker_Count, 9).Value = Sheet.Cells(x, 1).Value
            'Calculate and write Yearly Change in column
            Sheet.Cells(Ticker_Count, 10).Value = Sheet.Cells(x, 6).Value - Sheet.Cells(y, 3).Value
                'Formating Yearly Change
                If Sheet.Cells(Ticker_Count, 10).Value < 0 Then
                    'If true, change background color to red
                    Sheet.Cells(Ticker_Count, 10).Interior.ColorIndex = 3
                 Else
                    'If false, set background color to green
                    Sheet.Cells(Ticker_Count, 10).Interior.ColorIndex = 4
                End If

            'Calculate and write Percent_Change
            If Sheet.Cells(y, 3).Value <> 0 Then
                Percent_Change = ((Sheet.Cells(x, 6).Value - Sheet.Cells(y, 3).Value) / Sheet.Cells(y, 3).Value)
                    'Formating Percent_Change value as percentage
                    Sheet.Cells(Ticker_Count, 11).Value = Format(Percent_Change, "Percent")
            Else
                'If false, place a 0
                Sheet.Cells(Ticker_Count, 11).Value = Format(0, "Percent")
            End If
            'Calculate and write Total Volume
            Sheet.Cells(Ticker_Count, 12).Value = WorksheetFunction.Sum(Range(Sheet.Cells(y, 7), Sheet.Cells(x, 7)))
            'Adjust variables to make sure it works for next loop
                'Makes sure next ticker name is written underneath
                Ticker_Count = Ticker_Count + 1
                'Makes sure the correct row of values is used in calculations
                y = x + 1
                
        End If
    Next x
            
    'Assigning values to variables for 2nd looping
    Greatest_Vol = Sheet.Cells(2, 12).Value
    Greatest_Incr = Sheet.Cells(2, 11).Value
    Greatest_Decr = Sheet.Cells(2, 11).Value
    Last_Row_I = Sheet.Cells(Rows.Count, 9).End(xlUp).Row

    'Loop through 1st table to fill in summary table
    For x = 2 To Last_Row_I
        'Finding Greatest_Vol
        If Sheet.Cells(x, 12).Value > Greatest_Vol Then
            'If true, corresponding value becomes the new Greatest_Vol
            Greatest_Vol = Sheet.Cells(x, 12).Value
            'Write corresponding ticker name in summary table
            Sheet.Cells(4, 16).Value = Sheet.Cells(x, 9).Value
        Else
            'If false, Greatest_Vol stays the same
            Greatest_Vol = Greatest_Vol
        End If
                
        'Finding Greatest_Incr
        If Sheet.Cells(x, 11).Value > Greatest_Incr Then
            'If true, corresponding value becomes the new Greatest_Incr
            Greatest_Incr = Sheet.Cells(x, 11).Value
            'Write correspondng ticker name in summary table
            Sheet.Cells(2, 16).Value = Sheet.Cells(x, 9).Value
        Else
            'If false, Greatest_Incr value stays the same
            Greatest_Incr = Greatest_Incr
        End If

        'Finding Greatest_Decr
        If Sheet.Cells(x, 11).Value < Greatest_Decr Then
            'If true, corresponding value becomes new Greatest_Decr
            Greatest_Decr = Sheet.Cells(x, 11).Value
            'Write corresponding ticker name in summary table
            Sheet.Cells(3, 16).Value = Sheet.Cells(x, 9).Value
        Else
            'If false, Greatest_Decr value remains the same
            Greatest_Decr = Greatest_Decr
        End If
                
        'Filling in summary table + formating
        Sheet.Cells(2, 17).Value = Format(Greatest_Incr, "Percent")
        Sheet.Cells(3, 17).Value = Format(Greatest_Decr, "Percent")
        Sheet.Cells(4, 17).Value = Format(Greatest_Vol, "Scientific")
    Next x
            
            
Next Sheet
End Sub
