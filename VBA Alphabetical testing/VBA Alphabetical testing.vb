Sub StockAnalysis()


    'Set worksheet name
    Dim ws As Worksheet
    Set ws = Worksheets("A")
    
    '-----------------------
    'LOOP THROUGH ALL SHEETS
    '-----------------------
    For Each ws in Worksheets

        '-----------------------
        'CREATE SUMMARY CHART LABELS
        '-----------------------
        
        'Create Summary Chart Labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Create Challenge Chart Labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        '-----------------------
        'APPLY SUMMARY CHART FORMATTING
        '-----------------------

        'Format Summary Chart column as percent
        ws.Range("K:K").Style = "Percent"
        ws.Range("Q2:Q3").Style = "Percent"


        '-----------------------
        'SET VARIABLES TO HOLD TICKER, OPEN/CLOSE DATES, AND TOTAL VOLUME NAMES
        '-----------------------

        'Set initial variable to hold the ticker name
        Dim Ticker_Name As String

        'Set initial variable to hold the total stock volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0

        'Set initial variable to hold the ticker year open price
        Dim Year_Open_Price As Double
        Year_Open_Price = ws.Cells(2, 3).Value

        'Set initial variable to hold the ticker year close price
        Dim Year_Close_Price As Double
        Year_Close_Price = ws.Cells(2, 6).Value

        'Set first day of year
        Dim Year_Open_Date As Double
        Year_Open_Date = ws.Cells(2, 2).Value

        'Set last day of year
        Dim Year_Close_Date As Double
        Year_Close_Date = ws.Cells(2, 2).Value

        'Set initial variable to hold the yearly change
        Dim Yearly_Change As Double
        Yearly_Change = 0

        'Set initial variable to hold the percent change
        Dim Percent_Change As Double
        Percent_Change = 0


        'Keep track of the location for each ticker in the summary table
        Summary_Table_Row = 2

        'Keep track of last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



        '-----------------------
        'LOOP THROUGH ALL TICKERS IN SHEET
        '-----------------------
        For i = 2 To LastRow


                '-----------------------
                'INSERT THE TICKER AND TOTAL VOLUME INTO THE SUMMARY TABLE
                '-----------------------

                ' Check if we are still within the same ticker name, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    'Set the Ticker Name
                    Ticker_Name = ws.Cells(i, 1).Value

                    'Add to the Total Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                    'Add to the Year Close Price
                    Year_Close_Price = ws.Cells(i, 6).Value

                    'Calculate the Yearly Change
                    Yearly_Change = Year_Close_Price - Year_Open_Price

                    'Calculate the Percent Change 
                        'Avoid dividing by 0.  If Year Open price is 0, divide by 1 instead of Year Open Price
                    If Year_Open_Price = 0 Then
                        Percent_Change = Yearly_Change / 1
                        'Otherwise, divide by Year Open Price 
                    Else
                        Percent_Change = Yearly_Change / Year_Open_Price
                    End If
                                        
                    ' Print the Ticker Name in the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                    ' Print the Yearly Change in the Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                    ' Print the Percent Change in the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                            'Format Percentage Column
                            If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                                'Format as Green
                                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 50
                            ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
                                'Format as Red
                                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                            End If

                    ' Print the Total Stock Volume to the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    ' Reset the Total Stock Volume
                    Total_Stock_Volume = 0

                    'Reset the Year Open Date
                    Year_Open_Date = ws.Cells(i + 1, 2).Value
                    Year_Open_Price = ws.Cells(i + 1, 3).Value
                    Year_Close_Date = ws.Cells(i + 1, 2).Value
                    Year_Close_Price = ws.Cells(i + 1, 6).Value
                    ' If the cell immediately following a row has the same ticker stock...
                Else

                    ' Add to the Total Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    ' Set the opening value for the first day of the year
                    
                        'Reset Year Open date to beginning of year if stock has changed
                        If ws.Cells(i, 2).Value < Year_Open_Date Then
                            Year_Open_Date = ws.Cells(i, 2).Value
                            Year_Open_Price = ws.Cells(i, 3).Value
                        End If

                        'Set the Year Close date and the Year Close Price to value on last day of year
                        If Cells(i, 2).Value > Year_Close_Date Then
                            Year_Close_Date = ws.Cells(i, 2).Value
                            Year_Close_Price = ws.Cells(i, 6).Value
                        End If

                End If
                
        Next i

        '------------------------------
        ' FIND MAX STOCK VOLUME
        '------------------------------
        'Set initial variable to hold the ticker name
        Dim Volume_Ticker As String
        Dim Volume_Row As Long

        'Set initial variable to hold range for total stock volume
        Dim Stock_Volume_Range As Range
        Set Stock_Volume_Range = ws.Range("L:L")

        'Set initial variable to hold the total stock volume. Set volume and print to summary table
        Total_Stock_Volume = Application.WorksheetFunction.Max(Stock_Volume_Range)
        ws.Range("Q4").Value = Total_Stock_Volume

        'Find row where total stock volume is located. Print to Summary Table
        Volume_Row = Application.WorksheetFunction.Match(Total_Stock_Volume, Stock_Volume_Range, 0)
        Volume_Ticker = ws.Range("I" & Volume_Row).Value
        ws.Range("P4").Value = Volume_Ticker


        '------------------------------
        ' FIND MIN PERCENT DECREASE 
        '------------------------------
        Dim Min_Percent_Ticker As String
        Dim Min_Percent_Row As Long

        'Set initial variable to hold range for min percent decrease
        Dim Min_Decrease_Range As Range
        Set Min_Decrease_Range = ws.Range("K:K")

        'Set initial variable to hold min % decrease. Set % decrease and print to summary table
        Dim Min_Percent_Decrease As Double
        Min_Percent_Decrease = Application.WorksheetFunction.Min(Min_Decrease_Range)
        ws.Range("Q3").Value = Min_Percent_Decrease

        'Find row where min % decrease is located. Print to Summary Table
        Min_Percent_Row = Application.WorksheetFunction.Match(Min_Percent_Decrease, Min_Decrease_Range, 0)
        Min_Percent_Ticker = ws.Range("I" & Min_Percent_Row).Value
        ws.Range("P3").Value = Min_Percent_Ticker


        '------------------------------
        ' FIND MAX PERCENT DECREASE 
        '------------------------------
        Dim Max_Percent_Ticker As String
        Dim Max_Percent_Row As Long

        'Set initial variable to hold range for Max percent decrease
        Dim Max_Decrease_Range As Range
        Set Max_Decrease_Range = ws.Range("K:K")

        'Set initial variable to hold Max % decrease. Set % decrease and print to summary table
        Dim Max_Percent_Decrease As Double
        Max_Percent_Decrease = Application.WorksheetFunction.Max(Max_Decrease_Range)
        ws.Range("Q2").Value = Max_Percent_Decrease

        'Find row where Max % decrease is located. Print to Summary Table
        Max_Percent_Row = Application.WorksheetFunction.Match(Max_Percent_Decrease, Max_Decrease_Range, 0)
        Max_Percent_Ticker = ws.Range("I" & Max_Percent_Row).Value
        ws.Range("P2").Value = Max_Percent_Ticker

    '------------------------------
    ' SHEET FIXES COMPLETE
    '------------------------------
    Next ws





End Sub