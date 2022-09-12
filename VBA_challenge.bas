Attribute VB_Name = "Module1"
Sub VBA_challenge()

    Dim Ticker As String
    Dim Trading_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim Total_Rows As Long
    Dim Ticker_First_Date_Row As Double
    Dim Ticker_First_Date_Open As Double
    Dim Ticker_Last_Date_Close As Double
    Dim Yearly_Change As Double
    Dim Yearly_Percent_Change As Double
        
    'Setting variables that will be used for counters and running totals & finding total rows in sheet
    Trading_Volume = 0
    Summary_Table_Row = 2
    Ticker_First_Date_Row = 2
    Total_Rows = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To Total_Rows
        'If next ticker <> current ticker ready to print final totals to summary table for current ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Set current ticker symbol
            Ticker = Cells(i, 1).Value
            
            'Add final cell to running total
            Trading_Volume = Trading_Volume + Cells(i, 7).Value
            
            'Printing ticker to summary table
            Range("I" & Summary_Table_Row).Value = Ticker
            
            'Printing trading volume to summary table
            Range("L" & Summary_Table_Row).Value = Trading_Volume
            
            'Grab first date open value
            Ticker_First_Date_Open = Cells(Ticker_First_Date_Row, 3)
            
            'Grab last date close value
            Ticker_Last_Date_Close = Cells(i, 6)
            
            'Calculating & priting yearly change
            Range("J" & Summary_Table_Row).Value = Ticker_Last_Date_Close - Ticker_First_Date_Open
            
            'Calculating & priting percent change
            Range("K" & Summary_Table_Row).Value = (Ticker_Last_Date_Close - Ticker_First_Date_Open) / Ticker_First_Date_Open
            
            'Variable resets are performed below to prepare for next ticker
            'Add next row to summary table row counter
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset trading volume running total for next ticker symbol
            Trading_Volume = 0
            
            'Reset ticker first date to start counting for next ticker symbol
            Ticker_First_Date_Row = i + 1
            

        'Else still counting & adding to toal for current ticker
        Else
            Trading_Volume = Trading_Volume + Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Adding headers to summary table
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    'Adding Conditional formatting to Yearly Change column
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'General column format
    Columns("J:J").Select
    Selection.NumberFormat = "$#,##0.00"
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Columns("L:L").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    

End Sub


