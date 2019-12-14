Sub runallworksheets()
'----------------------------------------------------------------------------------------------------------------------------
'                             Challenge 2
'----------------------------------------------------------------------------------------------------------------------------
'Declare in memory variable
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
'For loop to run subprocedure in all worksheets
    For Each xSh In Worksheets
        xSh.Select
'Function to start subprocedure
        Call summarydatafinal
    Next
    Application.ScreenUpdating = True
End Sub

Sub summarydatafinal()

'Declare in memory all the variables
Dim ws As Worksheet
Dim ticker As String
Dim lastRow As Long
Dim Summary_Table_Row As Integer
Dim openValueRow As Long
Dim openValue As Double
Dim closeValue As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim volumeTotal As Variant
Dim rng As Range
Dim maxTotalVolume As Variant
Dim maxPercentChange As Double
Dim minPercentChange As Double
Dim lastRowTotalVolume As Integer
Dim lastRowPercentChange As Integer
Dim maxTotalVolumeTicker As Range
Dim maxPercentChangeTicker As Range
Dim minPercentChangeTicker As Range

For Each ws In Worksheets

'Name the columns headers since this values will remain constant in all worksheets
    Range("I1,P1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

'Variable to find the last row of the ticker column
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Variable to collect data results below each other in their respective column
    Summary_Table_Row = 2

'Variable to set the count to 0 for each stock
    volumeTotal = 0

'Number 2 assinged to the initial value of "i" to skip the header
    i = 2

'Variable to store the value at which each stock will opens
    openValue = Round(Cells(i, 3).Value, 2)

'Set up of iteration from first value in row "2" to the last row
        For i = 2 To lastRow
   
'Set up to find when a ticker value is different from the next
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                
'Variable to store the value of each ticker and print it on column "I"
                ticker = Cells(i, 1).Value
                Range("I" & Summary_Table_Row).Value = ticker
        
'Variable to store the value when the stock closes
                closeValue = Round(Cells(i, 6).Value, 2)
                                
'Variable to calculate the change of the stock value over the year and print it on column "J"
                yearlyChange = closeValue - openValue
                Range("J" & Summary_Table_Row).Value = yearlyChange
            
'Conditional to change cell color to green for positive values
                If yearlyChange > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
 'Conditional to change cell color to red for negative values
                ElseIf yearlyChange < 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                        
'Conditional to avoid dividing by zero error
                If openValue <> 0 Then
'Variable to calculate the percent change of the stock value over the year and print in on column "K"
                percentChange = ((closeValue - openValue) / openValue) * 100
                Range("K" & Summary_Table_Row).Value = Round(percentChange, 2)
                End If
            
'Variable to store the value of the "next" stock at when it opens
                openValue = Round(Cells(i + 1, 3).Value, 2)
            
'Variable to store the last stock value as the stock ticker changes and print it on column "L"
                volumeTotal = volumeTotal + Cells(i, 7).Value
                Range("L" & Summary_Table_Row).Value = volumeTotal
                
'Variable to increase the row value by 1 to collect data below each other
                Summary_Table_Row = Summary_Table_Row + 1
'Reset of the total variable to zero as the count moves to another stock
                volumeTotal = 0
            
'Conditional to store the addition of the total volume as it goes through each iteration of the same stock
            Else
               volumeTotal = volumeTotal + Cells(i, 7).Value
            
            
            End If
    
        Next i

'----------------------------------------------------------------------------------------------------------------------------
'                             Challenge 1
'----------------------------------------------------------------------------------------------------------------------------


'Variable to find the last row of the "Total Stock Volume" list
    lastRowTotalVolume = Cells(Rows.Count, 12).End(xlUp).Row

'To find the maximum value in the Total Stock Volume list
        Set rng = Range(Cells(2, 12), Cells(lastRowTotalVolume, 12))
        maxTotalVolume = Application.WorksheetFunction.Max(rng)
'To find the ticker associated with the the maximum volume stock
        Set maxTotalVolumeTicker = Range(Cells(2, 12), Cells(lastRowTotalVolume, 12)).Find(maxTotalVolume, lookat:=xlWhole)
'To print both values (ticker and amount)
        Range("Q4").Value = maxTotalVolume
        Range("P4").Value = maxTotalVolumeTicker.Offset(, -3)
    
'Variable to find the last row of the "Percent Change" list
    lastRowPercentChange = Cells(Rows.Count, 11).End(xlUp).Row

'Set up to find the maximum value in the Percent Change list
        Set rng = Range(Cells(2, 11), Cells(lastRowPercentChange, 11))
        maxPercentChange = Application.WorksheetFunction.Max(rng)
'To find the ticker associated with the the maximum percent change
        Set maxPercentChangeTicker = Range(Cells(2, 11), Cells(lastRowTotalVolume, 11)).Find(maxPercentChange, lookat:=xlWhole)
'To print both values (ticker and amount)
        Range("Q2").Value = Format(maxPercentChange / 100, "0.00%")
        Range("P2").Value = maxPercentChangeTicker.Offset(, -2)

'Set up to find the minimun value in the Percent Change list
        Set rng = Range(Cells(2, 11), Cells(lastRowPercentChange, 11))
        minPercentChange = Application.WorksheetFunction.Min(rng)
'To find the ticker associated with the the minimun percent change
        Set minPercentChangeTicker = Range(Cells(2, 11), Cells(lastRowTotalVolume, 11)).Find(minPercentChange, lookat:=xlWhole)
'To print both values (ticker and amount)
        Range("Q3").Value = Format(minPercentChange / 100, "0.00%")
        Range("P3").Value = minPercentChangeTicker.Offset(, -2)
    
Next ws

End Sub