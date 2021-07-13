Attribute VB_Name = "Module1"
Sub Reset()

For Each ws In ThisWorkbook.Worksheets

ws.Activate

Columns("I:P").Delete

Cells(1, 1).Select

Next ws

Worksheets(1).Select

End Sub

Sub TickerSummary()

'Setting up the summary table headers

'Label headers in summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Sorting the raw data in order by ticker then by date
'Sort column C, after column N and after column P

Range("A:G").Select
Selection.Columns.Sort key1:=Columns("a"), Order1:=xlAscending, Key2:=Columns("b"), Header:=xlYes

'Defining variables for the row loop that will feed the summary table

Dim NumRows As Long
Dim CurTicker As String
Dim StockVolume As Variant
Dim TickerRowStart As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

'Assigning the starting variables

NumRows = Cells(Rows.Count, 1).End(xlUp).Row 'for the To loop
StockVolume = 0 'starting at zero to sum up the stock volume per ticker
TickerRowStart = 2 'starting point
OpenPrice = Cells(2, 3).Value 'starting point

'Looping through all the rows to get the needed summarization for the summary table row for each Ticker

For i = 2 To NumRows

    'Looking ahead to see if Ticker changes

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'if it does
    
      'Before it changes, capture the summary of current ticker
      
      CurTicker = Cells(i, 1).Value 'This is the Current Ticker to print to summary table
      Range("I" & TickerRowStart).Value = CurTicker 'This prints the Current Ticker to the summary table
      StockVolume = StockVolume + Cells(i, 7).Value 'This is the counter that is summing each row of stock volume (in both If cases) - it resets after each ticker
      Range("L" & TickerRowStart).Value = StockVolume 'This prints the counter as it is since it's the last row before the Ticker changes
      ClosePrice = Cells(i, 6).Value 'This saves the close price column of the Current Ticker row to use in Yearly Change formula
      YearlyChange = (ClosePrice - OpenPrice) 'This calculates the Yearly Change
      Range("J" & TickerRowStart).Value = YearlyChange 'This prints the Yearly Change to the summary table
            
            'Need this if statement to avoid div/0 error for Percent formula when OpenPrice is 0
            If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
            ElseIf (OpenPrice = 0 And ClosePrice > 0) Then
                    PercentChange = 1
            Else
                    PercentChange = YearlyChange / OpenPrice
            End If

     Range("K" & TickerRowStart).Value = PercentChange 'This prints the percent change to the summary table then formats it
     Range("K" & TickerRowStart).NumberFormat = "0.00%"
     
     TickerRowStart = TickerRowStart + 1 'This keeps track of which row we're on
     
     StockVolume = 0 'This resets the Stock Volume since we got what we needed for the summary table already and we want it to start over with next ticker
     
     OpenPrice = Cells(i + 1, 3) 'Sets the Open Price for the next ticker since we know the next row is a new ticker and we need the OpenPrice from that row
     
    Else
        
            StockVolume = StockVolume + Cells(i, 7).Value 'This is the counter that is summing each row of stock volume (in both If cases as we pass through each row) - it resets after each ticker
            
    End If
        
    
Next i

'--------------------------------------------------------------------------
'All the formatting next
'--------------------------------------------------------------------------

  
'Set summary tables to bold
Range("I1:L1").EntireColumn.Font.Bold = True

'Autofit summary table columns
Range("A1:L1").EntireColumn.AutoFit

'Conditional formatting that will highlight positive change in green and negative change in red
    
    'Color code yearly change
    
    SummaryLastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To SummaryLastRow
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4 'Green
            Else
                Cells(i, 10).Interior.ColorIndex = 3 'Red
            End If
            
    Next i

'Setting last position
Cells(1, 1).Select
MsgBox ("Summary table complete.")


End Sub



