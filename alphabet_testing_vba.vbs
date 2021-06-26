Sub TickerTape()
'------------------------
'GOALS
'------------------------
    'a.  The ticker symbol.
    'b.  Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'c.  The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'd.  The total stock volume of the stock.
    'e.  You should also have conditional formatting that will highlight positive change in green and negative change in red.

'------------------------
'Loop Mechanism
'------------------------

For Each ws In Worksheets

'Dim all the things
Dim TickerName As String

Dim OpenValue As Double
OpenValue = 0

Dim CloseValue As Double
CloseValue = 0

Dim TotalVolume As Double
TotalVolume = 0

'Set where to start putting the summary data
Dim SummaryTableRow As Integer
SummaryTableRow = 2


'set the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Need to define the range of data
  ' Loop through all credit card purchases
  For i = 2 To LastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the TickerName and OpenValue
    TickerName = ws.Cells(i, 1).Value
    CloseValue = ws.Cells(i, 6).Value
      

      ' Add to the Brand Total
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
    ws.Range("j" & SummaryTableRow).Value = TickerName
      
      ' Calc the yearly change
    ws.Range("k" & SummaryTableRow).Value = CloseValue - OpenValue
      
      ' Calc the yearly change, checking for zero first values
      If OpenValue = 0 Then
    ws.Range("l" & SummaryTableRow).Value = "N/A"
    Else
    
    ws.Range("l" & SummaryTableRow).Value = (CloseValue - OpenValue) / OpenValue
    ws.Range("l" & SummaryTableRow).NumberFormat = "0.00%"
      
      End If
      
      ' Print the Total Volume to the Summary Table
      
    ws.Range("m" & SummaryTableRow).Value = TotalVolume
      

      ' Add one to the summary table row
    SummaryTableRow = SummaryTableRow + 1
      
      ' Reset the Total Volume
    TotalVolume = 0

    ' If the cell immediately following a row is the same brand...
   
   ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
     OpenValue = ws.Cells(i, 3).Value
      
    Else
    
      ' Add to the Brand Total
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value

      
    End If

  Next i
  
  Dim myrange As Range
  Set myrange = ws.Range("K2:l" & LastRow)

For Each Cell In myrange
    If Cell.Value < 0 Then
    Cell.Interior.Color = RGB(255, 0, 0)

    ElseIf Cell.Value > 0 Then
    Cell.Interior.Color = RGB(0, 255, 0)
    
 End If
    
Next

'set the last row of data set
ResultsLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Create the new ranges based on rows
Dim mygreatrange As Range
Set mygreatrange = ws.Range("L2:L" & ResultsLastRow)

Dim myvolumerange As Range
Set myvolumerange = ws.Range("M2:M" & ResultsLastRow)

'Dim the ref values, find the value of the function in question
Dim MaxValue As Double
MaxValue = WorksheetFunction.Max(mygreatrange)

Dim MinValue As Double
MinValue = WorksheetFunction.Min(mygreatrange)

Dim GreatestVolume As Double
GreatestVolume = WorksheetFunction.Max(myvolumerange)

'Set each value in it's appropriate cell and format as necessary
ws.Range("R2").Value = MaxValue
ws.Range("R2").NumberFormat = "0.00%"
ws.Range("R3").Value = MinValue
ws.Range("R3").NumberFormat = "0.00%"
ws.Range("R4").Value = GreatestVolume

'Find the corresponding ticker by matching the index value from the function results
Dim MaxValIndex As Double

MaxValIndex = WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("L2:L" & ResultsLastRow), 0)
ws.Range("Q2").Value = ws.Cells(MaxValIndex + 1, 10)

Dim MinValIndex As Double
MaxValIndex = WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("L2:L" & ResultsLastRow), 0)
ws.Range("Q3").Value = ws.Cells(MaxValIndex + 1, 10)

Dim VolumeIndex As Double
VolumeIndex = WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("M2:M" & ResultsLastRow), 0)
ws.Range("Q4").Value = ws.Cells(VolumeIndex + 1, 10)
 Next ws

End Sub

Sub Greatest()
'set the last row

ResultsLastRow = Cells(Rows.Count, 11).End(xlUp).Row

Dim mygreatrange As Range
Set mygreatrange = Range("L2:L" & ResultsLastRow)

Dim myvolumerange As Range
Set myvolumerange = Range("M2:M" & ResultsLastRow)

Dim MaxValue As Double
MaxValue = WorksheetFunction.Max(mygreatrange)

Dim MinValue As Double
MinValue = WorksheetFunction.Min(mygreatrange)

Dim GreatestVolume As Double
GreatestVolume = WorksheetFunction.Max(myvolumerange)


Range("R2").Value = MaxValue
Range("R2").NumberFormat = "0.00%"
Range("R3").Value = MinValue
Range("R3").NumberFormat = "0.00%"
Range("R4").Value = GreatestVolume

Dim MaxValIndex As Double

MaxValIndex = WorksheetFunction.Match(Range("R2").Value, Range("L2:L" & ResultsLastRow), 0)
Range("Q2").Value = Cells(MaxValIndex + 1, 10)

Dim MinValIndex As Double
MaxValIndex = WorksheetFunction.Match(Range("R3").Value, Range("L2:L" & ResultsLastRow), 0)
Range("Q3").Value = Cells(MaxValIndex + 1, 10)

Dim VolumeIndex As Double
VolumeIndex = WorksheetFunction.Match(Range("R4").Value, Range("M2:M" & ResultsLastRow), 0)
Range("Q4").Value = Cells(VolumeIndex + 1, 10)




End Sub
