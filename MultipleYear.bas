Attribute VB_Name = "Module1"
Sub VBAHW_MultipleYear_PS()


'Loop through all the sheets and do stuff in each of the sheets
'Use a name defined by me, ws, to get to each worksheet using the Worksheets keyword
For Each ws In Worksheets

'Set an initial variable for holding a ticker name
Dim Ticker As String

'Set an initial variable for holding the total volume per ticker
Dim TickerTotalVolume As Double
TickerTotalVolume = 0

'Keep track of the location for each ticker in the summary table, start at row 2
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Evaluate each sheet from the start of data to the last row of that sheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Place headers for new columns we are placing data, and have column wide enough to fit the Total Stock Volume header
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"
Range("J1").Columns.AutoFit


'Start looping within each sheet
For i = 2 To LastRow

'Check to see if we are considering the same ticker, start work in case if not.
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Get the ticker name and assign volume to that ticker
Ticker = ws.Cells(i, 1).Value
TickerTotalVolume = TickerTotalVolume + ws.Cells(i, 7).Value

'Place ticker name in a summary table
ws.Range("I" & SummaryTableRow).Value = Ticker

'Place volume for each ticker in the summary table
ws.Range("J" & SummaryTableRow).Value = TickerTotalVolume

'Go to the next row
SummaryTableRow = SummaryTableRow + 1

'Reset the ticker volume
TickerTotalVolume = 0


'Now evaluate in case the next row contains the same ticker
Else
TickerTotalVolume = TickerTotalVolume + ws.Cells(i, 7).Value


'Stop evaluating...
End If

'And move to the next row
Next i

'Go to the next spreadsheet
Next ws


End Sub
