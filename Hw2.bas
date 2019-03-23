Attribute VB_Name = "Module1"
Sub Stock()
'Create a string to loop through each ticker and calculate value

For Each ws In Worksheets
ws.Activate
    Set ActiveWorksheet = ActiveWorkbook.ActiveSheet
        
'Set an initial variable for holding the ticker letter
Dim Ticker As String

'Set an initial variable for holding the total per ticker
Dim Ticker_Total As Double
Ticker_Total = 0

'Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Volume As Double
Dim n As Double
n = WorksheetFunction.CountA(ActiveWorksheet.Columns(1))

n = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Volume"


'Loop through all stock value amounts
For i = 2 To n


'Check to see if stock value is the same, if not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Add to the Ticker Total
Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
Ticker = Cells(i, 1).Value


'Print the Ticker Letter in the Summary Table
Range("I" & Summary_Table_Row).Value = Ticker

'Print the Ticker Volume to the Summary Table
Range("J" & Summary_Table_Row).Value = Ticker_Total

'Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset the Ticker Total
Ticker_Total = 0

'If the cell immediately following a row is the same item...
Else:

'Add to the Ticker_Total
Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value



End If


Next i
Next ws

End Sub

