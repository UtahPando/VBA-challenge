VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Quarter_Report()

Dim ws As Worksheet

'Loop through each worksheet in Workbook
For Each ws In ThisWorkbook.Sheets

'Declare the data variables
Dim Ticker_symbol As String
Dim Opening_pr As Double
Dim Price_ch As Double
Dim Percent_ch As Double
Dim LastRow As Long
Dim wsCount As Integer
Dim Volume_Tot As Double
Dim Summary_Table_Row As Integer

'Set initial values for processing variables
    Volume_Tot = 0
    Summary_Table_Row = 2

'Set the Very 1st Opening Price
 Opening_pr = ws.Cells(2, 3).Value


'Count the number of rows in the worksheet(?)
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Loop through rows Within each workbook, while <ticker> is same
' save out the new value of <close>

'Create Header rows before proceeding into the table
      ws.Range("I" & 1).Value = "Ticker"
      ws.Range("J" & 1).Value = "Quarterly Change"
      ws.Range("K" & 1).Value = "Percent Change"
      ws.Range("L" & 1).Value = "Total Stock Volume"

'Loop through every row in the current quarter's (worksheet's) data
For i = 2 To LastRow

    ' Check if we are still within the same ticker symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Add the final value to the Volume_Tot Total
      Volume_Tot = Volume_Tot + ws.Cells(i, 7).Value
      
      'Calculates the changes before resetting Opening_pr to 0
      Price_ch = ws.Cells(i, 6) - Opening_pr
      
      'Writes out the quarterly change for the stock symbol
      ws.Range("J" & Summary_Table_Row).Value = Price_ch
      
    'Format Cells conditionally
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' Green fill for positive values
        ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) 'Red fill for negative values
        End If
      
      'Write out the Summary values for the percent change and Total Volume
      ws.Range("K" & Summary_Table_Row).Value = Price_ch / Opening_pr
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      ws.Range("L" & Summary_Table_Row).Value = Volume_Tot
      

      ' Resets the Ticker Symbol, Openinging Price, and Volume Total for the new stock
      Ticker_symbol = ws.Cells(i, 1).Value
      Opening_pr = ws.Cells(i + 1, 3).Value
      Volume_Tot = 0

      ' Print the Ticker_Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_symbol

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    Else 'Handle the case when the stock symbol is still the same, updating totals
      ' Add to the Volume_Tot Total
      Volume_Tot = Volume_Tot + ws.Cells(i, 7).Value
      Price_ch = ws.Cells(i, 6) - Opening_pr

    End If
     
  Next i

'Create the variables needed to Loop through the resulting summary table to find the Max % Increase, % Decline, and Volume

    Dim rng As Range
    Dim cell As Range
    Dim SummaryRows As Integer
    Dim MaxPctChg As Double
    Dim MinPctChg As Double
    Dim MaxVolume As Double
    
    MaxPctChg = -100
    MinPctChg = 100
    MaxVolume = 0
    
    SummaryRows = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

    ' Define the range to apply conditional formatting
    Set rng = ws.Range("K2:K" & SummaryRows) ' Update the range as needed
    
'LABEL THE SUMMARY TABLE
      ws.Range("P" & 1).Value = "Ticker"
      ws.Range("Q" & 1).Value = "Value"
      
      ws.Range("O" & 2).Value = "Greatest % Increase"
      ws.Range("O" & 3).Value = "Greatest % Decrease"
      ws.Range("O" & 4).Value = "Greatest Total Volume"
      ws.Range("L" & 1).Value = "Total Stock Volume"


    ' Loop through each cell in the resulting summary table range to find the Max % Increase
    For Each cell In rng
        If cell.Value > MaxPctChg Then
            MaxPctChg = cell.Value
            TickerSymbol = cell.Offset(0, -2)
        End If
    Next cell
    
    'Output and format the summary MaxPctChg for the quarter
    ws.Range("P2").Value = TickerSymbol
    ws.Range("Q2").Value = MaxPctChg
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ' Loop through each cell in the resulting summary table range to find the Max % Decline
    For Each cell In rng
        If cell.Value < MinPctChg Then
            MinPctChg = cell.Value
            TickerSymbol = cell.Offset(0, -2)
        End If
    Next cell
    
    'Output and format the Max percent decline
    ws.Range("P3").Value = TickerSymbol
    ws.Range("Q3").Value = MinPctChg
    ws.Range("Q3").NumberFormat = "0.00%"
    
    SummaryRows = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    Set rng = ws.Range("L2:L" & SummaryRows) ' Update the range as needed
    
    ' Loop through each cell in the resulting summary table range to find the Max Volume
    For Each cell In rng
        If cell.Value > MaxVolume Then
            MaxVolume = cell.Value
            TickerSymbol = cell.Offset(0, -3)
        End If
    Next cell
    
    'Output and format the Max volume value
    ws.Range("P4").Value = TickerSymbol
    ws.Range("Q4").Value = MaxVolume

Next ws

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
