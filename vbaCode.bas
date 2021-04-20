Attribute VB_Name = "Module1"
Sub MultiYearStockTicker()
    For Each ws In Worksheets
  ' Set an initial variable for holding the stock name
  Dim Stock_Name As String

  ' Set an initial variable for holding the stock total per ticker
  Dim Stock_Total As Double
  Stock_Total = 0

  ' Keep track of the location for each stock ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  
  Dim Stock_Open As Double
  Dim Stock_Close As Double
  Dim Yearly_Change As Double
  
  Yearly_Change = 0
  Stock_Open = ws.Cells(2, 3).Value
  

Dim max As Double
Dim tag As String
Dim tag2 As String
Dim min As Double
Dim maxVolume As Double
Dim tag3 As String


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


  ' Loop through all stock tickers
  For i = 2 To LastRow

    ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Stock_Close = ws.Cells(i, 6).Value
    

      ' Set the Stock name
      Stock_Name = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      
      Yearly_Change = (Stock_Close - Stock_Open)

      ' Print the Stock Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Stock_Name

      ' Print the Stock amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total
      
      ws.Range("J" & Summary_Table_Row).Value = (Stock_Close - Stock_Open)
      
      If Stock_Open = 0 Then
      Stock_Open = 1
         
      ws.Range("K" & Summary_Table_Row).Value = (Yearly_Change / Stock_Open)
      
      Else
      
      ws.Range("K" & Summary_Table_Row).Value = (Yearly_Change / Stock_Open)
      End If
      
    
        
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_Total = 0
      
      Stock_Open = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

    End If

    
    
  Next i
  
  
  For j = 2 To LastRow
        If ws.Cells(j, 10).Value < 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(j, 10).Value > 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 4
    
    End If
    
    
  max = Application.WorksheetFunction.max(ws.Range("K:K"))
  min = Application.WorksheetFunction.min(ws.Range("K:K"))
  maxVolume = Application.WorksheetFunction.max(ws.Range("L:L"))
    

     If ws.Cells(j, 11).Value = max Then
       tag = ws.Cells(j, 11).Offset(0, -2).Value
       
    ws.Cells(2, 16) = max
    ws.Cells(2, 15).Value = tag
     End If



    
    If ws.Cells(j, 11).Value = min Then
       tag2 = ws.Cells(j, 11).Offset(0, -2).Value
       ws.Cells(3, 16) = min
       ws.Cells(3, 15).Value = tag2
     End If
   
  
    
   
    
    If ws.Cells(j, 12).Value = maxVolume Then
       tag3 = ws.Cells(j, 12).Offset(0, -3).Value
       
        ws.Cells(4, 16) = maxVolume
        ws.Cells(4, 15).Value = tag3
     End If
  
    
    Next j


Next ws
End Sub
