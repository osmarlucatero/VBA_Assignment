# VBA_assignment
 ' Create loop for all worksheets

 For Each ws In Worksheets
 
' Assign variables

 Dim ticker As String
    Dim vol_sum As Double
    vol_sum = 0

    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double

' Name generation for columns that will be added
    
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "Yearly_change"
    ws.Cells(1, 12).Value = "Total Stock Vol"
    ws.Cells(1, 11).Value = "Yearly_percentage"

    Summary_Table_Row = 2
    
' Create loop for table
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow

' In case a value is = 0

      If year_open = 0 Then
          
         year_open = ws.Cells(i, 3).Value
         
      End If

' Defining where the data should map to based on their ticker category
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          ticker = ws.Cells(i, 1).Value
          
          year_close = ws.Cells(i, 6).Value
          yearly_change = year_close - year_open
          
' Changed to percent with formating in excel (could add *100 here if not)
          yearly_percentage = (yearly_change / year_open)

          vol_sum = vol_sum + ws.Cells(i, 7).Value

'Map where the summary table values will go
          
          ws.Range("j" & Summary_Table_Row).Value = yearly_change
          ws.Range("I" & Summary_Table_Row).Value = ticker
          ws.Range("K" & Summary_Table_Row).Value = yearly_percentage
          ws.Range("L" & Summary_Table_Row).Value = vol_sum

          Summary_Table_Row = Summary_Table_Row + 1

' Reset value of ticker category before moving to next
          vol_sum = 0

      Else

          vol_sum = vol_sum + Cells(i, 7).Value

      End If

    Next i
    
 ' Create labels for calculations

     ws.Cells(2, 14).Value = "Greatest % Increase"
     ws.Cells(3, 14).Value = "Greatest % Decrease"
     ws.Cells(4, 14).Value = "Greatest Volume"

' Calculate max and min for summary table

     ws.Cells(2, 15).Value = Application.WorksheetFunction.Max(ws.Range("K2:K3001"))
     ws.Cells(3, 15).Value = Application.WorksheetFunction.Min(ws.Range("K2:K3001"))
     ws.Cells(4, 15).Value = Application.WorksheetFunction.Max(ws.Range("L2:L3001"))
 Next ws

End Sub
