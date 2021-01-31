Attribute VB_Name = "Module3"
Sub stock_analysis_homework()
'--------------------
' SET ALL VARIABLES
'--------------------

 ' set variable for holding the stock symbol
 Dim stock_name As String
 
 'set variable for holding stock value
 Dim stock_value As Double
 stock_value = 0
 
 'set variable stock volume
 Dim stock_volume As Long
 stock_volume = 0
 
 'set variables for values
 Dim year_open As Double
 Dim year_close As Double
 Dim yearly_change As Double
 
 year_open = 0
 year_close = 0
 yearly_change = 0
 
 'keep track of summary table
 Dim summary_table_row As Integer
 summary_table_row = 2
 
 ' find last row
 Dim last_row As Long
 last_row = Cells(Rows.Count, 1).End(x1up).Row
 
'--------------------
'SET LOOP FOR ALL SHEETS
'--------------------

 ' loop through all sheets
 For Each ws In Worksheets

 ' add summary table to each sheet
 ws.Cells(1, 9).Value = "ticker"
 ws.Cells(1, 10).Value = "yearly change"
 ws.Cells(1, 11).Value = "percent change"
 ws.Cells(1, 12).Value = "total stock volume"
 
 ' add conditional formatting
 Dim my_range As Range
 Dim cond1 As FormatCondition, cond2 As FormatCondition
 Set my_range = Columns("K")
 Set cond1 = my_range.FormatConditions.Add(xlCellValue, xlGreater, "0")
 Set cond2 = my_range.FormatConditions.Add(xlCellValue, x1Less, "0")
 
 'define rules for formatting
 With cond1
 .Interior.Color = vbGreen
 .Font.Color = vbWhite
 End With
 
 With cond2
 .Interior.Color = vbRed
 .Font.Color = vbWhite
 End With
 
 ' add percentage formatting
 ws.Columns("K").NumberFormat = "0.00%"
 
'--------------------
'RUN LOOPS FOR STOCK VARIABLES
'--------------------

 ' loop through all stock dates
 For i = 2 To last_row
 
  'check if we are still in the same stock
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
   ' set the stock name
   stock_name = Cells(i, 1).Value
   
   ' add to the volume total
   stock_volume = stock_volume + Cells(i, 7).Value
   
   ' add year close stock value
   year_close = year_close + Cells(i, 6).Value
   
   ' caluclate stock value
   stock_value = (year_open - year_close)
   
   ' calculate yearly change in stock value
   yearly_change = ((year_open - year_close) / 100)
   
   ' print the stock symbol in the summary table
   Range("I" & summary_table_row).Value = stock_name
   
   ' print the stock value change in the summary table
   Range("J" & summary_table_row).Value = stock_value
   
   ' print the stock percentage in the summary table
   Range("K" & summary_table_row).Value = yearly_change
   
   ' print stock volume in the summary table
   Range("L" & summary_table_row).Value = stock_volume
   
   ' add one to summary table row
   summary_table_row = summary_table_row + 1
   
   ' reset values
   stock_volume = 0
   year_open = 0
   year_close = 0
   yearly_change = 0
   
  Else
  
   ' add to the stock volume total
   stock_volume = stock_volume + Cells(i, 7).Value
    
    If Cells(i, 2).Value = "*0101" Then
    
    ' set year open value
    year_open = year_open + Cells(i, 6).Value
    
    End If
    
  End If
  
 Next i

Next ws

End Sub
