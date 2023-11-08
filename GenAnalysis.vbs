Attribute VB_Name = "GenAnalysis"

Sub GenAnalysis()
'Module Two VBA-Challenge
'John Ellis
'2023-11-09
'
'Add summary tabs for each year (<year> Summary) which:
'Add Columns for:
'    -<ticker symbol>,
'    -<Yearly Change> (closing price at end of year less the closing price at the beginning of the year)
'    -<Percentage change> from opening price and closing price at end of the year.
'    -<total stock volume> for the year
'    -will have one record (row) per stock per year
'    -add conditional formatting to the Yearly Change and the Percentage Change columns which is fills the cell in RED for negative volumes and fills the cell with GREEN for positive values
'    -for each year identify
'       - the stock with the Greatest % increase
'       - the stock with the Greatest % decrease
'       - the stock with the Greatest total volume.
'

'Set up the variables

Dim dStart, dStop As Date
Dim dataSheet As Worksheet
Dim numSheets, sheetCount, dataRowCount, summaryRowCount, dataLastRow, summaryLastRow, CurrentCloseDate, CurrentOpenDate, RED, GREEN As Integer
Dim CurrentStock, sMaxUp, sMaxDwn, sMaxVol As String
Dim YearStockVol, vMaxVol, VmaxDwn, VMaxUp, YearChg, PerChg, YearStockOpen, YearStockClose, maxDate, minDate As Double
dStart = Time 'Log the start time'

'Set Color Values
RED = 3
GREEN = 4


numSheets = ThisWorkbook.Worksheets.Count 'store the workbook's work sheet count.
'Application.ScreenUpdating = False

'Iterate through each worksheet in the workbook
For sheetCount = 1 To numSheets

    dataLastRow = ThisWorkbook.Worksheets(sheetCount).Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set up stock summary table
    ThisWorkbook.Worksheets(sheetCount).Range("I1").Value = "Ticker"
    ThisWorkbook.Worksheets(sheetCount).Range("J1").Value = "Yearly Change"
    ThisWorkbook.Worksheets(sheetCount).Range("K1").Value = "Percent Change"
    ThisWorkbook.Worksheets(sheetCount).Range("L1").Value = "Total Stock Volume"
   
    'Set up high performers
    ThisWorkbook.Worksheets(sheetCount).Range("P1").Value = "Ticker"
    ThisWorkbook.Worksheets(sheetCount).Range("Q1").Value = "Value"
    ThisWorkbook.Worksheets(sheetCount).Range("O2").Value = "Greatest % Increase"
    ThisWorkbook.Worksheets(sheetCount).Range("O3").Value = "Greatest % Decrease"
    ThisWorkbook.Worksheets(sheetCount).Range("O4").Value = "Greatest Total Volume"
    summaryLastRow = ThisWorkbook.Worksheets(sheetCount).Cells(Rows.Count, 9).End(xlUp).Row
    
    'Set up variables for storing page data
    YearStockVol = 0 'set the yearly volume to 0
    YearStockOpen = 0 'set the stock's open value for the year to 0
    YearStockClose = 0 'set the stock's close value for the year to 0
    CurrentCloseDate = 0 'set the  close date for the year to 0
    CurrentOpenDate = 9999999999# 'set open date for the year to 9999999999
    VMaxUp = 0 'Set the stored Value of the stock with the Maximum Price increase to 0
    VmaxDwn = 9999999999# 'Set the stored Value of the stock with the Maximum Price Decrease to 0
    vMaxVol = 0 'Set the value of the stored Volume of the stock with the Most stock traded to 0
    summaryRowCount = 2
    
    For dataRowCount = 2 To dataLastRow
        CurrentStock = ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 1).Value 'store the stock ticker
        
        If CLng(ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 2).Value) > CurrentCloseDate Then 'check if the date on the data row is more recent than the one stored
            YearStockClose = ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 6).Value 'if it is update the Close Value for the Year
            CurrentCloseDate = CLng(ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 2).Value) 'set the close date the current data row's close date.
        End If
        
        If CLng(ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 2).Value) < CurrentOpenDate Then 'check if the date on the data row is older than the one stored
            YearStockOpen = ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 3).Value 'if it is update the open value with for the Year.
            CurrentOpenDate = CLng(ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 2).Value) 'set the open date to the current data row's open date.
        End If
        
        YearStockVol = YearStockVol + ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount, 7).Value 'Add the current rows volume to the stored volume for the current stock.
        
        If CurrentStock <> ThisWorkbook.Worksheets(sheetCount).Cells(dataRowCount + 1, 1).Value Then 'No More Stock for this Ticker so write out the stocks yearly data
            
            ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 9).Value = CurrentStock 'Write the stock ticker to the summary table
            
            With ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 11) 'Writeout the stock's the Percent change for the stock to the summary table and format the cell.
                .Value = YearStockClose / YearStockOpen - 1 'Calculate the Yearly % Change
                .NumberFormat = "0.00%"
            End With
            
            ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 12).Value = YearStockVol 'Writeout the Stock's volume for the year
            
            With ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 10) 'Format the Cells 'Writeou the Stock's Change for the year and format the cells.
                .Value = YearStockClose - YearStockOpen 'Calculate the yearly change.
                .NumberFormat = "0.00"
      
                If YearStockClose < YearStockOpen Then 'if the stock dropped in value formate the change value to RED otherwise Green
                    .Interior.ColorIndex = RED
                Else
                    .Interior.ColorIndex = GREEN
                End If
            
            End With
            
            'Reset the tracked values for the current stock to starting values
            
            YearStockVol = 0
            CurrentCloseDate = 0
            CurrentOpenDate = 9999999999#
            'Check for Top Perform Stocks ans store the value's
            
            If ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 11).Value > VMaxUp Then 'if the current stock's %Change is Greater than what's stored update what's stored
                VMaxUp = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 11).Value
                sMaxUp = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 9).Value
            End If
            If ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 11).Value < VmaxDwn Then ' if the current stock's % Change is less that whats stored then update what's stored
                VmaxDwn = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 11).Value
                sMaxDwn = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 9).Value
            End If
            If ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 12) > vMaxVol Then 'if the current stock's volume is greater than what's store then update what's stored.
                vMaxVol = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 12).Value
                sMaxVol = ThisWorkbook.Worksheets(sheetCount).Cells(summaryRowCount, 9).Value
            End If
              
            summaryRowCount = summaryRowCount + 1
            
        End If
            
         
    Next dataRowCount
    

        
'    Next summaryRowCount
    
    'Write out the high performer results to the worksheet and formate the cells
    ThisWorkbook.Worksheets(sheetCount).Range("P2").Value = sMaxUp
    With ThisWorkbook.Worksheets(sheetCount).Range("Q2")
        .Value = VMaxUp
        .NumberFormat = "0.00%"
    End With
    
    ThisWorkbook.Worksheets(sheetCount).Range("P3").Value = sMaxDwn
    With ThisWorkbook.Worksheets(sheetCount).Range("Q3")
        .Value = VmaxDwn
        .NumberFormat = "0.00%"
    End With
    
    ThisWorkbook.Worksheets(sheetCount).Range("P4").Value = sMaxVol
    ThisWorkbook.Worksheets(sheetCount).Range("Q4").Value = vMaxVol
    
    ThisWorkbook.Worksheets(sheetCount).Cells.EntireColumn.AutoFit 'Clean up the column widths
          
Next sheetCount 'go to the next sheet
'once all sheets are processed pop a message with the processing time
dStop = Time
'Application.ScreenUpdating = True
MsgBox ("Done! in " & Format(dStop - dStart, "HH:MM:SS"))


End Sub

