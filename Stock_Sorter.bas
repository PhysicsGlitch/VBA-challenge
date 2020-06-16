Attribute VB_Name = "Stock_Sorter"
Sub StockSorter():
    Dim Wksht_Count As Integer
    Dim Wksht_Iterator As Integer
    Dim TotalSV As Double
    Dim YearlyChange As Double
    Dim new_ticker As Long
    Dim emptyrow As Long
    Dim rowcounter As Long
    Dim rowcount As Long
             
    Wksht_Count = ActiveWorkbook.Worksheets.Count
          
    For Wksht_Iterator = 1 To Wksht_Count
                   
        Worksheets(Wksht_Iterator).Activate
        ActiveSheet.Cells.ClearFormats
        ActiveSheet.Range("I:Q").Columns.ClearContents
        
        'I ran into some format issues where if I ran the code twice it would break because VBA didn't like finding the greatest of a Percent value.
        'So I just cleared the formats and put a style sheet at the end where I could adjust everything about the style in one place.
        'Also I put a general clear for columns I to Q so that if the code was ran again it would reset everything instead of add data at the bottom of these columns.
        
        'Column and Row Headers
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        rowcount = ActiveSheet.UsedRange.Rows.Count
        
        TotalSV = 0
        'set the value for incrementing the sum of Total Stock Value to 0
        new_ticker = 2
        'this just counts the first row where the new ticker in the for loop is found. The first ticker starts at 2 since it is the first stock below the header.
        
        
        For rowcounter = 2 To rowcount
        
            If Cells(rowcounter, 1) <> Cells(rowcounter + 1, 1) Then
                TotalSV = Cells(rowcounter, 7).Value + TotalSV
                emptyrow = Range("I:I").CurrentRegion.Rows.Count
                'Instead of creating another variable to increment the row count of where to put the next line of data in our totals, like we did in class.
                'I just created a function with a variable emptyrow for excel to check the next empty row in the current region and then fill in data there.
                'The only consideration with this way is that if the module was run again it wouldn't overwrite the data so you'd need a line to clear these columns to run again.
                Cells(emptyrow + 1, 9).Value = Cells(rowcounter, 1).Value
                Cells(emptyrow + 1, 12).Value = TotalSV
                'Find first none 0 start for opening stock value
                'this function uses the worksheet function to to find the first entry that has an opening price greater than 0 for each ticker.
                'there are 2 possibilites here: if only one line of data is missing or if the info for an entire ticker is missing.
                'I wrote code for both possibilites since both occur in PLNT. In 2014 there is no PLNT info, while in 2015 it seems like it was a new stock halfway through the year
                'The basic logis is that if the Total is zero, you just need to overwrite the data as all zeros and then the code will work for both contingencies.
                    
                If Cells(new_ticker, 3).Value = 0 Then
                    Dim firstnon0row As Long
                    Dim non0iter As Long
                    Dim non0range As Range
                    
                    Set non0range = Range(Cells(new_ticker, 3), Cells(rowcount, 3))
                        For non0iter = new_ticker To rowcount
                            If Cells(non0iter, 3).Value > 0 Then
                            firstnon0row = non0iter
                            Exit For
                        End If
                        Next
                        
                        'This if statement just disambiguates the case of whether all info is missing for a stock or only the first few lines.
                        If TotalSV = 0 Then
                            Cells(emptyrow + 1, 10).Value = 0
                            Cells(emptyrow + 1, 11).Value = 0
                        Else
                            YearlyChange = Cells(rowcounter, 6).Value - Cells(firstnon0row, 3).Value
                            Cells(emptyrow + 1, 10).Value = YearlyChange
                            Cells(emptyrow + 1, 11).Value = YearlyChange / Cells(firstnon0row, 3).Value
                        End If
                        
                 Else
                 
                 
                    YearlyChange = Cells(rowcounter, 6).Value - Cells(new_ticker, 3).Value
                    Cells(emptyrow + 1, 10).Value = YearlyChange
                    Cells(emptyrow + 1, 11).Value = YearlyChange / Cells(new_ticker, 3).Value
                    
                End If
                
                
                    
                
                new_ticker = rowcounter + 1
                'This line just finds the first row where the new ticker begins.
                'Also, I didn't know if we were supposed to care if a stock started with 0 total volume. But I cross-checked the numbers in the readme
                'and the numbers in the example sheet calculated yearly change and percent change even if the new stock started with a total of 0 so I did the same.
                TotalSV = 0
                
            Else
                TotalSV = TotalSV + Cells(rowcounter, 7).Value
                             
            End If
        
        Next rowcounter
        
        'Find greatest increase and decrease function
                
        Dim maxSVrow As Long
        Dim maxincreaserow As Long
        Dim maxdecreaserow As Long
        Dim maxSVrg As Range
        Dim maxvaluesrg As Range
        
        Set maxSVrg = Range("L2", Range("L2").End(xlDown))
        Set maxvaluesrg = Range("K2", Range("K2").End(xlDown))
     
        maxSVrow = maxSVrg.Find(WorksheetFunction.Max(maxSVrg)).Row
        maxincreaserow = maxvaluesrg.Find(WorksheetFunction.Max(maxvaluesrg)).Row
        maxdecreaserow = maxvaluesrg.Find(WorksheetFunction.Min(maxvaluesrg)).Row
        
        Range("P4").Value = Cells(maxSVrow, 9).Value
        Range("Q4").Value = Cells(maxSVrow, 12).Value
        Range("P2").Value = Cells(maxincreaserow, 9).Value
        Range("Q2").Value = Cells(maxincreaserow, 11).Value
        Range("P3").Value = Cells(maxdecreaserow, 9).Value
        Range("Q3").Value = Cells(maxdecreaserow, 11).Value
        
        'Format Section: This is where I put the whole "Style Sheet" for the spreadsheet so I could change things without having to go back into the for loop.
        
        'Conditional formatting color conditions
        
        Dim colorrange As Range
        Dim redpercent As FormatCondition
        Dim greenpercent As FormatCondition
        
        Set colorrange = Range("J2", Range("J2").End(xlDown))
        Set greenpercent = colorrange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set redpercent = colorrange.FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        With greenpercent
            .Interior.Color = RGB(0, 128, 0)
            End With
        With redpercent
            .Interior.Color = RGB(255, 0, 0)
            End With
                        
        'Column width formatting
        
        Range("I:I").ColumnWidth = 10
        Range("J:J").ColumnWidth = 12
        Range("K:K").ColumnWidth = 14
        Range("L:L").ColumnWidth = 17
        Range("O:O").ColumnWidth = 20
        Range("P:P").ColumnWidth = 10
        Range("Q:Q").ColumnWidth = 13
           
        'change values to percents in final totals
        
        Range("K:K").NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        
                
    Next Wksht_Iterator
            
End Sub
