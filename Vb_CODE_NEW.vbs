Sub create_report()

    ' loop and calculate each sheet

    For Each ws In Worksheets

        Dim WorksheetName As String
        Dim TrickerName As String
        Dim OpenValue As Double
        Dim CloseValue As Double
        
        Dim source As Worksheet
        Dim destination As Worksheet
        
        Dim Year As String
        Dim J As Integer
        
        
        Dim Percent As Double
        Dim TotalStock As Double
        
        

        ' create sheet for each Year

        WorksheetName = ws.Name
        Year = "Report" & WorksheetName
        Sheets.Add.Name = Year
        
        ' wirte collumns headers

        Cells(1, 1).Value = "Ticker"
        Cells(1, 2).Value = "YearlyChange"
        Cells(1, 3).Value = "PercentChange"
        Cells(1, 4).Value = "Total Stock Volume"
        
        ' set current sheet as source and report sheet as destination

        Set source = ThisWorkbook.Sheets(ws.Name)
        Set destination = ThisWorkbook.Sheets(Year)
        
        ' keep the last row and last column of the source sheet
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' set the active page

        Worksheets(ws.Name).Activate
        
        ' keep these values as default

        TrickerName = Cells(2, 1).Value
        OpenValue = Cells(2, 3)
        CloseValue = Cells(2, 6)
        
        J = 2
        
        TotalStock = Cells(2, 7)
        Percent = 0
        
        'loop on all the rows and calculate percentage

        For i = 2 To LastRow
                
                If source.Cells(i + 1, 1).Value <> source.Cells(i, 1).Value Then
                    
                    CloseValue = Cells(i, 6)
                    If OpenValue = 0 Then
                        Percent = 0
                    Else
                        Percent = Round(((CloseValue - OpenValue) * 100) / OpenValue, 2)
                    End If
                    
                    destination.Cells(J, 1).Value = TrickerName
                    destination.Cells(J, 2).Value = CloseValue - OpenValue
                    destination.Cells(J, 3).Value = Percent
                    destination.Cells(J, 4).Value = TotalStock
                    If Percent < 0 Then
                       destination.Cells(J, 2).Interior.ColorIndex = 3
                    Else
                       destination.Cells(J, 2).Interior.ColorIndex = 4
                    End If
                    
                    
                    TrickerName = source.Cells(i + 1, 1).Value
                    OpenValue = source.Cells(i + 1, 3)
                    CloseValue = source.Cells(i + 1, 6)
                    TotalStock = Cells(2, 7)
                    Percent = 0
                    
                    
                    J = J + 1
                    
                Else
                    TotalStock = TotalStock + source.Cells(i, 7)
                    CloseValue = source.Cells(i, 6)
                
                End If
                

        Next i

        ' set default values for Greatestincrease and GreatestDecrease and Greatestvolume
        
        Greatestincrease = destination.Cells(2, 3).Value
        TickerGreatestincrease = destination.Cells(2, 1).Value
        
        GreatestDecrease = destination.Cells(2, 3).Value
        TickerGreatestDecrease = destination.Cells(2, 1).Value
        
        Greatestvolume = destination.Cells(2, 4).Value
        TickerGreatestvolume = destination.Cells(2, 1).Value
        
        
        
        LastRow = destination.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' calculate Greatestincrease and GreatestDecrease

        For i = 2 To LastRow
            

                If destination.Cells(i, 3).Value > Greatestincrease Then
                    Greatestincrease = destination.Cells(i, 3).Value
                    TickerGreatestincrease = destination.Cells(i, 1).Value
                End If
                

                
                If destination.Cells(i, 3).Value < GreatestDecrease Then
                    GreatestDecrease = destination.Cells(i, 3).Value
                    TickerGreatestDecrease = destination.Cells(i, 1).Value
                End If
                

            
            If destination.Cells(i, 4).Value > Greatestvolume Then
                Greatestvolume = destination.Cells(i, 4).Value
                TickerGreatestvolume = destination.Cells(i, 1).Value
            End If
                
                
        
        Next i
        
        
        destination.Cells(2, 6).Value = "Greatest % increase"
        destination.Cells(2, 7).Value = TickerGreatestincrease
        destination.Cells(2, 8).Value = Greatestincrease
        
        destination.Cells(3, 6).Value = "Greatest % Decrease"
        destination.Cells(3, 7).Value = TickerGreatestDecrease
        destination.Cells(3, 8).Value = GreatestDecrease
        
        destination.Cells(4, 6).Value = "Greatest total volume"
        destination.Cells(4, 7).Value = TickerGreatestvolume
        destination.Cells(4, 8).Value = Greatestvolume
        
        
    Next ws


End Sub



