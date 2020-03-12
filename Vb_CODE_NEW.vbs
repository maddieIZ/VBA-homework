Sub create_report()
        
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
        
        

        
        WorksheetName = ws.Name
        Year = "Report" & WorksheetName
        Sheets.Add.Name = Year
        
        Cells(1, 1).Value = "Ticker"
        Cells(1, 2).Value = "YearlyChange"
        Cells(1, 3).Value = "PercentChange"
        Cells(1, 4).Value = "Total Stock Volume"
        
        
        Set source = ThisWorkbook.Sheets(ws.Name)
        Set destination = ThisWorkbook.Sheets(Year)
        
       
    
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        Worksheets(ws.Name).Activate
        
        TrickerName = Cells(2, 1).Value
        OpenValue = Cells(2, 3)
        CloseValue = Cells(2, 6)
        
        J = 2
        
        TotalStock = Cells(2, 7)
        Percent = 0
        
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
        
        Greatestincrease = destination.Cells(2, 3).Value
        TickerGreatestincrease = destination.Cells(2, 1).Value
        
        GreatestDecrease = destination.Cells(2, 3).Value
        TickerGreatestDecrease = destination.Cells(2, 1).Value
        
        Greatestvolume = destination.Cells(2, 4).Value
        TickerGreatestvolume = destination.Cells(2, 1).Value
        
        
        
        LastRow = destination.Cells(Rows.Count, 1).End(xlUp).Row
        
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



