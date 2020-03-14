Sub VBA_Wall_Street()

 For Each ws In Worksheets
 

        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim GreatestIncreaseTicker As String
        Dim GreatestPercentageIncrease As Double
        Dim GreatestDecreaseTicker As String
        Dim GreatestPercentageDecrease As Double
        Dim GreatestTotalVolumeTicker As String
        Dim GreatestTotalVolumeValue As Double
        
        TotalStockVolume = 0
        GreatestIncreaseTicker = ""
        GreatestPercentageIncrease = 0
        GreatestDecreaseTicker = ""
        GreatestPercentageDecrease = 0
        GreatestTotalVolumeTicker = ""
        GreatestTotalVolumeValue = 0
    
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(2, 1).Value
        
        OpenPrice = ws.Cells(2, 3).Value
        

        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            ClosePrice = ws.Cells(i, 6).Value
            
            YearlyChange = ClosePrice - OpenPrice
            
            PercentChange = YearlyChange / OpenPrice
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
            
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
            ws.Range("K" & Summary_Table_Row).Value = PercentChange

            Summary_Table_Row = Summary_Table_Row + 1
            
            TotalStockVolume = 0
            
            Else
            
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
            End If
            
            If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
            
            If PercentChange > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = PercentChange
                GreatestIncreaseTicker = ws.Cells(i, 1).Value
            End If
            If PercentChange < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = PercentChange
                GreatestDecreaseTicker = ws.Cells(i, 1).Value
            End If
            If TotalStockVolume > GreatestTotalVolumeValue Then
                GreatestTotalVolumeValue = TotalStockVolume
                GreatestTotalVolumeTicker = ws.Cells(i, 1).Value
            End If
            
        Next i
            ws.Range("O2").Value = GreatestIncreaseTicker
            ws.Range("O3").Value = GreatestDecreaseTicker
            ws.Range("O4").Value = GreatestTotalVolumeTicker
            ws.Range("P2").Value = GreatestPercentageIncrease
            ws.Range("P3").Value = GreatestPercentageDecrease
            ws.Range("P4").Value = GreatestTotalVolumeValue
            
            ws.Range("A1:S1").EntireColumn.AutoFit
            ws.Range("A1:S1").Font.Bold = True
            ws.Columns(11).NumberFormat = "0.00%"
            ws.Columns(12).NumberFormat = "#,#00.0#"
            ws.Range("P2:P3").NumberFormat = "0.00%"
            ws.Range("P4").NumberFormat = "#,#00.0#"
  Next ws
    
   
    
End Sub


