Sub Market_exchange()
 
 Dim WorksheetName As String
    Dim Ticker As String
    Dim Opening As Double
    Dim Yearly_change As Double
    Dim Percent_Change As Double
    Dim j As Integer
    Dim i As Long
    Dim Total_Stock_Volume As Double
    
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestTotalVolumeTicker As String
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    
 For Each ws In Worksheets
                            
                            
    'Print the headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
                            
                            
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Yearly_change = 0
    
    Percentage_Changed = 0
    
    Opening = ws.Cells(2, 3).Value
    
    Total_Stock_Volume = 0
    
    j = 0
    
    For i = 2 To LastRow
    
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            Yearly_change = ws.Cells(i, 6).Value - Opening
            
          'Percentage caluculation
        If Opening <> 0 Then
        
            Percent_Change = (Yearly_change / Opening)
            
        Else: Percent_Change = 0
            
        End If
            
          ' output values
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
          
            ws.Range("j" & 2 + j).Value = Yearly_change
            
            ws.Range("k" & 2 + j).Value = Percent_Change
          
            ws.Range("l" & 2 + j).Value = Total_Stock_Volume
            
            ws.Range("k" & 2 + j).NumberFormat = "0.00%"
            
            ws.Range("q" & 2 + j).NumberFormat = "0.00%"
            
          'Color code
        Dim YearlyChangeCell As Range
                
        Set YearlyChangeCell = ws.Range("J" & 2 + j)
       
        If Yearly_change >= 0 Then
                    
                YearlyChangeCell.Interior.ColorIndex = 4
        Else
                
                YearlyChangeCell.Interior.ColorIndex = 3
        
        End If
        
            
           
            ' greatest % increase and % decrease
        If Percent_Change > GreatestIncrease Then
                
            GreatestIncrease = Percent_Change
                
            GreatestIncreaseTicker = ws.Cells(i, 1).Value
            
        ElseIf Percent_Change < GreatestDecrease Then
                
            GreatestDecrease = Percent_Change
                
            GreatestDecreaseTicker = ws.Cells(i, 1).Value
            
        End If
                
                ' greatest total volume
            
        If Total_Stock_Volume > GreatestTotalVolume Then
                
            GreatestTotalVolume = Total_Stock_Volume
                
            GreatestTotalVolumeTicker = ws.Cells(i, 1).Value
            
        End If
           
        
          ' incriment variables
          Total_Stock_Volume = 0
          
          Yearly_change = 0
          
          Opening = ws.Cells(i + 1, 3).Value
          
          j = j + 1
          
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        
        End If
        
    Next i
    
        ws.Range("P2").Value = GreatestIncreaseTicker
        
        ws.Range("P3").Value = GreatestDecreaseTicker
        
        ws.Range("P4").Value = GreatestTotalVolumeTicker
        
        ws.Range("Q2").Value = GreatestIncrease
        
        ws.Range("Q3").Value = GreatestDecrease
        
        ws.Range("Q4").Value = GreatestTotalVolume

    
 Next ws
 
    
End Sub

