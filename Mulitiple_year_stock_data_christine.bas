Attribute VB_Name = "Module1"
Sub Challenge()
    Dim ws As Worksheet
    'Generate in each worksheets
    For Each ws In Worksheets
        ws.Activate
        Dim i, lastrow As Long
        Dim Row As Integer
        Dim Ticker As String
        Dim date_open, date_close As Double
        
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim volume As Double
        Dim first_alpha  As Long
        
        
        Row = 1
        first_alpha = 1
        volume = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
       Range("I1") = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
      
        
        
    
        
            For i = 2 To lastrow
           
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Row : For answers position
                    Row = Row + 1
                
                ' each ticker first place
                    first_alpha = first_alpha + 1
                
                ' Q1: Ticker Symbol
                    Cells(Row, 9) = Cells(i, 1).Value
                
                ' open and close price of each symbol
                    date_open = Cells(first_alpha, 3).Value
                    date_close = Cells(i, 6).Value
                 
                 'calculate each ticker's yealy change and percentage change
                    If date_open = 0 Then
                    Yearly_Change = date_close - date_open
                    Percentage_Change = close_date
                    Else
                        Yearly_Change = date_close - date_open
                        Percentage_Change = Yearly_Change / date_open
                    End If
                
                    
                'start place to end place of each symbol
                    For j = first_alpha To i
                    ' volume of stock
                        volume = volume + Cells(j, 7).Value
                    Next j
                
                    Cells(Row, 10) = Yearly_Change
                    Cells(Row, 11) = Percentage_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                    Cells(Row, 12) = volume
                    ' Turn volume, Yearly_change, Percentage_Change to 0 for next calculation
                    volume = 0
                    Yearly_Change = 0
                    Percentage_Change = 0
                    ' turn first_alpha value to the last number of each symbol for next calculation
                    first_alpha = i
                
                
                End If
            Next i
            '-------------------------------------
            'Bonus
            'Greatest increase, decrease and volume table
            
            Dim g As Long
            Dim rise, reduce, great_volume As Double
            Dim now_change, previous_change As Double
            Dim now_volume, previous_volume As Double
            
            'Ticker name
            Dim rise_alpha, reduce_alpha As String
            Dim great_alpha As String
           
            Range("P1") = "Ticker"
            Range("Q1") = "Value"
            Range("O2") = "Greatest % Increase"
            Range("O3") = "Greatest % Decrease"
            Range("O4") = "Greatest Total Volume"
            rise = 0
            reduce = 0
            great_volume = 0
            lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
            
            
            'g starts from 3 to compare previous one
            For g = 3 To lastrow2
                 
                 'Assign cells value to varible
                 now_change = Cells(g, 11).Value
                 previous_change = Cells(g - 1, 11).Value
                 now_volume = Cells(g, 12).Value
                 previous_volume = Cells(g - 1, 12).Value
                 
                 ' find geatest increase ticker name and value
                 If rise > now_change And rise > previous_change Then
                    rise = rise
                ElseIf now_change > rise And now_change > previous_change Then
                    rise = now_change
                    rise_alpha = Cells(g, 9).Value
                ElseIf previous_change > rise And previous_change > now_change Then
                    rise = previous_change
                    rise_alpha = Cells(g - 1, 9).Value
                
                End If
                
                'find geatest decrease ticker name and value
                If reduce < now_change And reduce < previous_change Then
                    reduce = reduce
                
                ElseIf now_change < reduce And now_change < previous_change Then
                    reduce = now_change
                    reduce_alpha = Cells(g, 9).Value
                ElseIf previous_change < reduce And previous_change < now_change Then
                    reduce = previous_change
                    reduce_alpha = Cells(g - 1, 9).Value
                
                End If
                
                ' find geatest volume ticker name and value
                If great_volume > now_volume And great_volume > previous_volume Then
                    great_volume = great_volume
                
                ElseIf now_volume > great_volume And now_volume > previous_volume Then
                    great_volume = now_volume
                    great_alpha = Cells(g, 9).Value
                
                ElseIf previous_volume > great_volume And previous_volume > now_volume Then
                great_volume = previous_volume
                great_alpha = Cells(g - 1, 9).Value
                
                End If
            Next g
                
                Range("P2") = rise_alpha
                Range("P3") = reduce_alpha
                Range("P4") = great_alpha
                Range("Q2") = rise
                Range("Q2").NumberFormat = "0.00%"
                Range("Q3") = reduce
                Range("Q3").NumberFormat = "0.00%"
                Range("Q4") = great_volume
                
                
            Dim c As Long
            '-----------
            ' format cells color of Yearly Change's column
            For c = 2 To lastrow
                If Cells(c, 10).Value > 0 Then
                'ColorIndex 4 = green
                    Cells(c, 10).Interior.ColorIndex = 4
                ElseIf Cells(c, 10).Value < 0 Then
                'ColorIndex 3 = red
                    Cells(c, 10).Interior.ColorIndex = 3
                    
                End If
            Next c
                    
                
                
                
                
                
                
                 
        
        
        
            
       
            
            
            
            
            
            
           
      
    Next
End Sub



