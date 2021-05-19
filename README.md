Sub Stockdata()

    ' Set CurrentWs as a worksheet object variable.
    Dim CurrentWs As Worksheet
    
    
    ' Loop through all of the worksheets in the active workbook.
    For Each CurrentWs In Worksheets
    
        ' Set Variables
        Dim Ticker_Name As String
        Ticker_Name = " "
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Price_Change As Double
        Price_Change = 0
        Dim Price_Change_Percent As Double
        Price_Change_Percent = 0
        Dim max_ticker_name As String
        max_ticker_name = " "
        Dim min_ticker_name As String
        min_ticker_name = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_volume_ticker As String
        max_volume_ticker = " "
        Dim max_volume As Double
        max_volume = 0
        '----------------------------------------------------------------
         
        ' Location for Variables
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set initial row count
        Dim Lastrow As Long
        
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        
        If Need_Summary_Table_Header Then
            ' Set Titles
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            
            Need_Summary_Table_Header = True
        End If
        
        ' Set initial value of Open_Price
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning of Current Worksheet
        For i = 2 To Lastrow
        
      
            ' Check we are within the same ticker,
                If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ' Set the ticker name
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                
                ' Calculate Price_Change and Price_Change_Percent
                Close_Price = CurrentWs.Cells(i, 6).Value
                Price_Change = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    Price_Change_Percent = (Price_Change / Open_Price) * 100
               
                End If
                
                ' Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("J" & Summary_Table_Row).Value = Price_Change
    
                If (Price_Change > 0) Then
                    
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Price_Change <= 0) Then
                    
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Price_Change_Percent) & "%")
                ' Print the Ticker Name in the Summary Table, Column J
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset
                Price_Change = 0
                Close_Price = 0
                ' Capture next Ticker's Open_Price
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                
                ' Calculations
                If (Price_Change_Percent > max_percent) Then
                    max_percent = Price_Change_Percent
                    max_ticker_name = Ticker_Name
                ElseIf (Price_Change_Percent < min_percent) Then
                    min_percent = Price_Change_Percent
                    min_ticker_name = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > max_volume) Then
                    max_volume = Total_Ticker_Volume
                    max_volume_ticker = Ticker_Name
                End If
                
                ' Reset
                Price_Change_Percent = 0
                Total_Ticker_Volume = 0
                
            
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            
      
        Next i

                     
                CurrentWs.Range("Q2").Value = (CStr(max_percent) & "%")
                CurrentWs.Range("Q3").Value = (CStr(min_percent) & "%")
                CurrentWs.Range("P2").Value = max_ticker_name
                CurrentWs.Range("P3").Value = min_ticker_name
                CurrentWs.Range("Q4").Value = max_volume
                CurrentWs.Range("P4").Value = max_volume_ticker
                
           
        
     Next CurrentWs
End Sub

