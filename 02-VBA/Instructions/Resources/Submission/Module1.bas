Attribute VB_Name = "Module1"
Sub easy_macro()

    Dim Col As Double
    Dim Total_Vol As Double
    Dim WS_Count, WS_index As Integer
    Dim sheet As Worksheet

    ' Set WS_Count equal to the number of worksheets in the workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    'Loop through all sheets in the workbook
    For WS_index = 1 To WS_Count

        Sheets(WS_index).Activate
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"
        Col = 2
        Cells(Col, 9).Value = Cells(Col, 1).Value

        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For Row = 2 To LastRow
            If Cells(Row, 1).Value = Cells(Col, 9) Then
                Total_Vol = Total_Vol + Cells(Row, 7).Value
            Else
                Cells(Col, 10).Value = Total_Vol
                Total_Vol = Cells(Row, 7).Value
                Col = Col + 1
                Cells(Col, 9).Value = Cells(Row, 1).Value
            End If
        Next Row
     
        Cells(Col, 10).Value = Total_Vol
    Next WS_index
End Sub
Sub moderate_macro()
    Dim WS As Worksheet
    
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
      
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Col As Integer
        Col = 1
        Dim i As Long
        
       
        Open_Price = Cells(2, Col + 2).Value
         
        
        For i = 2 To LastRow

            If Cells(i + 1, Col).Value <> Cells(i, Col).Value Then
                
                Ticker_Name = Cells(i, Col).Value
                Cells(Row, Col + 8).Value = Ticker_Name
                
                Close_Price = Cells(i, Col + 5).Value
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Col + 9).Value = Yearly_Change
              
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Col + 10).Value = Percent_Change
                    Cells(Row, Col + 10).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + Cells(i, Col + 6).Value
                Cells(Row, Col + 11).Value = Volume
               
                Row = Row + 1
              
                Open_Price = Cells(i + 1, Col + 2)
                
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, Col + 6).Value
            End If
        Next i
        
        
        YCLastRow = WS.Cells(Rows.Count, Col + 8).End(xlUp).Row

        For j = 2 To YCLastRow
            If (Cells(j, Col + 9).Value > 0 Or Cells(j, Col + 9).Value = 0) Then
                Cells(j, Col + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Col + 9).Value < 0 Then
                Cells(j, Col + 9).Interior.ColorIndex = 3
            End If
        Next j
       
      
    Next WS
        
End Sub
Sub hard_macro()
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
       
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
     
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Col As Integer
        Col = 1
        Dim i As Long
        
    
        Open_Price = Cells(2, Col + 2).Value
         
        
        For i = 2 To LastRow
         
            If Cells(i + 1, Col).Value <> Cells(i, Col).Value Then
               
                Ticker_Name = Cells(i, Col).Value
                Cells(Row, Col + 8).Value = Ticker_Name
             
                Close_Price = Cells(i, Col + 5).Value
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Col + 9).Value = Yearly_Change
               
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Col + 10).Value = Percent_Change
                    Cells(Row, Col + 10).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + Cells(i, Col + 6).Value
                Cells(Row, Col + 11).Value = Volume
                
                Row = Row + 1
            
                Open_Price = Cells(i + 1, Col + 2)
                
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, Col + 6).Value
            End If
        Next i
        
        
        YCLastRow = WS.Cells(Rows.Count, Col + 8).End(xlUp).Row
      
        For j = 2 To YCLastRow
            If (Cells(j, Col + 9).Value > 0 Or Cells(j, Col + 9).Value = 0) Then
                Cells(j, Col + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Col + 9).Value < 0 Then
                Cells(j, Col + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        
        Cells(2, Col + 14).Value = "Greatest % Increase"
        Cells(3, Col + 14).Value = "Greatest % Decrease"
        Cells(4, Col + 14).Value = "Greatest Total Volume"
        Cells(1, Col + 15).Value = "Ticker"
        Cells(1, Col + 16).Value = "Value"
        
        For Z = 2 To YCLastRow
            If Cells(Z, Col + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Col + 15).Value = Cells(Z, Col + 8).Value
                Cells(2, Col + 16).Value = Cells(Z, Col + 10).Value
                Cells(2, Col + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Col + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Col + 15).Value = Cells(Z, Col + 8).Value
                Cells(3, Col + 16).Value = Cells(Z, Col + 10).Value
                Cells(3, Col + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Col + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Col + 15).Value = Cells(Z, Col + 8).Value
                Cells(4, Col + 16).Value = Cells(Z, Col + 11).Value
            End If
        Next Z
        
    Next WS

End Sub
