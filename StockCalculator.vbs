Attribute VB_Name = "Module1"

Public Sub Stock_Calculator()

    Dim ws As Worksheet
    Dim rng As Range
    Dim Ticker_Name As String
    Dim Volume_Total As Double
    Dim Summary_Total As Integer
    Dim Close_Price As Double
    Dim Open_Price As Double
    Dim Max_Volume As Double
    Dim G_Increse As Double
    Dim G_Decrease As Double
    

    For Each ws In ThisWorkbook.Worksheets
   
        With ws
        
        Volume_Total = 0
        Summary_Total = 2
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To last_row

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                       
            Ticker_Name = ws.Cells(i, 1).Value
                   
                If Summary_Total = 2 Then

                    Open_Price = ws.Cells(2, 3)
                    
                Else
                
                    Open_Price = ws.Cells(i + 1, 3).Value
                
                End If
        
                If Open_Price = 0 Then
            
                   ws.Range("K" & Summary_Total).Value = 0
                   
                Else
                
                   ws.Range("K" & Summary_Total).Value = ((Open_Price - Close_Price) / Open_Price)
                   
                End If
                
                        
            Close_Price = ws.Cells(i, 6).Value
            
            ws.Range("J" & Summary_Total).Value = Open_Price - Close_Price
            
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                   
            ws.Range("I" & Summary_Total).Value = Ticker_Name
                   
            ws.Range("L" & Summary_Total).Value = Volume_Total
                
                If ws.Range("J" & Summary_Total).Value > 0 Then
                        
                    ws.Range("J" & Summary_Total).Interior.Color = vbGreen
                        
                Else
                    
                    ws.Range("J" & Summary_Total).Interior.Color = vbRed
                        
                End If
                                                 
          
        
            Summary_Total = Summary_Total + 1
        
            Volume_Total = 0
                    
            Else
            
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                         
            End If
        
        Next i
        
    ws.Range("K:K").Style = "Percent"
            
    G_Increse = ws.Application.WorksheetFunction.Max(ws.Range("K:K"))
    
    G_Decrease = ws.Application.WorksheetFunction.Min(ws.Range("K:K"))

    Max_Volume = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    ws.Range("P2").Value = G_Increse
    
    ws.Range("P3").Value = G_Decrease
    
    ws.Range("P2:P3").Style = "Percent"
    
    ws.Range("P4").Value = Max_Volume

    ws.Range("I1").Value = "Ticker"
    
    ws.Range("J1").Value = "Yearly Change"
    
    ws.Range("K1").Value = "Percent Change"
    
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O4").Value = "Greates Total Volume"
    
    ws.Range("O3").Value = "Greates % Decrease"
    
    ws.Range("O2").Value = "Greates % Increase"
    
    ws.UsedRange.Columns.AutoFit
    
    End With
     
       
    Next ws

    
End Sub
