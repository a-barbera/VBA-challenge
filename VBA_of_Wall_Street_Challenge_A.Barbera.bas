Attribute VB_Name = "Module1"
Sub Ticker_PullandConsolidate()
    'BONUS apply module to multiple worksheets
    
   For Each WS In Worksheets
   WS.Activate
    
    'Header formatting
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

'format these columns

    Range("I1:P1").EntireColumn.AutoFit
    Range("I1:P1").Font.Bold = True
    Range("N1:N4").Font.Bold = True
    Range("K2:K1000").NumberFormat = "0.00%"

'main sub formula

    Dim I As Long
    Dim j As Long
    Dim lRow As Long
    Dim lCol As Long
    Dim TotalStock 'too long to be a long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim CurrentTicker As String
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column

    j = 2
    TotalStock = 0
    CurrentTicker = Cells(2, 1).Value

        For I = 2 To lRow
    
            If Cells(I, 1).Value = CurrentTicker Then
            TotalStock = TotalStock + Cells(I, 7)
            End If
    
            If Cells(I - 1, 1).Value <> CurrentTicker Then
            OpenPrice = Cells(I, 3).Value 'ending this container
    
            End If
    
            If Cells(I + 1, 1).Value <> CurrentTicker Then
            Cells(j, 9).Value = CurrentTicker
            Cells(j, 12).Value = TotalStock
            TotalStock = 0
            CurrentTicker = Cells(I + 1, 1).Value
            ClosePrice = Cells(I, 6).Value
            Cells(j, 10).Value = ClosePrice - OpenPrice
            Cells(j, 11).Value = (ClosePrice - OpenPrice) / OpenPrice
            j = j + 1
    
            End If
  
            Next I
    
        'BONUS finding the greatest increase, decrease, and total stock volume
    
    Dim GrPIn As Double
    Dim GrPDec As Double
    Dim GrTotVol As Double
    

    GrPIn = 0
    GrPDec = 0
    GrTotVol = 0
        
        For I = 2 To lRow
        
        If Cells(I, 11).Value > GrPIn Then
        GrPIn = Cells(I, 11).Value
        'set ticker to column q
        Range("O2").Value = Cells(I, 9).Value
        
        'set value to column r
        Range("P2").Value = FormatPercent(GrPIn)
        
        ElseIf Cells(I, 11).Value < GrPDec Then
        GrPDec = Cells(I, 11).Value
        
        'set ticker to column q
        Range("O3").Value = Cells(I, 9).Value
        
        'set value to column r
        Range("P3").Value = FormatPercent(GrPDec)
        
        End If
        
        If Cells(I, 12).Value > GrTotVol Then
        GrTotVol = Cells(I, 12).Value
        Range("o4").Value = Cells(I, 9).Value
        Range("P4").Value = Cells(I, 12).Value
        
        End If
        

        Next I
        
        'CONDITIONAL FORMATTING
        
    lRow = Cells(Rows.Count, 10).End(xlUp).Row
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column

        For I = 2 To lRow

            If Cells(I, 10).Value >= 0 Then
            Cells(I, 10).Interior.ColorIndex = 4

            ElseIf Cells(I, 10).Value <= 0 Then
            Cells(I, 10).Interior.ColorIndex = 3
            
            Else: Cells(I, 10).Value = blank
            Cells(I, 10).Interior.ColorIndex = 0
            End If
                                                  

        Next I

    Next WS
    
End Sub


