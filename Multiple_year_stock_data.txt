Sub stocks()
Dim i As Long
Dim j As Long
Dim ticker(1) As String
Dim cl As Double
Dim op As Double
Dim total
Dim diff As Double
Dim per As Double

For Each ws In Worksheets ' iterates through every worksheet

    ' formats colomn names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    NumRows = ws.Cells(Rows.Count, 1).End(xlUp).Row ' stores count of rows
    
    ' resets the variables
    j = 1
    perange = 0
    totrange = 0
    
    For i = 2 To NumRows ' iterates through every row
    
        If j = 1 Then
            ticker(0) = ws.Cells(i, 1).Value ' stores first ticker for comparison
            ticker(1) = ws.Cells(i, 1).Value
            total = total + ws.Cells(i, 7).Value
            op = ws.Cells(i, 3).Value
            j = j + 1
            
        Else
            ticker(1) = ws.Cells(i, 1).Value ' incrments ticker
            
            If ticker(0) <> ticker(1) Then ' compairs currention ticker to first ticker
            
                ' if block that prvents cl and op from being zero to factilitate calculations
                If cl < 1 Then
                    cl = 1
                End If
                
                If op < 1 Then
                    op = 1
                End If
                
                cl = ws.Cells(i - 1, 6) ' stores final close value
                diff = cl - op ' calcs difference
                per = diff / op ' calcs percent
                
                ws.Cells(j, 9).Value = ticker(0) ' stores ticker in excel doc
                ws.Cells(j, 10).Value = diff      ' stores difference in excel doc
                ws.Cells(j, 11).Value = per      ' stores percentange in excel doc
                ws.Cells(j, 12).Value = total    ' stores total stocks in excel doc
                ws.Cells(j, 11).Style = "Percent" ' changes percent different colomn to percent style
                
                ' conditionally formats percentage cell based on value
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                
                j = j + 1 ' incremnts counter
                
                ' resets variables for next group
                total = 0
                ticker(0) = ws.Cells(i, 1).Value
                ticker(1) = ws.Cells(i, 1).Value
                total = total + ws.Cells(i, 7)
                op = ws.Cells(i, 3).Value
                
                
            Else
                total = total + ws.Cells(i, 7).Value ' sums total
                
            End If
        End If
        Next i
        
        perange = ws.Range("k1", ws.Range("k1").End(xlDown)).Rows ' stores range of percents
        totrange = ws.Range("l1", ws.Range("l1").End(xlDown)).Rows ' stores range of total stocks
        
        ' calculates and stores the max percent, min percent, and largest volume values
        mn = Application.WorksheetFunction.Min(perange)
        mx = Application.WorksheetFunction.Max(perange)
        Large = Application.WorksheetFunction.Max(totrange)
        
        ws.Range("Q2").Value = mx ' stores the max percent change in the excel doc
        ws.Range("Q3").Value = mx ' stores the min percent change in the excel doc
        ws.Range("Q4").Value = Large ' stores the largets total volume in the excel doc
        
        ' finds the row that the max, min, and largest values
        mn_l = Application.WorksheetFunction.Match(mn, perange, 0)
        mx_l = Application.WorksheetFunction.Match(mx, perange, 0)
        Large_l = Application.WorksheetFunction.Match(Large, totrange, 0)
        
        ' uses the above information to find the corrisponding ticker value and stores it next to the information
        ws.Range("P2") = ws.Cells(mx_l, 9)
        ws.Range("P3") = ws.Cells(mn_l, 9)
        ws.Range("P4") = ws.Cells(Large_l, 9)
        
        ws.Range("Q2", "Q3").Style = "Percent" ' converts the max and min values to percents
        
Next ws
    

End Sub
