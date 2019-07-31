Attribute VB_Name = "Module1"
Sub hard()
    Dim WS_Count As Integer
    Dim stock As String
    Dim vol As Double
    Dim stkcnt As Long
    Dim first As Double
    Dim last As Double
    Dim inc As String
    Dim incv As Double
    Dim dec As String
    Dim decv As Double
    Dim gvol As String
    Dim gvolv As Double
    
    'count number of sheets
    WS_Count = ActiveWorkbook.Worksheets.Count
      
    
    'loop through all sheets
    For i = 1 To WS_Count
        vol = 0
        stkcnt = 1
        first = 0
        last = 0
        incv = 0
        decv = 0
        gvolv = 0
        
        'assign first ticker and opening price to variables
        stock = ActiveWorkbook.Worksheets(i).Cells(2, 1).Value
        first = ActiveWorkbook.Worksheets(i).Cells(2, 3).Value
        'label new columns
        ActiveWorkbook.Worksheets(i).Cells(1, 9) = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 10) = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11) = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 12) = "Total Stock Volume"
        ActiveWorkbook.Worksheets(i).Cells(2, 14) = "Greatest % Increase"
        ActiveWorkbook.Worksheets(i).Cells(3, 14) = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, 14) = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(i).Cells(1, 15) = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 16) = "Value"
        
        
        'count rows
        m = ActiveWorkbook.Worksheets(i).Cells(ActiveWorkbook.Worksheets(i).Rows.Count, "A").End(xlUp).Row
        'list first ticker
        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 9) = stock
 
        
            'loop through rows
            For j = 2 To m + 1
            

                'get closing price for current stock and add daily and % changes to list
                If ActiveWorkbook.Worksheets(i).Cells(j, 2).Value = 20161230 Then
                    last = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                    ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10) = last - first
                    If first <> 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = (last - first) / first
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11).NumberFormat = "0.00%"
                    ElseIf first = 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = "infinite"
                    End If
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j, 2).Value = 20151231 Then
                    last = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                    ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10) = last - first
                    If first <> 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = (last - first) / first
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11).NumberFormat = "0.00%"
                    ElseIf first = 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = "infinite"
                    End If
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j, 2).Value = 20141231 Then
                    last = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                    ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10) = last - first
                    If first <> 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = (last - first) / first
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11).NumberFormat = "0.00%"
                    ElseIf first = 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 11) = "infinite"
                    End If
                End If
                    
                'add volume if ticker matches current stock
                If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value = stock Then
                    vol = vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    
                'updates current stock if ticker is different adds volume total to list
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> stock Then
                    ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 12) = vol
                    ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 9) = stock
                    'reassigns current stock and opening price
                    stock = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                    first = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                    'colors change field for +/-
                    If ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10).Value > 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10).Interior.ColorIndex = 4
                    ElseIf ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10).Value < 0 Then
                        ActiveWorkbook.Worksheets(i).Cells((stkcnt + 1), 10).Interior.ColorIndex = 3
                    End If
                    'resets volume for new current stock and adds one to the ticker count
                    vol = ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    stkcnt = stkcnt + 1
                End If
                
            Next j
            
            
            
            'count rows
            q = ActiveWorkbook.Worksheets(i).Cells(ActiveWorkbook.Worksheets(i).Rows.Count, "I").End(xlUp).Row
            'get values for greatest and least % change and greatest volume
            incv = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(i).Range(ActiveWorkbook.Worksheets(i).Cells(2, 11), ActiveWorkbook.Worksheets(i).Cells(q, 11)))
            decv = Application.WorksheetFunction.Min(ActiveWorkbook.Worksheets(i).Range(ActiveWorkbook.Worksheets(i).Cells(2, 11), ActiveWorkbook.Worksheets(i).Cells(q, 11)))
            gvolv = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(i).Range(ActiveWorkbook.Worksheets(i).Cells(2, 12), ActiveWorkbook.Worksheets(i).Cells(q, 12)))
            
            For w = 2 To q
                'get tickers
                If ActiveWorkbook.Worksheets(i).Cells(w, 11).Value = decv Then
                    dec = ActiveWorkbook.Worksheets(i).Cells(w, 9).Value
                ElseIf ActiveWorkbook.Worksheets(i).Cells(w, 11).Value = incv Then
                    inc = ActiveWorkbook.Worksheets(i).Cells(w, 9).Value
                End If
                If ActiveWorkbook.Worksheets(i).Cells(w, 12).Value = gvolv Then
                    gvol = ActiveWorkbook.Worksheets(i).Cells(w, 9).Value
                End If
            Next w

            'adds the superlatives to their proper cells
            ActiveWorkbook.Worksheets(i).Cells(2, 15) = inc
            ActiveWorkbook.Worksheets(i).Cells(3, 15) = dec
            ActiveWorkbook.Worksheets(i).Cells(4, 15) = gvol
            ActiveWorkbook.Worksheets(i).Cells(2, 16) = incv
            ActiveWorkbook.Worksheets(i).Cells(2, 16).NumberFormat = "0.00%"
            ActiveWorkbook.Worksheets(i).Cells(3, 16) = decv
            ActiveWorkbook.Worksheets(i).Cells(3, 16).NumberFormat = "0.00%"
            ActiveWorkbook.Worksheets(i).Cells(4, 16) = gvolv
           
    Next i
End Sub
