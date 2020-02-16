Attribute VB_Name = "Module1"
Sub VBAStocks()



'define everything
Dim ticker As String
Dim vol As Long
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'Set volume to 0
vol = 0

On Error Resume Next

'loop through each worksheet
For Each ws In Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


    Summary_Table_Row = 2
    

    'Determine the last row in seet
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through the columns
    
        For i = 2 To lastrow
        
            'Checks if we're within the some ticker to assign values to first and last stock price
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            year_open = ws.Cells(i, 3).Value
             
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


            'Set the ticker name
            ticker = ws.Cells(i, 1).Value
            
            'Add to the volume
            vol = vol + ws.Cells(i, 7).Value
            
            'Find year open
            year_close = ws.Cells(i, 6).Value

            'Find yearly change and percent change
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_open
            

            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the volume
            vol = 0
            
            'If the cell immediately following this row is the same ticker...
            
            Else
        
            'Add to the volume
            vol = vol + ws.Cells(i, 7).Value

            End If
            

        Next i

ws.Columns("K").NumberFormat = "0.00%"


       'format columns colors
    Dim rg As Range
    Dim x As Long
    Dim y As Long
    Dim color_cell As Range

    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    y = rg.Cells.Count

    For x = 1 To y
    Set color_cell = rg(x)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next x

Next ws


End Sub

