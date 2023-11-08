Sub Ticker()

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent As Double
Dim Volume As LongLong
Dim Opening As Double
Dim openingCell As Long
Dim Closing As Double
Dim LastRow As Long
Dim Summary_Table_Row As Integer
Dim Greatest_In As Double
Dim Greatest_D As Double
Dim GTV As LongLong
Dim GI_Ticker As String
Dim GD_Ticker As String
Dim GTV_Ticker As String

    


For Each ws In Worksheets
	  'autofit columns 
        ws.Cells.Columns.AutoFit
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Yearly_Change = 0
    Percent = 0
    openingCell = 1
    Summary_Table_Row = 2
    Greatest_D = 0
    Greatest_In = 0
    GTV = 0

    'name headers for new tables
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % of Increase"
    ws.Cells(3, 15).Value = "Greatest % of Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    	  'for loop for summary table
        For i = 2 To LastRow
        
        
		'loop down rows to find ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
			'add up volume values together for same tickers
                Volume = Volume + ws.Cells(i, 7).Value
            
			'find opening value for ticker
                Opening = ws.Cells(openingCell, 3).Value
			'find closing value for ticker
                Closing = ws.Cells(i, 6).Value
			'calculate yearly change 
                Yearly_Change = Closing - Opening
            
				'calculate percent change
                    If Opening > 0 Then
                        Percent = (Yearly_Change / Opening) * 100
                    End If
            
			'input ticker names, yearly change, percent change and total stock volume to summary table
                ws.Cells(Summary_Table_Row, 9).Value = Ticker
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Round(Percent, 2) & "%"
                ws.Cells(Summary_Table_Row, 12).Value = Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Volume = 0
                openingCell = 1
    
        
            Else
            
            	'add value of current cell to total running volume value
                Volume = Volume + ws.Cells(i, 7).Value
                    If openingCell = 1 Then
                        openingCell = i
                    End If
                
            End If
        
        Next i
        
        For j = 2 To 3001
        
		'loop through percent change to find greatest increase value
            If ws.Range("K" & j).Value > Greatest_In Then
            Greatest_In = ws.Range("K" & j).Value

		'find ticker for greatest increase value
            Greatest_In_Ticker = ws.Range("I" & j).Value
            
            End If
            
		'loop through percent change to find greatest decrease value
            If ws.Range("K" & j).Value < Greatest_D Then
            Greatest_D = ws.Range("K" & j).Value

		'find ticker for greatest decrease value
            Greatest_D_Ticker = ws.Range("I" & j).Value
            
            End If
            
		'loop through volume to find greatest volume value
            If ws.Range("L" & j).Value > GTV Then
            GTV = ws.Range("L" & j).Value
	
		'find ticker for greatest volume value
            GTV_Ticker = ws.Range("I" & j).Value
            
            End If
            
        Next j
         
		'input greatest increase %, greatest decrease % and toral values into table
            ws.Cells(2, 17).Value = Round(Greatest_In, 2) & "%"
            ws.Cells(2, 16).Value = Greatest_In_Ticker
            ws.Cells(3, 17).Value = Round(Greatest_D, 2) & "%"
            ws.Cells(3, 16).Value = Greatest_D_Ticker
            ws.Cells(4, 17).Value = GTV
            ws.Cells(4, 16).Value = GTV_Ticker
                 
        
        'for loop for green or red cell color in J column
        For k = 2 To 3001
            
		'change bg color of cell to green if positive
            If ws.Range("J" & k).Value >= 0 Then
            ws.Range("J" & k).Interior.ColorIndex = 4
            
            Else
            
		'change bg color to red if negative
            ws.Range("J" & k).Interior.ColorIndex = 3
            
            End If
        
        Next k
        
Next ws
        
End Sub



