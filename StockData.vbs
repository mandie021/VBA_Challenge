Attribute VB_Name = "Module1"

Sub StockMarket()
 
 
'Create a script that will loop through all the stocks for one year and output the following information:
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
Dim ws As Worksheet

 'declare variables
    Dim ticker As String
    Dim i As Long
    Dim oprice As Double
    Dim cprice As Double
    Dim TotalVolume As LongLong
        TotalVolume = 0
    Dim lastRow As Long
    Dim YrChange As Double
    Dim PrevAmount As Long
        PrevAmount = 2
    Dim PrecentC As Double
    Dim SummaryTableRow As Long
        SummaryTableRow = 2


For Each ws In Worksheets
ws.Activate
' Column Headers / Data Field Labels
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
        
'    'declare variables
'    Dim ticker As String
'    Dim i As Long
'    Dim oprice As Double
'    Dim cprice As Double
'    Dim TotalVolume As LongLong
'        TotalVolume = 0
'    Dim lastRow As Long
'    Dim YrChange As Double
'    Dim PrevAmount As Long
'        PrevAmount = 2
'    Dim PrecentC As Double
'    Dim SummaryTableRow As Long
'        SummaryTableRow = 2
  
    'determine last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    oprice = ws.Cells(2, 3).Value
    
    'reset SummaryRow
    SummaryTableRow = 2
    
    For i = 2 To lastRow
                        
        'check if we are in same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'ticker name
            ticker = ws.Cells(i, 1).Value
            
            'adding totalvalue
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'print ticker values to summary shhet
            ws.Range("I" & SummaryTableRow).Value = ticker
    
            ' Print The TotalVolume To The Summary Table
            ws.Range("L" & SummaryTableRow).Value = TotalVolume
           
            'set cprice
            cprice = ws.Cells(i, 6).Value
            
            'calculate YrChange
            YrChange = cprice - oprice
           
           
           'reset ticker toal
            TotalVolume = 0
            
            'create percent
            If cprice = 0 Then
                If oprice = 0 Then
                    PercentC = 0
                Else
                    PercentC = "NA"
                End If
            Else
                PercentC = FormatPercent(YrChange / cprice)
            End If
     
            'print yrchange
            ws.Cells(SummaryTableRow, 10).Value = YrChange
            
            'print PercentC
            ws.Range("K" & SummaryTableRow).Value = PercentC
            
            'set oprice
            oprice = ws.Cells(i + 1, 3)
             'conditional formatting for increase and decrease of total volume
    
            'make green
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            Else
                'make red
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
            ' Add One To The Summary Table Row
            SummaryTableRow = SummaryTableRow + 1
            
        Else
        
            'adding totalvalue
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
   
        End If
   
        
       

    Next i
    
    
Next ws



    'greatest increase/decress/+totalvalue

    'Determine last row 11- did not define before
'    lastRow = Cells(Rows.Count, 11).End(xlUp).Row
'
'    looping for the 'at a glace'
'    For j = 2 To lastRow
'
'    increase
'    If ws.Range("K" & j) > ws.Range("Q2").Value Then
'    ws.Range("Q2").Value = ws.Range("K" & j).Value
'    ws.Range("P2").Value = ws.Range("I" & i).Value
'
'    End If
'
'    decrease
'    If ws.Range("K" & j) < ws.Range("Q3").Value Then
'    ws.Range("Q3").Value = ws.Range("K" & j).Value
'    ws.Range("P3").Value = ws.Range("I" & j).Value
'
'
'    End If
'
'    volume
'    If ws.Range("L" & j) > ws.Range("Q4").Value Then
'    ws.Range("Q4").Value = ws.Range("L" & j).Value
'    ws.Range("P4").Value = ws.Range("I" & j).Value
'
'    End If
'    including the % symbol
'        ws.Range("Q2").NumberFormat = "0.00%"
'        ws.Range("Q3").NumberFormat = "0.00%"
'    Next j
    
  Columns("I:Q").AutoFit


End Sub




