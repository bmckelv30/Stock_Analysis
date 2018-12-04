VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

Sub stock()

    For Each ws In Worksheets
    
    ' Determine the Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add the Headers
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Total Stock Volume"
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
           
    ' Set totalvol to zero
    Dim totalvol As Double
    totalvol = 0
    
    Dim Ticker As String
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the Ticker
            Ticker = ws.Cells(i, 1).Value
    
            ' Add to the totalvol
            totalvol = totalvol + ws.Cells(i, "G").Value
    
            ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Print the Total to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = totalvol
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Brand Total
            totalvol = 0
    
        ' If the cell immediately following a row is the same Ticker...
        Else
        
          ' Add to the totalvol
          totalvol = totalvol + Cells(i, "G").Value
        End If
    Next i

    Next ws

End Sub

'Yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.

Sub yearlychg()

    For Each ws In Worksheets
    
    ' Determine the Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add new columns (if running after running stock module)
    ws.Range("J:J").EntireColumn.Insert
    ws.Range("K:K").EntireColumn.Insert
    
    'Add headers
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
           
    ' Set totals to zero
    Dim totalchg As Double
    Dim totop As Double
    Dim totcl As Double
    
    totalchg = 0
    totop = 0
    totcl = 0
    
    ' Set pct to zero
    Dim pct As Double
    pct = 0
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Add to the open total
            totop = totop + (ws.Cells(i, "C").Value)
        
            ' Add to the close total
            totcl = totcl + (ws.Cells(i, "F").Value)
            
            ' Add to the Total
            totalchg = totop - totcl
            
            'Update pct
            pct = totalchg / totop
            
            ' Print the Total to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = totalchg
            ws.Range("J:J").NumberFormat = "#.################"
            ' Add colors
                If (totalchg >= 0) Then
                     ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
                ' Otherwise color it red
                Else
                    ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
                End If
            ' Print the Percent change to the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = pct
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Totals
            totalchg = 0
            totop = 0
            totcl = 0
            
            ' Reset pct
            pct = 0
        ' If the cell immediately following a row is the same Ticker...
        Else
          ' Add to the open total
          totop = totop + (ws.Cells(i, "C").Value)
          ' Add to the close total
          totcl = totcl + (ws.Cells(i, "F").Value)
        End If
    Next i
    Next ws

End Sub


