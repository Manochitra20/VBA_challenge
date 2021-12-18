Attribute VB_Name = "Module1"
Sub MultiYear_stock():

'Defining variables
Dim ws As Worksheet
Dim Ticker As String
Dim Total_volume As Long
Dim open_price As Double
Dim close_price As Double
Dim Yearly_change As Double
Dim Percent_change As Double
Dim Summary As Long

    ' Keeping track of the location for each stock in the summary row
        Summary = 2
      'starting Totalvolume with 0 and adding to it
        Total_volume = 0

    'Looping through each sheet in the workbook
    For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_change"
    ws.Range("K1").Value = "Percent_change"
    ws.Range("L1").Value = "Total_volume"
    
    'defining the initial open_price
    open_price = ws.Cells(2, 3).Value
    
    
    'Getting LastRow of the worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Looping thorugh all stocks for one year
    For t = 2 To LastRow
     
    'Check if we are still within the same year, if it is not...
        If ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
    'set ticker symbol
        Ticker = ws.Cells(t, 1).Value
    'add to totalvolume
        Total_volume = Total_volume + ws.Cells(t, 7).Value
        
    'print Ticker symbol in summary row
        ws.Range("I" & Summary).Value = Ticker
    'print total volume in summary row
        ws.Range("L" & Summary).Value = Total_volume
      
    
    'defining the close_price
    close_price = ws.Cells(t, 6).Value
    
    'find yearly_change
    Yearly_change = (close_price - open_price)
    'print yearly_change in summary row
    ws.Range("J" & Summary).Value = Yearly_change
    
    'find percent_change
    If open_price = 0 Then
        Percent_change = 0
    Else

    Percent_change = Yearly_change / open_price
    End If
    
    'print percent change in summary row
    ws.Range("K" & Summary).Value = Percent_change
    
     'conditional formatting the yearly-changeonditional formatting
        'that will highlight positive change in green and negative change in red
        If ws.Range("J" & Summary).Value > 0 Then
            ws.Range("J" & Summary).Interior.ColorIndex = 4
        Else
            ws.Range("J" & Summary).Interior.ColorIndex = 3
        End If
     'changing number format for percent change
     ws.Range("K" & Summary).NumberFormat = "0.00%"
    
        
     'Increment summary row to move to next row without overwriting
      Summary = Summary + 1
        
       'reset Totalvolume
       Total_volume = 0
        
        'reset open_price
        open_price = ws.Cells(t + 1, 3).Value
        
     
   End If
   
    Next t
    
    'reset summary for next worksheet
    Summary = 2
    
    Next ws
    

End Sub

