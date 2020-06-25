Attribute VB_Name = "Module1"
Sub vba_challenge()

'Assignment of variables

Dim ws As Worksheet
Dim ticker As String
Dim stockvol As Double
Dim sumtable As Integer
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Direction to pull through all worksheets in file
'Naming sumtable column headers

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change ($)"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Initial formatting
    
    ws.Range("K2:K3050").NumberFormat = "0.00%"
    
'Initial value setting

    sumtable = 2
    stockvol = 0
    open_price = ws.Cells(2, 3).Value
    
'Direction to loop through data provided

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'If statement and logic to pull data to sumtable
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            stockvol = stockvol + ws.Cells(i, 7).Value
            ws.Range("I" & sumtable).Value = ticker
            ws.Range("L" & sumtable).Value = stockvol
            close_price = ws.Cells(i, 6).Value
            yearly_change = close_price - open_price
            ws.Range("J" & sumtable).Value = yearly_change
            percent_change = yearly_change / open_price
            On Error Resume Next
            ws.Range("K" & sumtable).Value = percent_change
            open_price = ws.Cells(i + 1, 3).Value
            stockvol = 0
            sumtable = sumtable + 1
        
        Else
            stockvol = ws.Cells(i, 7).Value + stockvol
        
        End If
    
'Formatting
        If ws.Cells(i, 10) > 0 Then
            ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
        
        Else
            ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        
        End If
        
   Next i
   
   ws.Range("J" & sumtable).ClearContents
   
Next
    
End Sub
