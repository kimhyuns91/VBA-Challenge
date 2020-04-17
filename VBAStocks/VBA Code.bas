Attribute VB_Name = "Module2"
Sub VBAStock():

For Each ws In Worksheets

'Dimension the variables
Dim ticker As String
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_changed As Double
Dim percent_changed As Double
Dim total_volume As LongLong
Dim counter As Long

counter = 2
total_volume = 0

'Capture the initial yearly open value

yearly_open = ws.Range("C2").Value

'Find the last row/column of the worksheet

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1

'Loop through each of the row

For i = 2 To lastrow + 1
    
    'Create an if statement to find when the ticker changes
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Grab the ticker
        
            ticker = ws.Cells(i, 1).Value
        
        'Calculate the Yearly Change
                            
            yearly_closed = ws.Cells(i, 6).Value
            yearly_changed = yearly_closed - yearly_open
             
        'Calculate the percent change
            
             If yearly_open = 0 And yearly_closed > 0 Then
            
                    t = i
                
                While ws.Cells(t, 3).Value <> 0
                
                    yearly_open = ws.Cells(t, 3).Value
                    
                    t = t - 1
                
                Wend
                
             'replace  yeary change for when stock became available
                yearly_changed = yearly_closed - yearly_open
                
                percent_changed = yearly_changed / yearly_open
                
    
                
            ElseIf yearly_open = 0 And yearly_closed = 0 Then
            
                percent_changed = 0
            
            Else
            
                percent_changed = yearly_changed / yearly_open
            
            End If
            
            yearly_open = ws.Cells(i + 1, 3).Value
        
        'Add the total  Stock Volume
        
            total_volume = total_volume + ws.Cells(i, 7)
            
        'Display output
        
        ws.Cells(counter, 9).Value = ticker
        ws.Cells(counter, 10).Value = yearly_changed
        ws.Cells(counter, 11).Value = percent_changed
        ws.Cells(counter, 11).NumberFormat = "0.00%"
        ws.Cells(counter, 12).Value = total_volume
        counter = counter + 1
        total_volume = 0

         Else
        
            total_volume = total_volume + ws.Cells(i, 7)
    
    
    End If
    
Next i

' label the header

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Precent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



'Conditional formatting

new_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row - 1

For j = 2 To new_lastrow + 1

'Color green if positive and red if negative yearly change

    If ws.Cells(j, 10).Value > 0 Then
        
        ws.Cells(j, 10).Interior.ColorIndex = 4
        
    Else
        
        ws.Cells(j, 10).Interior.ColorIndex = 3
        
    End If
    
Next j
        
        
        
        

'Find the greatest % increase from calcualted values
 
 Dim p_max As Double
 Dim p_min As Double
 Dim tot_vol As LongLong
 Dim ticker1 As String
 Dim ticker2 As String
 Dim ticker3 As String
 
 
 p_max = 0
 p_min = 0
 tot_vol = 0
 
new_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
new_column = ws.Cells(1, Columns.Count).End(xlToLeft).Column
 
 For k = 2 To new_row

    If ws.Cells(k, 11).Value > p_max Then
    
    p_max = ws.Cells(k, 11).Value
    ticker1 = ws.Cells(k, 9).Value
    
    
 'Find the greatest % decrease from calcualted values
    
    ElseIf ws.Cells(k, 11).Value < p_min Then
    
    p_min = ws.Cells(k, 11).Value
    ticker2 = ws.Cells(k, 9).Value
    
    End If

'Find the greatest total volume from calcualted values
    If ws.Cells(k, 12).Value > tot_vol Then
    
    tot_vol = ws.Cells(k, 12).Value
    ticker3 = ws.Cells(k, 9).Value
    
   End If
    
  
  Next k
 
' display output

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 16).Value = ticker1
ws.Cells(2, 17).Value = p_max
ws.Cells(3, 16).Value = ticker2
ws.Cells(3, 17).Value = p_min
ws.Cells(4, 16).Value = ticker3
ws.Cells(4, 17).Value = tot_vol


ws.Range("Q2:Q3").NumberFormat = "0.00%"


ws.Columns("A:W").AutoFit


Next ws

End Sub
