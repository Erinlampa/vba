Sub tickerCounter()
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

'1. Set up loop to loop through all worksheets in workbook
'2. Set the variables
'3. Print the data


For Each ws In Worksheets

'Set variable for holding ticker name
     Dim ticker As String

'Set variable for volume total
     Dim volume As String
     volume = 0
    
'Set variable to hold ticker symbols
     Dim ticker_tracker_row As Integer
     ticker_tracker_row = 2
    

' Determine the Last Row
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     Dim Yearly_Change As Double

     Dim Year_Open As Double
    
'Assign the starting point for the year open stock value.
    Year_Open = ws.Cells(2, 3)

    Dim Year_Close As Double
    
    Dim Percent_Change As Double

       
'-------Begin Loop---------------
'loop through ticker and add ticker symbol

    For I = 2 To lastrow
    

    ' Check if within the same ticker symbol, if not

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        ticker = Cells(I, 1).Value
        
'Establish the stock value at year close for that ticker type
        Year_Close = Cells(I, 6).Value

'Determine the yearly change
        Yearly_Change = (Year_Close) - (Year_Open)
        
'Percent Change-------------------------------------------

 'If the year close value equals the year open value then set percent change to zero.
        If Year_Close = Year_Open Then
        Percent_Change = 0

'If the year open value is zero then set the percent change to 100%.

            ElseIf Year_Open = 0 Then
            Percent_Change = 1

'Otherwise, calculate the percent change by dividing the previously calculated yearly change by the
'year open value.
            Else

        Percent_Change = (Yearly_Change) / (Year_Open)

 End If
 
      
'Reset Yearly_Open
    Year_Open = ws.Cells(I + 1, 3)
    
'Add to the volume total
        volume = volume + Cells(I, 7).Value
        
'Print volume to ticker_tracker
        ws.Range("J" & ticker_tracker_row).Value = volume

'Print ticker symbol to ticker_tracker
        ws.Range("I" & ticker_tracker_row).Value = ticker
        
'Print yearly change
        ws.Range("K" & ticker_tracker_row).Value = Yearly_Change
        
'Print percent change
        ws.Range("L" & ticker_tracker_row).Value = Percent_Change
        

'Add one to ticker_tracker_row
        ticker_tracker_row = ticker_tracker_row + 1
        
       
' If next row is the same ticker symbol
             Else


' Add to the volume Total
        volume = volume + ws.Cells(I, 7).Value
                      
        End If

'---End Loop--------------------------------------------

  Next I
  
'Assign Final Results--------------

Dim max As Double
Dim min As Double
Dim vol As Double
Dim maxticker As String
Dim minticker As String
Dim volticker As String

'Calculate Final Results-------------

max = Application.WorksheetFunction.max(ws.Columns("L"))
min = Application.WorksheetFunction.min(ws.Columns("L"))
vol = Application.WorksheetFunction.max(ws.Columns("J"))


'Print data-------------------------
ws.Cells(2, 16).Value = max

ws.Cells(3, 16).Value = min

ws.Cells(4, 16).Value = vol


Next ws

End Sub
