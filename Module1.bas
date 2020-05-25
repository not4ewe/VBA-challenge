Attribute VB_Name = "Module1"
Sub Stock_Market()

'Create a script that will loop through all the stocks for one year and output the following information.
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.


'set variable for holding total volume for each ticker symbol in the summary table
Dim Ticker_Total As Double
Ticker_Total = 0

'Keep track of ticker symbol output
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Lablel columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

j = 2
i = 2

Cells(i, 9).Value = Cells(i, 1).Value

DateMinOpen = Cells(i, 3).Value

'find the last row
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through one year
For i = 2 To LastRow
          
    'check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set Ticker name
        Ticker_Name = Cells(i, 1).Value
     
        'Add to ticker total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
        'Output the ticker symbol to the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Name
    
        'output the ticker total to summary table
        Range("L" & Summary_Table_Row).Value = Ticker_Total
           
        'add one row
        Summary_Table_Row = Summary_Table_Row + 1
 
        Ticker_Total = 0
 
    'if the cell following a row is the same ticker
    Else

        'Add to ticker total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    End If
'--------------
'Yearly change
'--------------
    If Cells(i, 1).Value = Cells(j, 9).Value Then

        DateMaxClose = Cells(i, 6).Value

    Else

    'calculated fields
        Cells(j, 10).Value = DateMaxClose - DateMinOpen

    If DateMaxClose <= 0 Then

        Cells(j, 11).Value = 0

    Else
        Cells(j, 11).Value = (DateMaxClose / DateMinOpen) - 1

    End If

    'Format cell as percent
    Cells(j, 11).Style = "Percent"
   
    'conditional color format cells
    If Cells(j, 10).Value >= 0 Then

        Cells(j, 10).Interior.ColorIndex = 4

    Else

        Cells(j, 10).Interior.ColorIndex = 3

    End If

'reset variables'

DateMinOpen = Cells(i, 3).Value


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
j = j + 1
Cells(j, 9).Value = Cells(i, 1).Value

End If
'------------------------------------------------------------------
      
  Next i

End Sub

