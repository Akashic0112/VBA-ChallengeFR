'1. Create a script that will loop through all the stocks for one year for each run and take the following information.
'       The ticker symbol.
'       Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The total stock volume of the stock.
'2. You should also have conditional formatting that will highlight positive change in green and negative change in red

Sub Stocks()
    'Defne Dims
    'Dim wb As Workbook
    'Set wb = Test_run- Important?
   
Dim ws As Worksheet
Dim j As Integer
Dim Total_Volume As Long
Total_Volume = 0

Dim start As Integer

    'summary table column titles
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

rc = Cells(Rows.Count, 1).End(xlUp).Row
j = 2
   
'For Each ws In Worksheets
    For i = 2 To rc:
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          Total_Volume = Total_Volume + Cells(i, 7).Value
          Cells(j, 9).Value = Cells(i, 1).Value
          Cells(j, 12).Value = Total_Volume
          Total_Volume = 0
          j = j + 1
        Else
          Total_Volume = Total_Volume + Cells(i, 12).Value
        End If
           
                
    
    Next i
'Next ws
    
End Sub
