Luisa Challenge Module 2
Here is my code-solution for the Mod 2 Challenge. It is still incomplete, but
I gave my max to get close to what was asked.

Sub Challenge2()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'set a initial variable for the type of ticker
Dim Ticker As String
Dim First_Open As Double
Dim Last_Close As Double
Dim Change As Double
Dim Porcentage As Double

'set a initial variable for holding the total of volume per each ticker
Dim Total_Volume As Double
Total_Volume = 0

'keep track of the location for each ticker
Dim Summary_table_row As Integer
Summary_table_row = 2

'Creating the titles for Ticker and Total_Volume columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Loop through all ticker types
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
First_Open = Cells(2, 3).Value

For i = 2 To LastRow


'Check if we are still within the same ticker type
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Last_Close = Cells(i, 6).Value
Change = Last_Close - First_Open

'Calculating the %
Porcentage = Change / First_Open

'set the ticker type
Ticker = Cells(i, 1).Value

'add to the volume
Total_Volume = Total_Volume + Cells(i, 7).Value

'print the ticker in the summary table
Range("I" & Summary_table_row).Value = Ticker

'print the volume to the summary table
Range("L" & Summary_table_row).Value = Total_Volume

Range("J" & Summary_table_row).Value = Change

Range("K" & Summary_table_row).Value = Porcentage


'add one to the summary table row
Summary_table_row = Summary_table_row + 1

'Reset the Volume
Total_Volume = 0
First_Open = Cells(i + 1, 3).Value
'If the cell immediately following a row is the same ticker...
Else

'add to the volume
Total_Volume = Total_Volume + Cells(i, 7).Value


     End If
'adding color green for positive values and color red for negative values

If Cells(i, 10) > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
ElseIf Cells(i, 10) < 0 Then
Cells(i, 10).Interior.ColorIndex = 3
Else
Cells(i, 10).Interior.ColorIndex = 0
     
    End If
    
   Next i
   
    Next ws
 
End Sub




