Attribute VB_Name = "Module4"
'I decided to build the macro for this assignment step by step base on the challenge level and the pictures that were included
'I then built a macro to run them simultaneously.

Sub MACRO_TO_RUN()

Call PART_1
Call PART_2
Call PART_3

End Sub

Sub PART_1()

'format columns and assign headings
Columns("B:S").ColumnWidth = 0
Columns("H").ColumnWidth = 3
Columns("I").ColumnWidth = 8
Columns("L").ColumnWidth = 16
Columns("M").ColumnWidth = 3
Columns("S").ColumnWidth = 3

Range("L1").Value = "Total Stock Volume"
Range("T1").Value = "[ticker row#]"

Columns("L").NumberFormat = "#,##0"

Dim i As Long, j As Integer, k As Integer


    
    
'CREATE UNIQUE LIST OF TICKER

'copy list of tickers from column A to column I, and the remove all duplicates to create unique list
Range("A:A").Copy Range("I:I")
Range("I:I").RemoveDuplicates Columns:=1

'CREATE TEMPORARY LIST OF INDEX LOCATION

'Through evaluating the data, it is understood that the data for each ticker is located in a group. They are sorted in alphabetical order (primary) and by the date (secondary).
'This makes data within the first and the last row of the same ticker can be treated as a group, which enables the use of much more efficient functions such as sum instead of sumif.
'To be able to do this, the beginning and the end rows for each ticker need to be located, using a match function.
'The result is stored in a temporary list as it will be for different functions later.

'while creating list the index location of the begining of each tickers,  the number of row in the TICKERS LIST is being counted(tick_cnt)

Dim tick_cnt As Integer
j = 1
tick_cnt = 0

Do While Cells(j, 9).Value <> ""
    j = j + 1
    tick_cnt = tick_cnt + 1
    Cells(j, 20) = Application.Match(Cells(j, 9), Range("A:A"), 0)
Loop

'Find location of the last row
'to detemine the last row of the last data group, we need the location of the last row.

Dim row_cnt As Long

    With ActiveSheet
    row_cnt = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

Cells(j, 20) = row_cnt + 1

'store the value of number of ticker
Range("I1").Value = "Ticker"
Range("U3").Value = "tick_cnt"
Range("V3").Value = tick_cnt - 1


'CALCULATE_ANNUAL_TOTAL_VOLUME
'for each of the ticker, we sum, the value in "Total Volume" column that are in the rows between b_row and e_row
'create variables to mark the begining and end row of a ticker as b_row and e_row

Dim b_row As Long, e_row As Long

    For k = 2 To tick_cnt
    
        'assign value for begining row and end row
        b_row = Cells(k, 20)
        e_row = Cells(k + 1, 20) - 1
            
        'total annual value (sum) between row b_row to e_row
        Cells(k, 12) = Application.Sum(Range(Cells(b_row, 7), Cells(e_row, 7)))

    Next k

End Sub

Sub PART_2()

'prepare column and headings
Columns("C").ColumnWidth = 8.43
Columns("F").ColumnWidth = 8.43
Columns("J:K").ColumnWidth = 8.43

Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"

Columns("J").NumberFormat = "$#,##0.00"
Columns("k").NumberFormat = "0.00%"

'create variable b_price and e_price as beging of the year and end of the year prices define them as double
Dim b_price As Double, e_price As Double
Dim i As Long, j As Long, k As Integer
Dim row_cnt As Long, tick_cnt As Integer

tick_cnt = Range("V3").Value

'for each row of tickers

    For k = 2 To tick_cnt + 1
        'extrect the value of b_price and e_price
        'i = the number of row, of b_price
        i = Cells(k, 20)
        b_price = Cells(i, 3)
        
        j = Cells(k + 1, 20) - 1
        e_price = Cells(j, 6)
        
        'count value for column yearly change
        Cells(k, 10).Value = e_price - b_price
        
        'color cell green if positive, red if negative
        If Cells(k, 10) < 0 Then
            Cells(k, 10).Interior.ColorIndex = 3
            Else
            Cells(k, 10).Interior.ColorIndex = 4
        End If
        
        'calculate percentage change, while taking caring situation when divisor is 0
        If b_price <> 0 Then
            Cells(k, 11).Value = (e_price / b_price) - 1
            Else
            Cells(k, 11).Value = 0
        End If
        
    Next k

End Sub

Sub PART_3()

'prepare column and headings
Columns("N").ColumnWidth = 21
Columns("O").ColumnWidth = 8.43
Columns("P").ColumnWidth = 15

Range("p2:p3").NumberFormat = "0.00%"
Range("p4").NumberFormat = "#,##0"

Range("o1").Value = "Ticker"
Range("p1").Value = "Value"
Range("n2").Value = "[Greatest % Increase]"
Range("n3").Value = "[Greatest % Decrease]"
Range("n4").Value = "[Greatest Total Volume]"


'FIND_MAX_MIN
'create variable max_pct, min_pct and max_vol
Dim i As Integer, max_pct As Double, tick_cnt As Integer

'Look for Max percentage
'going through the list of percentage change, find the greatest % increase
'initiate max_pct as 0 and max iteration as the number of tickers (tick_cnt)

tick_cnt = Range("v3").Value
max_pct = 0
    For i = 2 To tick_cnt + 1
        If Cells(i, 11).Value > max_pct Then
            max_pct = Cells(i, 11).Value
            Range("o2").Value = Cells(i, 9)
            Range("p2").Value = max_pct
        End If
    Next i


'Look for Min Percentage, then find the matching tickers
Range("p3").Value = Application.Min(Range("k:k"))
Range("o3") = Cells(Application.Match(Range("p3").Value, Range("k:k"), 0), 9)


'Look for Max Volume, then find the matching tickers
Range("p4").Value = Application.Max(Range("L:L"))
Range("o4") = Cells(Application.Match(Range("p4").Value, Range("L:L"), 0), 9)


'delete temporary columns
Columns("A:E").ColumnWidth = 8.43
Columns("Q:Z").Delete


End Sub
