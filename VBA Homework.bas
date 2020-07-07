Attribute VB_Name = "Module1"
Sub alphabetical_test()

'name variables
Dim open_amt As Variant
Dim close_amt As Variant

Dim i As Double
Dim j As Double

j = 2
i = 2

'List out the header Row
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
'Set new created ticker name equal to original ticker name. Created ticker will write over with each iteration that has ticker change
Cells(j, 9).Value = Cells(j, 1).Value

'Set the open amount of first ticker outside of loop
open_amt = Cells(i, 3).Value

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

'check if original ticker name is equal to created ticker name. If it doesn't...
If Cells(i, 1).Value = Cells(j, 9).Value Then
'They matched. I need to store volume amt each match
volume_total = volume_total + Cells(i, 7).Value
'They matched. I need close amount
close_amt = Cells(i, 6).Value

Else
'Set Year Change by using last close amount-open amt
Cells(j, 10).Value = close_amt - open_amt

    'check against division by zero
        If close_amt <= 0 Then
        Cells(j, 11).Value = 0
       
       Else
        Cells(j, 11).Value = (close_amt / open_amt) - 1
                End If
'format style percent change column
Cells(j, 11).Style = "Percent"

 If Cells(j, 10).Value >= 0 Then
 
 'format yearly change color
Cells(j, 10).Interior.ColorIndex = 4
    Else
    Cells(j, 10).Interior.ColorIndex = 3
    End If

Cells(j, 12).Value = volume_total
open_amt = Cells(i, 3).Value
volume_total = Cells(i, 7).Value

'Move to next row of created ticker column and prepare column for next iteration
j = j + 1
Cells(j, 9).Value = Cells(i, 1).Value

End If

Next i

'set new J value and year change for next ticker
Cells(j, 10).Value = close_amt - open_amt

'If statement so not dividing by zero and format cell
If close_amt <= 0 Then
Cells(j, 11).Value = 0
 Else
 Cells(j, 11).Value = (close_amt / open_amt) - 1
End If
Cells(j, 11).Style = "Percent"
If Cells(j, 10).Value >= 0 Then
Cells(j, 10).Interior.ColorIndex = 4
Else
Cells(j, 10).Interior.ColorIndex = 3
End If
Cells(j, 12).Value = volume_total
'refit all columns
Columns("I:Q").EntireColumn.AutoFit
Cells(1, 1).Select


End Sub

