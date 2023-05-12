Sub AlphabetTest():

'     |\      _,,,---,,_
'ZZZzz /,`.-'`'    -.  ;-;;,_
'     |,4-  ) )-,_. ,\ (  `'-'
'    '---''(_/--'  `-'\_)


 'setting up
Dim ws As Worksheet
ws_num = ThisWorkbook.Worksheets.Count

'MsgBox ws_num

For Each ws In ThisWorkbook.Worksheets

 ' Dim(istyfy)ing my variables
    Dim Summary_Table_Row As Long
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim i As Long
    Dim j As Integer
    Dim LastRow As Long

  ' Set the headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

    ' Set initial values
    j = 0
    Total_Stock_Volume = 0
    Yearly_Change = 0
    Summary_Table_Row = 2

  ' get the last row number --> Class Modules - Credit Card Example
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To LastRow

 ' If ticker changes then enter in the table --> Class modules - Credit Card Example
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

  ' Getting the variables --> Class modules 
     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

'-----------------------------------------------------------------------------------
'Dwight & Mr. Murderbritches credits     |\__/,|   (`\
'                                      _.|o o  |_   ) )
'                                      -(((---(((--------

            ' Handle zero total volume --> Thank you Dwight!
            If Total_Stock_Volume = 0 Then
                ' print the results
                ws.Cells(2 + j, 9).Value = Cells(i, 1).Value
                ws.Cells(2 + j, 10).Value = 0
                ws.Cells(2 + j, 11).Value = "%" & 0
                ws.Cells(2 + j, 12).Value = 0

            Else
                ' Find First non zero starting value --> Thank you Dwight!
                If ws.Cells(Summary_Table_Row, 3) = 0 Then
                    For find_value = Summary_Table_Row To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            Summary_Table_Row = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Find Change
                Yearly_Change = (ws.Cells(i, 6) - ws.Cells(Summary_Table_Row, 3))
                Percent_Change = (Yearly_Change / ws.Cells(Summary_Table_Row, 3))


                ' start of the next ticker
                Summary_Table_Row = i + 1

                ' Paste my values --> https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
                ws.Cells(2 + j, 9).Value = Cells(i, 1).Value
                ws.Cells(2 + j, 10).Value = Yearly_Change
                ws.Cells(2 + j, 10).NumberFormat = "0.00"
                ws.Cells(2 + j, 11).Value = Percent_Change
                ws.Cells(2 + j, 11).NumberFormat = "0.00%"
                ws.Cells(2 + j, 12).NumberFormat = "0"
                ws.Cells(2 + j, 12).Value = Total_Stock_Volume
                ws.Cells(2 + j, 12).NumberFormat = "#,###,##0"

            End If

  ' reset variables for new stock ticker --> I got stuck here; Dwight helped me out with the j variable
            Total_Stock_Volume = 0
            Yearly_Change = 0
            j = j + 1
            

        ' If ticker is still the same add results
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

        End If

    Next i
'-----------------------------------------------------------------------------------
' Percent change color cell formatting --> Class Module for GradeBook 
    For k = 2 To LastRow
                
        If ws.Cells(k, 10).Value > 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 4
                    
            ElseIf ws.Cells(k, 11).Value < 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
                    
            ElseIf ws.Cells(i, 11).Value = 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 0
        End If
               
    Next k

    ' Create a separate table
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
 
 
    'Find the maximum value / https://stackoverflow.com/questions/45305322/vba-code-to-find-out-maximum-value-in-a-range-and-add-1-to-the-value-in-the-othe
 
        ws.Range("P2") = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Range("P3") = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        ws.Cells(3, 16).NumberFormat = "0.00%"

    'Cell formatting https://www.automateexcel.com/vba/format-numbers/
        ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        ws.Cells(4, 16).NumberFormat = "#,###,##0"

    'Find the ticker matching the greatest/smallest values - might be ugly, but it works! Mr. Murderbritches & Dwight would disapprove...
    '/\___/\
    '\ -.- /
    '`-.^.-'
    '  /"\

            For p = 2 To LastRow
                    If ws.Cells(p, 11).Value = ws.Cells(2, 16).Value Then
                    ws.Cells(2, 15).Value = ws.Cells(p, 9).Value
                End If
            Next p

            For q = 2 To LastRow
                    If ws.Cells(q, 11).Value = ws.Cells(3, 16).Value Then
                    ws.Cells(3, 15).Value = ws.Cells(q, 9).Value
                End If
            Next q

            For l = 2 To LastRow
                    If ws.Cells(l, 12).Value = ws.Cells(4, 16).Value Then
                    ws.Cells(4, 15).Value = ws.Cells(l, 9).Value
                End If
            Next l


Next ws

End Sub


