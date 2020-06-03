Attribute VB_Name = "Module1"
Sub stockmarket()

    'Setting Dimensions
    Dim total As Double
    Dim change As Double
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim i As Long
    Dim j As Integer
    
    'Setting Headers
     Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"


    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Variable Values
    j = 0
    total = 0
    change = 0
    start = 2
    
     For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    total = total + Cells(i, 7).Value

        
    If total = 0 Then
               
    Range("I" & 2 + j).Value = Cells(i, 1).Value
    Range("J" & 2 + j).Value = 0
    Range("K" & 2 + j).Value = "%" & 0
    Range("L" & 2 + j).Value = 0

    Else
          
    If Cells(start, 3) = 0 Then
    For find_value = start To i
    If Cells(find_value, 3).Value <> 0 Then
    start = find_value
    Exit For
    End If
    Next find_value
    End If

    ' Percent Change
    change = (Cells(i, 6) - Cells(start, 3))
    percentChange = Round((change / Cells(start, 3) * 100), 2)

    ' new stock tickers
    start = i + 1

    ' input results
    Range("I" & 2 + j).Value = Cells(i, 1).Value
    Range("J" & 2 + j).Value = Round(change, 2)
    Range("K" & 2 + j).Value = "%" & percentChange
    Range("L" & 2 + j).Value = total

    'creating colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell.Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell.Interior.Color = vbRed
            End With
       End Select
    Next g

            End If

            'Have to reset the variables for every different stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

 
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i
    
 Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

   

End Sub



