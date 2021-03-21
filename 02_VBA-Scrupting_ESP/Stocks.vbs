Sub Stocks():

    For Each ws In Worksheets
        ws.Activate

        'add new cells
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        'define variables and assign values
        Dim Last_Row As Long
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

        Dim Ticker_Name As String

        Dim Ticker_Row As Integer
        Ticker_Row = 2

        Dim Row_Count As Integer
        Row_Count = 0

        Dim Ticker_Total_Vol As LongLong
        Ticker_Total_Vol = 0

        Dim Ticker_Start_Price As Double
        Ticker_Start_Price = 0

        Dim Ticker_End_Price As Double
        Ticker_End_Price = 0

        Dim CondRange As Range
        Set CondRange = Range("J2:J" & Last_Row)

        'loop through all Cells
        For i = 2 To Last_Row

            'check if ticker has changed
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                'set ticker name
                Ticker_Name = Cells(i, 1).Value
    
                'Add to ticker total volume
                Ticker_Total_Vol = Ticker_Total_Vol + Cells(i, 7).Value
    
                'print ticker name in ticker row
                Range("I" & Ticker_Row).Value = Ticker_Name

                'print ticker total volume
                Range("L" & Ticker_Row).Value = Ticker_Total_Vol

                'capture ticker start & end prices
                Ticker_Start_Price = Cells(i - Row_Count, 3)
                Ticker_End_Price = Cells(i, 6)

                If Ticker_Start_Price = 0 Then

                    'print ticker yearly change
                    Range("J" & Ticker_Row).Value = "N/A"

                    'print ticker yearly % change
                    Range("K" & Ticker_Row).Value = "N/A"

                Else

                    'print ticker yearly change
                    Range("J" & Ticker_Row).Value = Ticker_End_Price - Ticker_Start_Price

                    'print ticker yearly % change
                    Range("K" & Ticker_Row).Value = Str((Ticker_End_Price - Ticker_Start_Price) / Ticker_Start_Price * 100) + "%"

                End If

                'add 1 to the ticker row
                Ticker_Row = Ticker_Row + 1
     
                'reset ticker total volume
                Ticker_Total_Vol = 0

                'reset row count
                Row_Count = 0

                'reset ticker start & end price
                Ticker_Start_Price = 0
                Ticker_End_Price = 0

            Else

                'add to ticker total volume
                Ticker_Total_Vol = Ticker_Total_Vol + Cells(i, 7).Value

                'add to row count
                Row_Count = Row_Count + 1
     
            End If
    
        Next i

        'conditional formatting

        For Each cell In CondRange
            If cell.Value = "N/A" Then
                cell.Interior.ColorIndex = 0
                cell.Font.ColorIndex = 1
            ElseIf cell.Value > 0 Then
                cell.Interior.ColorIndex = 50
                cell.Font.ColorIndex = 51
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 22
                cell.Font.ColorIndex = 30
            End If
            
        Next

    Next 

    MsgBox ("Complete")

End Sub


