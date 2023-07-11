Public Class Class1

End Class
Sub stockanalysis()


    ' Set CurrentWs as a worksheet object variable.

    Dim CurrentWs As Worksheet

    Dim Need_Summary_Table_Header As Boolean

    Dim COMMAND_SPREADSHEET As Boolean

    Need_Summary_Table_Header = False       'Set Header flag
    COMMAND_SPREADSHEET = True              'Hard part flag

    ' Loop through all of the worksheets in the active workbook.

    For Each CurrentWs In Worksheets

        ' Set an initial variables for tickername as text
        Dim tickername As String
        tickername = " "

        ' Set an initial variable for tickervolume, yearlychange, percentchange

        Dim yearlychange As Double
        yearlychange = 0

        Dim percentchange As Double
        percentchange = 0

        Dim tickervolume As Double
        tickervolume = 0

        ' Set an initial variable for open and close price of stock

        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0

        ' Set an initial variable for Best, Worst Yearly Change and Total Volume

        Dim MAX_TICKERNAME As String
        MAX_TICKERNAME = " "

        Dim MIN_TICKERNAME As String
        MIN_TICKERNAME = " "

        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0

        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0

        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0

        Dim MAX_VOLUMETICKER As String
        MAX_VOLUMETICKER = " "


        ' Set an summary table for data

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Counts the number of rows
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each row
        ' Set initial value of Open Price
        Open_Price = CurrentWs.Cells(2, 3).Value

        ' Loop through each row
        For i = 2 To Lastrow

            ' Check if we are still within the same ticker, if it is not...
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then

                ' Close Price Loop
                Close_Price = CurrentWs.Cells(i, 6).Value
                yearlychange = Close_Price - Open_Price

                If Open_Price <> 0 Then
                    percentchange = (yearlychange / Open_Price) * 100
                Else
                    ' Unlikely, but it needs to be checked to avoid program crushing
                    ' MsgBox( "For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet." )
                End If

                ' Set the tickername
                tickername = CurrentWs.Cells(i, 1).Value


                ' Add to the tickervolume
                tickervolume = tickervolume + CurrentWs.Cells(i, 7).Value

                ' Print the tickername in the Summary Table
                CurrentWs.Range("J" & Summary_Table_Row).Value = tickername

                ' Print the yearlychange in the Summary Table
                CurrentWs.Range("K" & Summary_Table_Row).Value = yearlychange
                If (yearlychange > 0) Then
                    CurrentWs.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (yearlychange <= 0) Then
                    CurrentWs.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

                ' Print the percentagechange in the Summary Table
                CurrentWs.Range("L" & Summary_Table_Row).Value = (CStr(percentchange) & "%")

                ' Print the tickervolume in the Summary Table
                CurrentWs.Range("M" & Summary_Table_Row).Value = tickervolume

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset the tickervolume, yearlychange, percentchange, Close_Price, Open_Price

                percentchange = 0
                Close_Price = 0
                Open_Price = CurrentWs.Cells(i + 1, 3).Value

                ' If then for Best Percent Change, Worst Percent Change

                If (yearlychange > MAX_PERCENT) Then
                    MAX_PERCENT = yearlychange
                    MAX_TICKERNAME = tickername

                ElseIf (yearlychange < MIN_PERCENT) Then
                    MIN_PERCENT = yearlychange
                    MIN_TICKERNAME = tickername

                End If

                If (tickervolume > MAX_VOLUME) Then
                    MAX_VOLUME = tickervolume
                    MAX_VOLUMETICKER = tickername

                End If

                ' Hard part adjustments to resetting counters

                yearlychange = 0
                tickervolume = 0

                ' If the cell immediately following a row is the ticker...
            Else

                ' Add to the tickervolume

                tickervolume = tickervolume + CurrentWs.Cells(i, 7).Value

            End If

        Next i


        ' Analysis of Max Percent Change, Min Percent Change, Max Tickername, Min Tickername, Max Volume, Min Volume

        CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
        CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
        CurrentWs.Range("P2").Value = MAX_TICKERNAME
        CurrentWs.Range("P3").Value = MIN_TICKERNAME
        CurrentWs.Range("Q4").Value = MAX_VOLUME
        CurrentWs.Range("P4").Value = MAX_VOLUMETICKER

    Next CurrentWs


End Sub