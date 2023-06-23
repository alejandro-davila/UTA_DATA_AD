# VBA-challenge
Assignment 2 - VBA - Stock Data - ADavila

The VBA code utilitzed in this project includes snippets and ideas sourced from various contributors and public resources. I would like to acknowledge the following individuals and organizations for their valuable code contributions:

UTA_Bootcamp

Microsoft - learn.microsoft.com

StackOverflow - stackoverflow.com

SuperExcelVBA - superexcelvba.com

Automate Excel - automateexcel.com

Their knowledge have greatly assisted in the development of the code below:

'VBA-challenge_Assignment2_A.Davila


Sub Module2Challenge_AD()

For Each ws In Worksheets
ws.Activate

            'Section_A - used to create all of my conditions 
    Dim worksheetname As String
    Dim TickerName As String
    Dim YearlyChange As Double
    Dim PercentChange As Single
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim i As Long
    Dim J As Long    
    Dim n As Long
    Dim Y As Long
    Dim SolutionTable As Long
    Dim themax As Double
    Dim themin As Double
    Dim maxvol As Double
    Dim ticker2 As String    
    Dim SummaryTable As Integer
        SummaryTable = 2
    Dim VolumeofStock As Double
        VolumeofStock = 0
    Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim Header As Integer
        Header = 9
            'end of Section_A


            'Section_B - This section will create headers and format them in their assigned cell location.
        Cells(1, Header).Value = "Ticker"
        Cells(1, Header + 1).Value = "Yearly Change"
        Cells(1, Header + 2).Value = "Percent Change"
        Cells(1, Header + 3).Value = "Total Stock Volume"
        Cells(1, Header + 6).Value = "Ticker"
        Cells(1, Header + 7).Value = "Value"
        Cells(2, Header + 5).Value = "Greatest % Increase"
        Cells(2, Header + 5).Font.Bold = True
        Cells(2, Header + 5).HorizontalAlignment = xlHAlignCenter
        Cells(3, Header + 5).Value = "Greatest % Decrease"
        Cells(3, Header + 5).Font.Bold = True
        Cells(3, Header + 5).HorizontalAlignment = xlHAlignCenter
        Cells(4, Header + 5).Value = "Greatest Total Volume"
        Cells(4, Header + 5).Font.Bold = True
        Cells(4, Header + 5).HorizontalAlignment = xlHAlignCenter
            'end of Section_B


            'Section_C - This section finds the Ticker, Openning, and Closing data
        For i = 2 To LastRow
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                TickerName = Cells(i, 1).Value
                    Range("I" & SummaryTable).Value = TickerName
                OpenPrice = Cells(i, 3).Value
                        'To view the OpenPrice amounts in column R use - Range("R" & SummaryTable).Value = OpenPrice
            
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ClosePrice = Cells(i, 6).Value
                        'To view the ClosePrice amounts in column S use - Range("S" & SummaryTable).Value = ClosePrice
                
                YearlyChange = ClosePrice - OpenPrice
                    Range("J" & SummaryTable).Value = YearlyChange
                PercentChange = YearlyChange / OpenPrice
                    Range("K" & SummaryTable).Value = PercentChange
                SummaryTable = SummaryTable + 1
                VolumeofStock = 0
            Else
            End If
        Next i
            'Section_C ends.           
                
            'Section_D - This section helps find the total Volume without being affected by Section_C. 
        SummaryTable = 2
        
        For J = 2 To LastRow
            If Cells(J + 1, 1).Value <> Cells(J, 1).Value Then
                VolumeofStock = VolumeofStock + Cells(J, 7).Value
                    Range("L" & SummaryTable).Value = VolumeofStock
                SummaryTable = SummaryTable + 1
                VolumeofStock = 0
            Else
                VolumeofStock = VolumeofStock + Cells(J, 7).Value
            End If
        Next J
            'Section_D ends.    

            'Section_E - This section deals with the formating of the solution table.
        SolutionTable = Cells(Rows.Count, "I").End(xlUp).Row
        Range("K2:K" & SolutionTable).NumberFormat = "0.00%"
    
        For n = 2 To SolutionTable
            If Cells(n, 10).Value <= 0 Then
                Cells(n, 10).Interior.ColorIndex = 3
            Else
                Cells(n, 10).Interior.ColorIndex = 4
            End If
        Next n
            'Section_E ends.
           
            'Section_F - This section answers and formats the second solution table.
        themax = WorksheetFunction.Max(Range("K2:K" & SolutionTable))
            Range("P2") = themax
        themin = WorksheetFunction.Min(Range("K2:K" & SolutionTable))
            Range("P3") = themin
        maxvol = WorksheetFunction.Max(Range("L2:L" & SolutionTable))
            Range("P4") = maxvol
            Range("P2:P3").NumberFormat = "0.00%"
        For Y = 2 To LastRow
            If themax = Cells(Y, 11).Value Then
                ticker2 = Cells(Y, 9).Value
                    Range("o2").Value = ticker2
            ElseIf themin = Cells(Y, 11).Value Then
                ticker2 = Cells(Y, 9).Value
                    Range("o3").Value = ticker2
        
            ElseIf maxvol = Cells(Y, 12).Value Then
                ticker2 = Cells(Y, 9).Value
                    Range("o4").Value = ticker2
            End If
        Next Y
        
        For HeaderB = 9 To 24
            Cells("1", HeaderB).Font.Bold = True
            Cells("1", HeaderB).HorizontalAlignment = xlHAlignCenter
            Columns(HeaderB).AutoFit
        Next HeaderB
            'Section_F ends.


 
Next ws
 

End Sub
