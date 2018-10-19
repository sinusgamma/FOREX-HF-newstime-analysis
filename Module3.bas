Attribute VB_Name = "Module3"
'Module level variables

    Dim LastRow_D As Integer ' on Derived Sheet
    Dim LastColumnNbr_D As Integer ' on Derived Sheet
    Dim LastColumnLtr_D As String ' on Derived Sheet
    Dim GroupFirst As Integer 'first column of asset
    Dim GroupFirstLtr As String
    Dim GroupLastLtr As String
    Dim NewSheetName As String
    Dim AllBinNbr As Integer
    

Sub a4_GenerateCharts()

    'Use if chart doesn't have data, blank cells
    'On Error Resume Next
    
    Dim sh As Worksheet
    Dim SheetNameEnd As String
    Dim i As Integer
        
    'Sheet limits on Derived sheet
    LastRow_D = Sheets("Derived").UsedRange.Rows.Count 'Number of Rows in Derived sheet
    LastColumnNbr_D = Sheets("Derived").UsedRange.Columns.Count 'Number of columns in Derived sheet
    LastColumnLtr_D = ConvertToLetter(LastColumnNbr_D)
    
    'Go through all the assets and asset_CR
    For GroupFirst = 14 To LastColumnNbr_D Step 8 'the first column of the asset
        
        GroupFirstLtr = ConvertToLetter(GroupFirst)
        GroupLastLtr = ConvertToLetter(GroupFirst + 7)
    
        'Get sheet names
        NewSheetName = Left(Sheets("Derived").Range(GroupFirstLtr & "1").Value, 6) 'Get first 6 asset character for Name and create sheet
        SheetNameEnd = Right(Sheets("Derived").Range(GroupFirstLtr & "1").Value, 3)
        If SheetNameEnd = "_CR" Then
            NewSheetName = NewSheetName & SheetNameEnd
        End If
        
        'Make sheets
        If SheetExists(NewSheetName) = False Then
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = NewSheetName
        End If
        Set sh = ActiveWorkbook.Worksheets(NewSheetName)
                
        'Populate the sheets
        Call AllTimeframeScatter(sh)
        Call PartialTimeframeScatter(sh)
        Call BoxPlot(sh, "Act-For < 0")  ' the Act-For tag is checked for formating, not only for title
        Call BoxPlot(sh, "Act-For > 0")
        Call DoubleChartData(sh)
        Call DoubleChartDraw(sh)
        
    Next GroupFirst
    
End Sub

Sub AllTimeframeScatter(shLocal As Worksheet)
    'ALL TIMEFRAME SCATTER PLOTS
    
    Dim i As Integer ' counter for news columns
    Dim k As Integer ' counter for same asset timeframe columns
    Dim iLtr As String
    Dim kLtr As String
    Dim Xname As String
    Dim ActCol As String
        
    For i = 6 To 13 ' the news part column numbers - they remain the same on every asset chart, dont need variable
        
        If i <> 10 Then 'because we skip the ff_id column, which is 10

            iLtr = ConvertToLetter(i)

            ' Name of x axis is the predictor
            Xname = Range("Derived!$" & iLtr & "$1").Value
            
            ' Add and use chart
            Set chrt = shLocal.Shapes.AddChart.Chart
            With chrt

            'Data
                .ChartType = xlXYScatter

                ' the data for scatterplot - (predictor, swings)
                For k = 1 To 8 ' for all timeframe columns in asset
                    ActCol = ConvertToLetter(GroupFirst + k - 1)
                    .SeriesCollection.NewSeries
                    .FullSeriesCollection(k).Name = "=Derived!$" & ActCol & "$1"
                    .FullSeriesCollection(k).Values = "=Derived!$" & ActCol & "$2:$" & ActCol & "$" & LastRow_D
                    .FullSeriesCollection(k).XValues = "=Derived!$" & iLtr & "$2:$" & iLtr & "$" & LastRow_D
                Next k

                'Titles
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Xname
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "swing"
                .Axes(xlCategory).HasMajorGridlines = True

                'Formatting
                .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
                .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = True
                .Axes(xlCategory, xlPrimary).Border.Color = black
                .Axes(xlCategory, xlPrimary).Border.Weight = 3
                .Axes(xlValue, xlPrimary).Border.Color = black
                .Axes(xlValue, xlPrimary).Border.Weight = 3


                .ChartStyle = 240
                .ChartArea.Height = 150
                .ChartArea.Width = 180
                .HasLegend = False

                'Position
                .Parent.Top = 0
                If i < 10 Then
                .Parent.Left = (i - 6) * 180
                Else
                .Parent.Left = (i - 7) * 180 'because we skip the ff_id column
                End If
            End With
        End If
    Next i
End Sub

Sub PartialTimeframeScatter(shLocal As Worksheet)
    'PARTIAL TIMEFRAME PLOTS SCATTER PLOTS
    
    Dim i As Integer
    Dim iLtr As String
    Dim j As Integer
    Dim jLtr As String
    Dim ChartGridRow As Integer
    Dim BegColLtr As String
    Dim EndColLtr As String
    Dim Xname As String
    Dim Yname As String
    
    ChartGridRow = 0
      
    For j = GroupFirst To (GroupFirst + 6) Step 2 'for all timeframe - step 2 because h/l pairs
            ChartGridRow = ChartGridRow + 1 'count the row for chart grid
            BegColLtr = ConvertToLetter(j)
            EndColLtr = ConvertToLetter(j + 1)
 
    For i = 6 To 13 'for all predictor

        If i <> 10 Then 'because we skip the ff_id column, which is 10

            iLtr = ConvertToLetter(i)

            ' Name of y axis is the swing timeframe
            Yname = Range("Derived!$" & BegColLtr & "$1").Value
            ' Name of x axis is the predictor
            Xname = Range("Derived!$" & iLtr & "$1").Value

            ' Add and use chart
            Set chrt = shLocal.Shapes.AddChart.Chart
            With chrt

            'Data
                .ChartType = xlXYScatter
                
                ' the data for scatterplot - (predictor, swings) - low and high
                .SeriesCollection.NewSeries
                .FullSeriesCollection(1).Name = "=Derived!$" & BegColLtr & "$1"
                .FullSeriesCollection(1).Values = "=Derived!$" & BegColLtr & "$2:$" & BegColLtr & "$" & LastRow_D
                .FullSeriesCollection(1).XValues = "=Derived!$" & iLtr & "$2:$" & iLtr & "$" & LastRow_D
                
                .SeriesCollection.NewSeries
                .FullSeriesCollection(2).Name = "=Derived!$" & EndColLtr & "$1"
                .FullSeriesCollection(2).Values = "=Derived!$" & EndColLtr & "$2:$" & EndColLtr & "$" & LastRow_D
                .FullSeriesCollection(2).XValues = "=Derived!$" & iLtr & "$2:$" & iLtr & "$" & LastRow_D


                'Titles
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Xname
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Yname
                .Axes(xlCategory).HasMajorGridlines = True

                'Trend lines
                .FullSeriesCollection(2).Trendlines.Add
                .FullSeriesCollection(2).Trendlines(1).Format.Line.DashStyle = msoLineSolid
                .FullSeriesCollection(2).Trendlines(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)

                .FullSeriesCollection(1).Trendlines.Add
                .FullSeriesCollection(1).Trendlines(1).Format.Line.DashStyle = msoLineSolid
                .FullSeriesCollection(1).Trendlines(1).Format.Line.ForeColor.RGB = RGB(0, 0, 255)

                'Formatting
                .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
                .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = True
                .Axes(xlCategory, xlPrimary).Border.Color = black
                .Axes(xlCategory, xlPrimary).Border.Weight = 3
                .Axes(xlValue, xlPrimary).Border.Color = black
                .Axes(xlValue, xlPrimary).Border.Weight = 3


                .ChartStyle = 240
                .ChartArea.Height = 150
                .ChartArea.Width = 180
               .HasLegend = False

                'Position
                .Parent.Top = ChartGridRow * 150
                If i < 10 Then
                .Parent.Left = (i - 6) * 180
                Else
                .Parent.Left = (i - 7) * 180 'because we skip the ff_id column
                End If
            End With
        End If
    Next i
    Next j
End Sub

Sub BoxPlot(shLocal As Worksheet, chTitle As String)
    'BOX PLOTS
    
    Dim boxChrt As Shape
    Dim j As Integer
    Dim r As Range
    Dim ActColLtr As String
                  
    Set boxChrt = shLocal.Shapes.AddChart2(406, xlBoxwhisker)
    boxChrt.Select

    ActiveChart.ChartTitle.Text = chTitle
    ActiveChart.ChartArea.Height = 300
    ActiveChart.ChartArea.Width = 540
    ActiveChart.HasLegend = True
    ActiveChart.Parent.Top = 780
    
    If chTitle = "Act-For > 0" Then 'Act-For>0
    ActiveChart.Parent.Left = 540
    Else
    ActiveChart.Parent.Left = 0 'Act-For<0
    End If

    For j = (GroupFirst) To (GroupFirst + 7) 'for all timeframe in asset make a box
        ActColLtr = ConvertToLetter(j)
        
        If chTitle = "Act-For > 0" Then 'Act-For>0
            For i = 2 To LastRow_D
                If Range("Derived!K" & i).Value > 0 Then
                    If r Is Nothing Then
                        Set r = Range("Derived!" & ActColLtr & i)
                    Else
                       Set r = Union(r, Range("Derived!" & ActColLtr & i))  ' here we have the range where the swing are from act-for > 0 range
                    End If
                End If
            Next i
        End If
        
        If chTitle = "Act-For < 0" Then 'Act-For<0
            For i = 2 To LastRow_D
                If Range("Derived!K" & i).Value < 0 Then
                    If r Is Nothing Then
                        Set r = Range("Derived!" & ActColLtr & i)
                    Else
                       Set r = Union(r, Range("Derived!" & ActColLtr & i))  ' here we have the range where the swing are from act-for < 0 range
                    End If
                End If
            Next i
        End If

        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection(j - GroupFirst + 1).Name = "=Derived!$" & ActColLtr & "$1"                 ' (j-Groupfirst +1) goes from 2 to 7 - very very ugly solution
        ActiveChart.FullSeriesCollection(j - GroupFirst + 1).Values = r

        Set r = Nothing   ' make new range

    Next j

    ActiveChart.SetElement (330)  ' first the data, the the grid
    
End Sub

Sub DoubleChartData(shLocal As Worksheet)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
            
    ' Make Header
    Sheets("Derived").Range(ConvertToLetter(GroupFirst) & "1:" & ConvertToLetter(GroupFirst + 7) & "1").Copy Destination:=Sheets(NewSheetName).Range("B75")
    Sheets(NewSheetName).Range("J75").Value = "Extreme"
    Sheets(NewSheetName).Range("K75").Value = "Bins"

    ' Make Row names
    Sheets(NewSheetName).Range("A76").Value = "max"
    Sheets(NewSheetName).Range("A77").Value = "min"
    Sheets(NewSheetName).Range("A78").Value = "med"
    Sheets(NewSheetName).Range("A79").Value = "FirstBinSize"

    'Count Highest Swing
    FString = "=MAX(Derived!" & GroupFirstLtr & "2:" & GroupFirstLtr & LastRow_D & ")"
    'Count for first cell
    Sheets(NewSheetName).Range("B76").Formula = FString
    'Count for row
    Sheets(NewSheetName).Range("B76").AutoFill Destination:=Sheets(NewSheetName).Range("B76:I76")

    'Count Lowest Swing
    FString = "=MIN(Derived!" & GroupFirstLtr & "2:" & GroupFirstLtr & LastRow_D & ")"
    'Count for first cell
    Sheets(NewSheetName).Range("B77").Formula = FString
    'Count for row
    Sheets(NewSheetName).Range("B77").AutoFill Destination:=Sheets(NewSheetName).Range("B77:I77")

    'Count Median
    FString = "=MEDIAN(Derived!" & GroupFirstLtr & "2:" & GroupFirstLtr & LastRow_D & ")"
    'Count for first cell
    Sheets(NewSheetName).Range("B78").Formula = FString
    'Count for row
    Sheets(NewSheetName).Range("B78").AutoFill Destination:=Sheets(NewSheetName).Range("B78:I78")

    'Extremes
    Sheets(NewSheetName).Range("J76").Formula = "=MAX(B76:I76)"
    Sheets(NewSheetName).Range("J77").Formula = "=MIN(B77:I77)"

    'Bin Number
    ' FirstBinSize - the average of high and low 5 minute median /4 - later bins: FirstBinSixe*2^x
    Sheets(NewSheetName).Range("J79").Formula = "=(ABS(D78)+ABS(E78))/8"
    Sheets(NewSheetName).Range("K76").Formula = "=ROUNDUP(LOG(ABS(J76/J79),2),0)"
    Sheets(NewSheetName).Range("K77").Formula = "=IF(J77<0,-1*ROUNDUP(LOG(ABS(J77/J79),2),0),ROUNDUP(LOG(ABS(J77/J79),2),0))"

    ' Build the swing column

    Sheets(NewSheetName).Range("B81").Value = "A-f<0"
    Sheets(NewSheetName).Range("K81").Value = "A-f>0"


    'Build column space
    If (Sheets(NewSheetName).Range("K77").Value < 0) Then
    i = Sheets(NewSheetName).Range("K77").Value ' the negativ columns
    Else
    i = 0
    End If
    If (Sheets(NewSheetName).Range("K76").Value > 0) Then
    j = Sheets(NewSheetName).Range("K76").Value ' the positive columns
    Else
    j = 0
    End If

    AllBinNbr = 82 + Abs(i) + j

    For k = i To j
    Sheets(NewSheetName).Range("A" & 82 + Abs(i) + k).Value = k ' with the +abs the first row always 82
    Next k

    'Build Bin limits
    FString = "=SIGN(A82)*$J$79*POWER(2,ABS(A82))"
    'Count for first cell
    Sheets(NewSheetName).Range("J82").Formula = FString
    'Count for Column
    Sheets(NewSheetName).Range("J82").AutoFill Destination:=Sheets(NewSheetName).Range("J82:J" & AllBinNbr)
    Sheets(NewSheetName).Range("J82:J" & AllBinNbr).NumberFormat = "0.00000"

    'Count cases A-f<0, and between limits
    FString = "=COUNTIFS(Derived!$K$2:$K$" & LastRow_D & ",""<0"",Derived!" & GroupFirstLtr & "$2:" & GroupFirstLtr & "$" & LastRow_D & ",""<"" & " & NewSheetName & "!$J82,Derived!" & GroupFirstLtr & "$2:" & GroupFirstLtr & "$" & LastRow_D & ","">"" & " & NewSheetName & "!$J81)"
            ' FString = "=COUNTIFS(Derived!$K$2:$K$86,""<0"",Derived!N$2:N$86,""<"" & audusd!$J82,Derived!N$2:N$86,"">"" & audusd!$J81)"
    'Count for first cell
    Sheets(NewSheetName).Range("B82").Formula = FString
    'Count for first Column
    Sheets(NewSheetName).Range("B82").AutoFill Destination:=Sheets(NewSheetName).Range("B82:B" & AllBinNbr)
    ' Copy formula through every row
            For k = 82 To AllBinNbr
            Sheets(NewSheetName).Range("B" & k).AutoFill Destination:=Sheets(NewSheetName).Range("B" & k & ":I" & k)
            Next k

    'Count cases A-f>0, and between limits
    FString = "=COUNTIFS(Derived!$K$2:$K$" & LastRow_D & ","">0"",Derived!" & GroupFirstLtr & "$2:" & GroupFirstLtr & "$" & LastRow_D & ",""<"" & " & NewSheetName & "!$J82,Derived!" & GroupFirstLtr & "$2:" & GroupFirstLtr & "$" & LastRow_D & ","">"" & " & NewSheetName & "!$J81)"
                ' FString = "=COUNTIFS(Derived!$K$2:$K$86,"">0"",Derived!N$2:N$86,""<"" & audusd!$J82,Derived!N$2:N$86,"">"" & audusd!$J81)"
    'Count for first cell
    Sheets(NewSheetName).Range("K82").Formula = FString
    'Count for first Column
    Sheets(NewSheetName).Range("K82").AutoFill Destination:=Sheets(NewSheetName).Range("K82:K" & AllBinNbr)
    ' Copy formula through every row
            For k = 82 To AllBinNbr
            Sheets(NewSheetName).Range("K" & k).AutoFill Destination:=Sheets(NewSheetName).Range("K" & k & ":R" & k)
            Next k

End Sub

Sub DoubleChartDraw(shLocal As Worksheet)

    Dim i As Integer
    Dim BegColLtr As String
    Dim EndColLtr As String

'A-f < 0 side
    For i = 2 To 8 Step 2 'The column number of low data
        BegColLtr = ConvertToLetter(i)
        EndColLtr = ConvertToLetter(i + 1)

    'Left
      ' Add and use chart
        Set chrt = shLocal.Shapes.AddChart.Chart
        With chrt
        .ChartType = xlBarClustered

    'Data
        .SeriesCollection.NewSeries
        .FullSeriesCollection(1).Values = "=" & NewSheetName & "!$" & BegColLtr & "$82:$" & BegColLtr & AllBinNbr
        .SeriesCollection.NewSeries
        .FullSeriesCollection(2).Values = "=" & NewSheetName & "!$" & EndColLtr & "$82:$" & EndColLtr & AllBinNbr

        .FullSeriesCollection(1).XValues = "=" & NewSheetName & "!$J$82:$J$" & AllBinNbr ' to add the vertical axis values

    'Format
        .Axes(xlValue).ReversePlotOrder = True
        .ChartGroups(1).GapWidth = 60
        .SetElement (msoElementChartTitleCenteredOverlay)
        .ChartTitle.Text = "A-F<0"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = False

        .ChartArea.Height = 150
        .ChartArea.Width = 180
        .HasLegend = False

        .Axes(xlValue).Delete

    'Position
        .Parent.Top = (i / 2) * 150
        .Parent.Left = 7 * 180

        End With
    Next i

'A-f > 0 side
    For i = 11 To 17 Step 2 'The column number of low data
        BegColLtr = ConvertToLetter(i)
        EndColLtr = ConvertToLetter(i + 1)

    'Left
      ' Add and use chart
        Set chrt = shLocal.Shapes.AddChart.Chart
        With chrt
        .ChartType = xlBarClustered

    'Data
        .SeriesCollection.NewSeries
        .FullSeriesCollection(1).Values = "=" & NewSheetName & "!$" & BegColLtr & "$82:$" & BegColLtr & AllBinNbr
        .SeriesCollection.NewSeries
        .FullSeriesCollection(2).Values = "=" & NewSheetName & "!$" & EndColLtr & "$82:$" & EndColLtr & AllBinNbr

        .FullSeriesCollection(1).XValues = "=" & NewSheetName & "!$J$82:$J$" & AllBinNbr ' to add the vertical axis values

    'Format
        .Axes(xlValue).ReversePlotOrder = False
        .ChartGroups(1).GapWidth = 60
        .SetElement (msoElementChartTitleCenteredOverlay)
        .ChartTitle.Text = "A-F>0"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = False

        .ChartArea.Height = 150
        .ChartArea.Width = 180
        .HasLegend = False

        .Axes(xlValue).Delete

    'Position
        .Parent.Top = ((i - 9) / 2) * 150
        .Parent.Left = 7.7 * 180

        End With
    Next i

End Sub


' Converts Column Number to Letter
Function ConvertToLetter(ColNum As Integer) As String
    ConvertToLetter = Split(Cells(1, ColNum).Address, "$")(1)
End Function

'Check if sheet name is used already
Function SheetExists(SheetName As String)
    On Error GoTo no:
    WorksheetName = Worksheets(SheetName).Name
    SheetExists = True
    Exit Function
no:
    SheetExists = False
End Function
