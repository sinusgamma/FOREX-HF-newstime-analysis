Attribute VB_Name = "Module5"
Option Explicit

Public Type DoubleLong
    db As Double
    lg As Long
End Type

Sub t1_tickData()
Attribute t1_tickData.VB_ProcData.VB_Invoke_Func = " \n14"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' as openprice is the first row before news, not after
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LastRow As Long
Dim LastColumnNbr As Long
Dim LastColumnLtr As String
Dim key As Variant
Dim fRow As Long ' first row of news
Dim lRow As Long ' last row of news
Dim oRow As Long ' first after news row - the open price row
Dim dictAr(3) As Long
Dim FString As String
    
    'Check if chart exists and delete old data before rebuild
    If (CheckIfSheetExists("TickData") <> True) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "TickData"
    Else
        Sheets("TickData").Cells.Delete
    End If
    
    'Copy original to TickData sheet
    Sheets(1).UsedRange.Copy Sheets("TickData").Range("A1")
    
    ' Got to the TickData sheet
    Sheets("TickData").Activate

    ' Sheet limits on TickData sheet
    LastRow = Sheets("TickData").UsedRange.Rows.Count 'Number of Rows in TickData sheet
    LastColumnNbr = Sheets("TickData").UsedRange.Columns.Count 'Number of columns in TickData sheet , NEEDS recalculate later
    LastColumnLtr = ConvertToLetter(LastColumnNbr)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the unique news_id-s and the first and last row of news
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim newsDict As New Scripting.dictionary
    Set newsDict = getUnique(Sheets("TickData").Range("D2:D" & LastRow))
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set the date-type
    Sheets("TickData").Range("A:A").NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    Sheets("TickData").Range("E:E").NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Count Time Difference
    Sheets("TickData").Range("N1").Value = "time_diff"
    FString = "=IF(E2<A2,TEXT(A2-E2,""mm:ss.000"")*86400,TEXT(E2-A2,""mm:ss.000"")*(-86400))" 'because excel doesnt like negative time
    'Count for first cell
    Sheets("TickData").Range("N2").Formula = FString
    'Count for column
    Sheets("TickData").Range("N2").AutoFill Destination:=Sheets("TickData").Range("N2:N" & LastRow)
    Sheets("TickData").Range("N:N").NumberFormat = "0.000"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Before, After, Open Dictionary FillUp
    Sheets("TickData").Range("O1").Value = "BAO"
    FString = "=IF(E2>A2,""B"",""A"")"
    'Count for first cell
    Sheets("TickData").Range("O2").Formula = FString
    'Count for column
    Sheets("TickData").Range("O2").AutoFill Destination:=Sheets("TickData").Range("O2:O" & LastRow)
    For Each key In newsDict.Keys()
        fRow = newsDict(key)(0)
        lRow = newsDict(key)(1)
        Sheets("TickData").Range("O" & fRow).Value = "BBB" 'indicate the first row of news Before
        Sheets("TickData").Range("O" & lRow).Value = "AAA" 'indicate the last row of news After
        'FString = "=MATCH(MIN(IF(N" & frow & ":N" & lrow & ">0,N" & frow & ":N" & lrow & ")),N" & frow & ":N" & lrow & ",0)"
        oRow = Evaluate("MATCH(MIN(IF(N" & fRow & ":N" & lRow & ">0,N" & fRow & ":N" & lRow & ")),N" & fRow & ":N" & lRow & ",0)") 'first afternews row - REALATIVE FROM the FIRST row
        Sheets("TickData").Range("O" & fRow + oRow - 2).Value = "OOO" ' this gives the first row before news, instead after news
        dictAr(0) = fRow
        dictAr(1) = lRow
        dictAr(2) = oRow + fRow - 2 ' this gives the first row before news, instead after news
        newsDict(key) = dictAr ' from now newsDict has the first last and open rows set
    Next key
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' count Ask and Bid swing from open Bid price, AskJump, BidJump, FullJump
    Sheets("TickData").Range("P1").Value = "Ask-OBid"
    Sheets("TickData").Range("Q1").Value = "Bid-OBid"
    Sheets("TickData").Range("S1").Value = "AskJump"
    Sheets("TickData").Range("T1").Value = "BidJump"
    Sheets("TickData").Range("U1").Value = "FullJump"
    Sheets("TickData").Range("V1").Value = "Ask-OAsk"
    Sheets("TickData").Range("W1").Value = "Bid-OAsk"
    
    Dim FStringAsk As String, FStringBid As String
    For Each key In newsDict.Keys()
        fRow = newsDict(key)(0)
        lRow = newsDict(key)(1)
        oRow = newsDict(key)(2) 'oRow is the first row before news, not after
        
    ' count Ask and Bid swing from open Bid price
        FStringAsk = "=B" & fRow & "-$C$" & oRow & ""
        FStringBid = "=C" & fRow & "-$C$" & oRow & ""
        'Count for first cell
        Sheets("TickData").Range("P" & fRow).Formula = FStringAsk
        Sheets("TickData").Range("Q" & fRow).Formula = FStringBid
        'Count for column
        Sheets("TickData").Range("P" & fRow).AutoFill Destination:=Sheets("TickData").Range("P" & fRow & ":P" & lRow)
        Sheets("TickData").Range("Q" & fRow).AutoFill Destination:=Sheets("TickData").Range("Q" & fRow & ":Q" & lRow)
        
    ' Count AskJump, BidJump, FullJump
        'Count for first cell
        Sheets("TickData").Range("S" & fRow + 1).Formula = "=P" & fRow + 1 & "-P" & fRow & ""
        Sheets("TickData").Range("T" & fRow + 1).Formula = "=Q" & fRow + 1 & "-Q" & fRow & ""
        Sheets("TickData").Range("U" & fRow + 1).Formula = "=IF(P" & fRow + 1 & ">P" & fRow & ",MAX(ABS(Q" & fRow + 1 & "-P" & fRow & "),ABS(P" & fRow + 1 & "-Q" & fRow & ")),-MAX(ABS(Q" & fRow + 1 & "-P" & fRow & "),ABS(P" & fRow + 1 & "-Q" & fRow & ")))"
        'Count for column
        Sheets("TickData").Range("S" & fRow + 1).AutoFill Destination:=Sheets("TickData").Range("S" & fRow + 1 & ":S" & lRow)
        Sheets("TickData").Range("T" & fRow + 1).AutoFill Destination:=Sheets("TickData").Range("T" & fRow + 1 & ":T" & lRow)
        Sheets("TickData").Range("U" & fRow + 1).AutoFill Destination:=Sheets("TickData").Range("U" & fRow + 1 & ":U" & lRow)
        
    ' count Ask and Bid swing from open Ask price
        FStringAsk = "=B" & fRow & "-$B$" & oRow & ""
        FStringBid = "=C" & fRow & "-$B$" & oRow & ""
        'Count for first cell
        Sheets("TickData").Range("V" & fRow).Formula = FStringAsk
        Sheets("TickData").Range("W" & fRow).Formula = FStringBid
        'Count for column
        Sheets("TickData").Range("V" & fRow).AutoFill Destination:=Sheets("TickData").Range("V" & fRow & ":V" & lRow)
        Sheets("TickData").Range("W" & fRow).AutoFill Destination:=Sheets("TickData").Range("W" & fRow & ":W" & lRow)
    Next key
    

    Sheets("TickData").Range("B:C").NumberFormat = "0.00000"
    Sheets("TickData").Range("P:W").NumberFormat = "0.00000"
    
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Count Spread
    Sheets("TickData").Range("R1").Value = "Spread"
    FString = "=B2-C2"
    'Count for first cell
    Sheets("TickData").Range("R2").Formula = FString
    'Count for column
    Sheets("TickData").Range("R2").AutoFill Destination:=Sheets("TickData").Range("R2:R" & LastRow)
   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GENERATE THE EVENT CHARTS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SetEventCharts newsDict
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'COUNT THE STATISTICS on "TickStat"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    t2_getStat newsDict
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

' Converts Column Number to Letter
Function ConvertToLetter(ColNum As Long) As String
    ConvertToLetter = Split(Cells(1, ColNum).Address, "$")(1)
End Function

'Get unique news ID as key, and first row, last row as value - the first row after news is empty, returns array(firts, last, beginning)
Function getUnique(dataSet As Range) As Object '<<< remove Column

    Dim dkey() As String
    Dim dValue(3) As Long 'first row, last row, first row after news
    Dim rowCount As Long
    Dim dictionary As Object
    Dim i As Long

    rowCount = dataSet.Rows.Count
    Set dictionary = CreateObject("Scripting.Dictionary")

    ReDim dkey(rowCount) 'the new size of data array
    For i = 1 To UBound(dkey)
        dkey(i) = dataSet.Cells(i, 1).Value    '<<< using Cells
        If (dkey(i) > dkey(i - 1)) Then
            dValue(0) = i + 1 ' the first row of news
            dictionary(dkey(i)) = dValue
        Else
            dValue(1) = i + 1 ' the last row of news
            dictionary(dkey(i)) = dValue
        End If
    Next i

'    Dim v As Variant
'    For Each v In dictionary.Keys()
'        Debug.Print v    '<<<
'    Next v
    
    Set getUnique = dictionary
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call the charts of last x news event
''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetEventCharts(dict As Scripting.dictionary)
    
    If (CheckIfSheetExists("ChartsClose") <> True) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "ChartsClose"
    Else
        Sheets("ChartsClose").Cells.Delete
    End If
    
    'Check if chart exists and delete old data before rebuild
    If (CheckIfSheetExists("ChartsFar") <> True) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "ChartsFar"
    Else
        Sheets("ChartsFar").Cells.Delete
    End If
            
    Dim i As Long
    Dim counter As Long
    Dim axesDistance As Double
    Dim dkey As Long
    Dim ditem() As Long
    counter = 0
    
    'Draw the chart of last news events
    For i = dict.Count - 1 To dict.Count - 14 Step -1
        dkey = dict.Keys(i)
        ditem = dict.Items(i)
                    
        Call DrawEventChartClose(dkey, ditem, counter)
            'use the same axes distances on close charts
            If (counter = 0) Then
                axesDistance = ActiveChart.Axes(xlValue).MajorUnit
            Else
                ActiveChart.Axes(xlValue).MajorUnit = axesDistance
            End If
                    
        Call DrawEventChartFar(dkey, ditem, counter)
            
        counter = counter + 1
    Next i

End Sub

'Check if sheet exist
Function CheckIfSheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    CheckIfSheetExists = False
    For Each ws In Worksheets
        If SheetName = ws.Name Then
            CheckIfSheetExists = True
            Exit Function
        End If
    Next ws
End Function

'Generate Each Event Chart Close
Sub DrawEventChartClose(news_id As Long, ditem() As Long, counter As Long) ' ditem(firstrow, lastrow, openrow) of news event
    
    'activate the sheet to prevent errors
    Worksheets("ChartsClose").Activate
    
    Dim beginRow As Long
    
    If (ditem(0) > (ditem(2) - 8)) Then ' don't let the beginRow to use the end of the other news event data
        beginRow = ditem(0)
    Else
    beginRow = ditem(2) - 8
    End If

    Dim rSourceData As Range
    Set rSourceData = Worksheets("TickData").Range("P" & beginRow & ":U" & ditem(2) + 40)
        
    Dim oChart As Chart
    Set oChart = Sheets("ChartsClose").Shapes.AddChart2(201, xlColumnClustered).Chart
      
    With oChart
        .ChartTitle.Text = "ID:" & news_id & " " & Sheets("TickData").Range("E" & ditem(2)).Value
        .SetSourceData rSourceData
        .PlotBy = xlColumns 'or xlRows
        .FullSeriesCollection(1).ChartType = xlLine
        .FullSeriesCollection(2).ChartType = xlLine
        .FullSeriesCollection(3).ChartType = xlColumnClustered
        .FullSeriesCollection(4).ChartType = xlColumnClustered
        .FullSeriesCollection(5).ChartType = xlColumnClustered
        .FullSeriesCollection(5).ChartType = xlColumnClustered
        .FullSeriesCollection(1).XValues = "=TickData!$N$" & beginRow & ":$N" & ditem(2) + 40
        
        .ChartArea.Height = 500
        .ChartArea.Width = 500
        .HasLegend = False
        .Parent.Top = Fix(counter / 3) * 500 ' Fix gets the Long part
        .Parent.Left = (counter Mod 3) * 500
        .SetElement (msoElementPrimaryCategoryGridLinesMajor) 'axes
        .SetElement (msoElementPrimaryValueGridLinesMinorMajor) ' no legend
        
        'wider columns
        .ChartGroups(1).Overlap = 0
        .ChartGroups(1).GapWidth = 50
        
        'make main horizontal axislines wider
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoTrue
        .Axes(xlValue).MajorGridlines.Format.Line.Weight = 2
        
        'format datapoints on lines and column colors
        With .FullSeriesCollection(1)
            .Name = "=""Ask"""
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Fill.Visible = msoTrue
            .Format.Fill.Solid
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Line.Visible = msoTrue
            .Format.Line.Weight = 2
        End With
        With .FullSeriesCollection(2)
            .Name = "=""Bid"""
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Fill.Visible = msoTrue
            .Format.Fill.Solid
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Line.Visible = msoTrue
            .Format.Line.Weight = 2
        End With
        With .FullSeriesCollection(3)
            .Name = "=""Spread"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Fill.ForeColor.Brightness = 0.400000006
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(4)
            .Name = "=""AskJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = RGB(0, 176, 240)
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(5)
            .Name = "=""BidJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = RGB(255, 153, 102)
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(6)
            .Name = "=""FullJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Fill.ForeColor.Brightness = 0.349999994
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
    End With ' oChart
    
    'To select the chart for getting the axesdistance out of sub - very very ugly solution
    oChart.Axes(xlValue).MajorGridlines.Select
    
End Sub

'Generate Each Event Chart Far
Sub DrawEventChartFar(news_id As Long, ditem() As Long, counter As Long) ' ditem(firstrow, lastrow, openrow) of news event

    Dim rSourceData As Range
    Set rSourceData = Worksheets("TickData").Range("P" & ditem(0) & ":U" & ditem(1))
    Dim oChart As Chart
    Set oChart = Sheets("ChartsFar").Shapes.AddChart2(201, xlColumnClustered).Chart
      
    With oChart
        .ChartTitle.Text = "ID:" & news_id & " " & Sheets("TickData").Range("E" & ditem(2)).Value
        .SetSourceData rSourceData
        .PlotBy = xlColumns 'or xlRows
        .FullSeriesCollection(1).ChartType = xlLine
        .FullSeriesCollection(2).ChartType = xlLine
        .FullSeriesCollection(3).ChartType = xlColumnClustered
        .FullSeriesCollection(4).ChartType = xlColumnClustered
        .FullSeriesCollection(5).ChartType = xlColumnClustered
        .FullSeriesCollection(5).ChartType = xlColumnClustered
        .FullSeriesCollection(1).XValues = "=TickData!$N$" & ditem(0) & ":$N" & ditem(1)
        
        .ChartArea.Height = 500
        .ChartArea.Width = 500
        .HasLegend = False
        .Parent.Top = Fix(counter / 3) * 500 ' Fix gets the Long part
        .Parent.Left = (counter Mod 3) * 500
        
        'wider columns
        .ChartGroups(1).Overlap = 0
        .ChartGroups(1).GapWidth = 50
        
        'format datapoints on lines and column colors
        With .FullSeriesCollection(1)

        End With
        With .FullSeriesCollection(2)

        End With
        With .FullSeriesCollection(3)
            .Name = "=""Spread"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Fill.ForeColor.Brightness = 0.400000006
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(4)
            .Name = "=""AskJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = RGB(0, 176, 240)
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(5)
            .Name = "=""BidJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = RGB(255, 153, 102)
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
        With .FullSeriesCollection(6)
            .Name = "=""FullJump"""
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
            .Format.Fill.ForeColor.TintAndShade = 0
            .Format.Fill.ForeColor.Brightness = 0.349999994
            .Format.Fill.Transparency = 0
            .Format.Fill.Solid
        End With
    End With ' oChart
End Sub

'Round numbers in range to decimal places
Sub roundRange(workRange As Range)
    Dim x As Range
    Application.ScreenUpdating = False
    For Each x In workRange
        If IsNumeric(x.Value) Then
            x.Value = Round(x.Value, 5)
        End If
    Next x
    Application.ScreenUpdating = True
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calculate the news event statistics from "TickData" sheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub t2_getStat(dict As Scripting.dictionary)

    'Check if chart exists and delete old data before rebuild
    If (CheckIfSheetExists("TickStat") <> True) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "TickStat"
    Else
        Sheets("TickStat").Cells.Delete
    End If
    
    Sheets("TickStat").Activate

    'Set Header
    Sheets("TickStat").Range("A1").Value = "news_id"
    Sheets("TickStat").Range("B1").Value = "news_time"
    
    Sheets("TickStat").Range("C1").Value = "SpBeO Avr"
    Sheets("TickStat").Range("D1").Value = "SpBeO Max"
    Sheets("TickStat").Range("E1").Value = "SpAfO Max"
    Sheets("TickStat").Range("F1").Value = "SpAfO MaxTm"
    Sheets("TickStat").Range("G1").Value = "SpAfMax Med30"
    Sheets("TickStat").Range("H1").Value = "SpAfMax Avr30"
    
    Sheets("TickStat").Range("I1").Value = "AJBeAvr"
    Sheets("TickStat").Range("J1").Value = "BJBeAvr"
    Sheets("TickStat").Range("K1").Value = "AJBeMax"
    Sheets("TickStat").Range("L1").Value = "BJBeMax"
    Sheets("TickStat").Range("M1").Value = "AJAfMax"
    Sheets("TickStat").Range("N1").Value = "AJMaxTm"
    Sheets("TickStat").Range("O1").Value = "BJAfMax"
    Sheets("TickStat").Range("P1").Value = "BJMaxTm"
    Sheets("TickStat").Range("Q1").Value = "AJAfMed"
    Sheets("TickStat").Range("R1").Value = "BJAfMed"
    Sheets("TickStat").Range("S1").Value = "AJAfAvr"
    Sheets("TickStat").Range("T1").Value = "BJrAfAvr"

    
    Sheets("TickStat").Range("C:T").NumberFormat = "0.00000"
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Loop trough the news events
    Dim newsid As Variant
    Dim fRow As Long 'first row of event
    Dim lRow As Long 'last row of event
    Dim oRow As Long 'open row of event - the first row after news_time
    Dim sRow As Long: sRow = 5 'the first row of mass data
    Dim mRow As Long: mRow = sRow 'row counter in mass data on tickStat
    Dim maxAddr As DoubleLong ' custom type for maxAddress() return
    Dim LastRow As Long
    Dim formatRange As Range

        
    For Each newsid In dict.Keys
        mRow = mRow + 1 'so the first used row for data is the second on TickStat sheet / first is header
        fRow = dict(newsid)(0)
        lRow = dict(newsid)(1)
        oRow = dict(newsid)(2) 'oRow is the first row before news, not after
        
        'News ID and Time
        Sheets("TickStat").Range("A" & mRow).Value = newsid
        Sheets("TickStat").Range("B" & mRow).Value = Sheets("TickData").Range("E" & fRow).Value
        
        'Spread Statistics - spread always positive number - only magnitude
        Sheets("TickStat").Range("C" & mRow) = "=AVERAGE(TickData!R" & fRow & ":R" & oRow & ")"
        Sheets("TickStat").Range("D" & mRow) = "=MAX(TickData!R" & fRow & ":R" & oRow & ")"
        
        maxAddr = maxAddress(Sheets("TickData").Range("R" & oRow & ":R" & lRow)) 'goes from oRow, because the broker can be slow, and between oRow and oRow+1 can be swings, and the address is important, because that can be intrade

        Sheets("TickStat").Range("E" & mRow) = maxAddr.db
        Sheets("TickStat").Range("F" & mRow) = Sheets("TickData").Range("N" & maxAddr.lg).Value
        
        Sheets("TickStat").Range("G" & mRow).Formula = "=MEDIAN(TickData!R" & maxAddr.lg + 1 & ":R" & maxAddr.lg + 30 & ")"
        Sheets("TickStat").Range("H" & mRow) = "=AVERAGE(TickData!R" & maxAddr.lg + 1 & ":R" & maxAddr.lg + 30 & ")"
        
        'BidJump and AskJump statistics - can be positive and negative number !!! - magnitude and direction
        Sheets("TickStat").Range("I" & mRow) = "=AVERAGE(TickData!S" & fRow & ":S" & oRow & ")"
        Sheets("TickStat").Range("J" & mRow) = "=AVERAGE(TickData!T" & fRow & ":T" & oRow & ")"
          
    
    Next 'newsid
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    LastRow = Sheets("TickData").UsedRange.Rows.Count
    Set formatRange = Union(Sheets("TickStat").Range("C" & sRow & ":E" & LastRow), Sheets("TickStat").Range("G" & sRow & ":H" & LastRow))
    formatStatData formatRange
    formatStatDataTime (Sheets("TickStat").Range("F:F"))
    formatDataHeader (Sheets("TickStat").Range("1:1"))
    colorStatHeader

End Sub

'Find max and its address in range
Public Function maxAddress(rng As Range) As DoubleLong

Dim cell As Range

For Each cell In rng
    If IsNumeric(cell.Value2) Then
        If cell.Value2 > maxAddress.db Then
            maxAddress.db = cell.Value2
            maxAddress.lg = cell.Row
        End If
    End If
Next cell
End Function

Sub formatStatData(formatRange As Range)
'format data on TickStat
    formatRange.FormatConditions.AddColorScale ColorScaleType:=2
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 16776444
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueHighestValue
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
End Sub

Sub formatStatDataTime(formatRange As Range)
'format data time on TickStat
    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 26367
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With formatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
End Sub

Sub formatDataHeader(formatRange As Range)
'Format header on tickstat, freeze first row and fit cell width
    With formatRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Sub colorStatHeader()
'Spread format
    Range("C1:H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

