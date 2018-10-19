Attribute VB_Name = "Module4"
Sub a2_CalculateDerivedData()

'
' Uses sheet "Data" to copy News and Assets data to Derived sheet
'
'!!!!!!!!!!!!!!!! Firt sheet must be named : Data

    Sheets("Derived").Activate ' Got to the Derived sheet

    Dim NewColNbr As Integer
    Dim NewCol_F As String
    Dim NewCol_L As String
    Dim i As Integer
    Dim j As Integer
    Dim ActCol As String
    Dim LastRow As Integer
    Dim OpenColNbr As Integer
    Dim OpenColLtr As String
    Dim FString As String
    Dim TempRange As Range
    
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NEWS Copy
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Copy first cloumns with Derived data
        Sheets("Data").Columns("A:J").Copy Destination:=Sheets("Derived").Columns("A:J")
        
    ' Count number of Rows
        LastRow = Sheets("Data").UsedRange.Rows.Count
        
    'Calculate Derived derivatives
            'Actual-Forecast
        Sheets("Derived").Range("K1") = "Act-For"
        Sheets("Derived").Range("K2").Formula = "=IF(OR(ISBLANK(Data!F2),ISBLANK(Data!G2)),"""",Data!F2-Data!G2)"
        Sheets("Derived").Range("K2").AutoFill Destination:=Sheets("Derived").Range("K2:K" & LastRow)
        
            'Actual-Previous
        Sheets("Derived").Range("L1") = "Act-Pre"
        Sheets("Derived").Range("L2").Formula = "=IF(OR(ISBLANK(Data!F2),ISBLANK(Data!H2)),"""",Data!F2-Data!H2)"
        Sheets("Derived").Range("L2").AutoFill Destination:=Sheets("Derived").Range("L2:L" & LastRow)
        
            'Previous - ModifiedFrom
        Sheets("Derived").Range("M1") = "Pre-MFr"
        Sheets("Derived").Range("M2").Formula = "=IF(OR(ISBLANK(Data!H2),ISBLANK(Data!I2)),"""",Data!H2-Data!I2)"
        Sheets("Derived").Range("M2").AutoFill Destination:=Sheets("Derived").Range("M2:M" & LastRow)
        
            ' Conditional Formating News
        ConditionalFormating (Sheets("Derived").Range("F2:I" & LastRow)) ' the same nature

        ConditionalFormating (Sheets("Derived").Range("K2:K" & LastRow))
        ConditionalFormating (Sheets("Derived").Range("L2:L" & LastRow))
        ConditionalFormating (Sheets("Derived").Range("M2:M" & LastRow))
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ASSETS Copy
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        
    LastRow = Sheets("Data").UsedRange.Rows.Count 'Count number of Rows
    LastColumnSource = Sheets("Data").UsedRange.Columns.Count 'Number of columns in source sheet
            
    OpenColNbr = 11 'the first asset column - xxxyyy currency openprices in Data sheet - will be incremented for other assets
    OpenColLtr = "K" 'the first asset column - xxxyyy currency openprices in Data sheet - will be incremented for other assets
    


' Calculate assets - from open - coobroot
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Check if exist data - this checks the openPrice column in Data sheet
    While Not IsEmpty(Sheets("Data").Range(OpenColLtr & "1"))
    
        ' Calculate Swing from Open
        '''''''''''''''''''''''''''
        ' Finds first blank column
        NewColNbr = Sheets("Derived").Cells(1, Columns.Count).End(xlToLeft).Column + 1 ' get first empty column number
        NewCol_F = ConvertToLetter(NewColNbr) ' first created blank number
        NewCol_L = ConvertToLetter(NewColNbr + 7) ' last created blank number
    
        ' Set Price-OpenPrice Header
        Sheets("Data").Range(OpenColLtr & "1").Offset(0, 1).Resize(RowSize:=1, ColumnSize:=8).Copy Destination:=Sheets("Derived").Range(NewCol_F & "1") ' get the openPriceColumn - offset - resize - so get headers
        ' Set Price-Openprice formula
        FString = "=IF(OR(ISBLANK(Data!" & ConvertToLetter(OpenColNbr + 1) & "2),ISBLANK(Data!$" & ConvertToLetter(OpenColNbr) & "2)),"""",Data!" & ConvertToLetter(OpenColNbr + 1) & "2-Data!$" & ConvertToLetter(OpenColNbr) & "2)"
                        ' =IF(OR(ISBLANK(Data!L2),ISBLANK(Data!$K2)),"",Data!L2-Data!$K2) - check if source exists
        
        Sheets("Derived").Range(NewCol_F & "2").Formula = FString
        ' Copy formula to first row
        Sheets("Derived").Range(NewCol_F & "2").AutoFill Destination:=Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & 2)
        ' Copy formula through every column
            For i = NewColNbr To (NewColNbr + 7)
            ActCol = ConvertToLetter(i)
            Sheets("Derived").Range(ActCol & "2").AutoFill Destination:=Sheets("Derived").Range(ActCol & "2:" & ActCol & LastRow)
            Next i
            
        ' Format Number
        Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & LastRow).NumberFormat = "0.00000"
        ' Conditional Formating
        ConditionalFormatingBars (Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & LastRow))
        ' Format Border
        MakeBorder (Sheets("Derived").Range(NewCol_F & "1:" & NewCol_L & LastRow))
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           
        ' Calculate Coob Root Transformation
        ''''''''''''''''''''''''''''''''''''
        ' Finds first blank column
        'NewColNbr = Sheets("Derived").Cells(1, Columns.Count).End(xlToLeft).Column + 1 ' get first empty column number
        'NewCol_F = ConvertToLetter(NewColNbr) ' first created blank number
        'NewCol_L = ConvertToLetter(NewColNbr + 7) ' last created blank number
    
        ' Set Price-OpenPrice Header and _CR for coobroot
        'Sheets("Data").Range(OpenColLtr & "1").Offset(0, 1).Resize(RowSize:=1, ColumnSize:=8).Copy Destination:=Sheets("Derived").Range(NewCol_F & "1") ' get the openPriceColumn - offset - resize - so get headers
        'For Each MyCell In Sheets("Derived").Range(NewCol_F & "1:" & NewCol_L & "1")
         '   MyCell.Value = MyCell.Value & "_CR"
        'Next
        ' Set Coob Root formula - coob root of the calculated swing
        'FString = "=IF(" & ConvertToLetter(NewColNbr - 8) & "2="""","""",POWER(" & ConvertToLetter(NewColNbr - 8) & "2,1/3))" ' -8 because it is calculated from swings
        'Sheets("Derived").Range(NewCol_F & "2").Formula = FString
        ' Copy formula to first row
        'Sheets("Derived").Range(NewCol_F & "2").AutoFill Destination:=Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & 2)
        ' Copy formula through every column
        '    For i = NewColNbr To (NewColNbr + 7)
        '    ActCol = ConvertToLetter(i)
        '    Sheets("Derived").Range(ActCol & "2").AutoFill Destination:=Sheets("Derived").Range(ActCol & "2:" & ActCol & LastRow)
        '    Next i
            
        ' Format Number
        'Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & LastRow).NumberFormat = "0.00000"
        ' Conditional Formating
        'ConditionalFormatingBars (Sheets("Derived").Range(NewCol_F & "2:" & NewCol_L & LastRow))
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        OpenColNbr = OpenColNbr + 11 ' Find the next asset 11 column right
        OpenColLtr = ConvertToLetter(OpenColNbr)
    Wend
    
    ' Freeze Panels - at N2 - here begins the asset part - FreezPanes works only on active window
    Range("N2").Select
    ActiveWindow.FreezePanes = True
    
    'Customize Header
     MakeHeader (Sheets("Derived").Rows(1))
     
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Clear Blank Cells - no formula, no error value - only empty cell
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rng As Range
    Dim cell As Range
    Set rng = Sheets("Derived").UsedRange
    For Each cell In rng
        If cell = "" Then
           cell.ClearContents
        End If
    Next cell
    
    
End Sub

' Converts Column Number to Letter
Function ConvertToLetter(ColNum As Integer) As String
    ConvertToLetter = Split(Cells(1, ColNum).Address, "$")(1)
End Function


' ConditionalFormats the range to highlight swings
Sub ConditionalFormating(formatRange As Range)

    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
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
    
' ConditionalFormats the range to highlight swings with bars
Sub ConditionalFormatingBars(formatRange As Range)
    
    formatRange.FormatConditions.AddDatabar
    formatRange.FormatConditions(formatRange.FormatConditions.Count).ShowValue = True
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With formatRange.FormatConditions(1).BarColor
        .Color = 6343505
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).BarFillType = xlDataBarFillSolid
    formatRange.FormatConditions(1).Direction = xlContext
    formatRange.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    formatRange.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
    formatRange.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With formatRange.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With formatRange.FormatConditions(1).NegativeBarFormat.Color
        .Color = 8420607
        .TintAndShade = 0
    End With
    
End Sub
    

' Format Border
Sub MakeBorder(formatRange As Range)
    formatRange.Borders(xlDiagonalDown).LineStyle = xlNone
    formatRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With formatRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    formatRange.Borders(xlEdgeTop).LineStyle = xlNone
    formatRange.Borders(xlEdgeBottom).LineStyle = xlNone
    With formatRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    formatRange.Borders(xlInsideVertical).LineStyle = xlNone
    formatRange.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

' Customise Header
Sub MakeHeader(formatRange As Range)
    With formatRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With formatRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With formatRange.Font
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984740745262
    End With
    formatRange.Font.Bold = True
End Sub








