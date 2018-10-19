Attribute VB_Name = "Module9"
Sub AAC_ReportTablePoly()
Attribute AAC_ReportTablePoly.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AAC_ReportTablePoly Macro
'

'
    Sheets("ShortS").Select
    MakeTable "TableShort"
    ExtendSelection "TableShort"
    Interpolate
    Uptotwenty
    AddFormula
    Calc_Polynom "TableShort"
'
    Sheets("LongS").Select
    MakeTable "TableLong"
    ExtendSelection "TableLong"
    Interpolate
    Uptotwenty
    AddFormula
    Calc_Polynom "TableLong"

End Sub
Sub MakeTable(tableName As String)
    Range("A1:AP30").Select
    Range("AP1").Activate
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$AP$30"), , xlYes).Name = _
        tableName
    Range(tableName & "[#All]").Select
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight9"
    ActiveSheet.Range(tableName & "[#All]").RemoveDuplicates Columns:=Array(5, 12, 19), Header:=xlYes
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
End Sub
Sub Uptotwenty()
Attribute Uptotwenty.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Uptotwenty Macro
'

'
    Range("D23").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("D24").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("D23:D24").Select
    Selection.AutoFill Destination:=Range("D23:D42"), Type:=xlFillDefault
    Range("D23:D42").Select
End Sub
Sub AddFormula()
Attribute AddFormula.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AddFormula Macro
'

'
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Dim RowNb As Long
    RowNb = Selection.Rows.Count


    Range("E44").FormulaArray = "=INDEX(LINEST(E2:E" & (2 + RowNb - 1) & ",$D23:$D" & (23 + RowNb - 1) & "^{1,2}),1)"
    Range("E45").FormulaArray = "=INDEX(LINEST(E2:E" & (2 + RowNb - 1) & ",$D23:$D" & (23 + RowNb - 1) & "^{1,2}),1,2)"
    Range("E46").FormulaArray = "=INDEX(LINEST(E2:E" & (2 + RowNb - 1) & ",$D23:$D" & (23 + RowNb - 1) & "^{1,2}),1,3)"
    Range("E44:E46").Select
    Selection.AutoFill Destination:=Range("E44:M46"), Type:=xlFillDefault
    Range("E44:M46").Select

End Sub
Sub Calc_Polynom(tableName As String)
Attribute Calc_Polynom.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Calc_Polynom Macro
'

'
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[pointsAway]]"), SortOn:=xlSortOnValues, Order _
        :=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[pointsAway]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E44:E46").Select
    Selection.Copy
    Range("E48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[takeProfit]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F44:F46").Select
    Selection.Copy
    Range("F48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[stopLoss]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G44:G46").Select
    Selection.Copy
    Range("G48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H44:H50").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("I44").Select
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[breakevenTrigger]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("J44:J46").Select
    Selection.Copy
    Range("J48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[breakevenDistance]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("K2").Select
    Range("K44:K46").Select
    Selection.Copy
    Range("K48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[trailingStop]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("L44:L46").Select
    Selection.Copy
    Range("L48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[trailingAfter]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("M44:M46").Select
    Selection.Copy
    Range("M48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("E23").Select
    ActiveCell.FormulaR1C1 = "=(R48C * RC4^2) + (R49C * RC4) + R50C"
    Range("E23").Select
    Selection.AutoFill Destination:=Range("E23:E42"), Type:=xlFillDefault
    Range("E23:E42").Select
    Selection.AutoFill Destination:=Range("E23:M42"), Type:=xlFillDefault
End Sub
Sub ExtendSelection(tableName As String)
'
' ExtendSelection Macro
'
    
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Clear
    ActiveSheet.ListObjects(tableName).Sort.SortFields.Add _
        key:=Range(tableName & "[[#All],[pointsAway]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Dim RowNb As Long
    RowNb = Selection.Rows.Count
    
    If RowNb = 20 Then
        Return
    End If
    
    Dim StepToFill As Double
    Dim NeedRowNb As Long
       
    'get the steps to insert rows to get 20 sample
    NeedRowNb = 20 - RowNb
    FillStep = RowNb / (NeedRowNb + 1)
    
    Dim i As Integer
    Dim InsertPlace As Integer
        
    For i = NeedRowNb To 1 Step -1
        InsertPlace = CInt(i * FillStep) + 2
        Rows(InsertPlace & ":" & InsertPlace).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
End Sub

Sub Interpolate()

'The numbers
For i = 5 To 13
    'skip the time
    If i <> 8 Then
        For Each cell In Range(Cells(3, i), Cells(20, i))
            If cell.Value = "" Then
                y2 = cell.End(xlDown).Value
                y1 = cell.End(xlUp).Value
                y = (y1 + y2) / 2
                cell.Value = y
            End If
        Next
        
     End If
    
Next i

For k = 1 To 4
            For Each cell In Range(Cells(3, k), Cells(21, k))
            If cell.Value = "" Then
                y = cell.End(xlUp).Value
                cell.Value = y
            End If
        Next
Next k
End Sub
Sub AAD_ReportFinalPoly()
Attribute AAD_ReportFinalPoly.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AAD_ReportFinalPoly Macro
'
    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    Sheets(4).Name = "F_SHORT"
    
    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    Sheets(5).Name = "F_LONG"
    
    CopyData
    CopyPolynomData
    
End Sub
Sub CopyData()
Attribute CopyData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopyData Macro
'

'
    Sheets("ShortS").Select
    Range("TableShort[#All]").Select
    Range("TableShort[[#Headers],[PLrateCom]]").Activate
    Selection.Copy
    Sheets("F_SHORT").Select
    ActiveSheet.Paste
    With ActiveSheet
        .ListObjects(1).Name = "TableShortNew"
    End With
    
    Sheets("LongS").Select
    Range("TableLong[#All]").Select
    Range("TableLong[[#Headers],[PLrateCom]]").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F_LONG").Select
    ActiveSheet.Paste
    With ActiveSheet
        .ListObjects(1).Name = "TableLongNew"
    End With
End Sub

Sub CopyPolynomData()
Attribute CopyPolynomData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopyPolynomData Macro
'

'
    Sheets("F_SHORT").Select
    Range("E29").Select
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[pointsAway]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("E23:E42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[pointsAway]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[takeProfit]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("F23:F42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[takeProfit]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[stopLoss]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("G23:G42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[stopLoss]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[breakevenTrigger]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("J23:J42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[breakevenTrigger]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[breakevenDistance]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("K23:K42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[breakevenDistance]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[trailingStop]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("L23:L42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[trailingStop]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort. _
        SortFields.Add key:=Range("TableShortNew[[#All],[trailingAfter]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_SHORT").ListObjects("TableShortNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ShortS").Select
    Range("M23:M42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("TableShortNew[trailingAfter]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("F_LONG").Select
    Range("E24").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[pointsAway]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("E23:E42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[pointsAway]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[takeProfit]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F2:F22").Select
    Sheets("LongS").Select
    Range("F23:F42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[takeProfit]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[stopLoss]]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("G23:G42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[stopLoss]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[breakevenTrigger]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("J23:J42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[breakevenTrigger]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[breakevenDistance]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("K23:K42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[breakevenDistance]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[trailingStop]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("L23:L42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[trailingStop]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort.SortFields _
        .Add key:=Range("TableLongNew[[#All],[trailingAfter]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F_LONG").ListObjects("TableLongNew").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("LongS").Select
    Range("M23:M42").Select
    Selection.Copy
    Sheets("F_LONG").Select
    Range("TableLongNew[trailingAfter]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
