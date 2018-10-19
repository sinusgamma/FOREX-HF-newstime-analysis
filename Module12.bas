Attribute VB_Name = "Module12"
Sub AAC_ReportTableRandom()
'
' AAC_ReportTableRandom Macro
'

'
    Sheets("ShortS").Select
    MakeTable "TableShort"
    ExtendSelection "TableShort"
    Interpolate
    Calc_Random

    Sheets("LongS").Select
    MakeTable "TableLong"
    ExtendSelection "TableLong"
    Interpolate
    Calc_Random

End Sub

Sub MakeTable(tableName As String)
    Range("A1:AW30").Select
    Range("AW1").Activate
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$AW$30"), , xlYes).Name = _
        tableName
    Range(tableName & "[#All]").Select
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight9"
    ActiveSheet.Range(tableName & "[#All]").RemoveDuplicates Columns:=Array(5, 15, 26), Header:=xlYes
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
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

Dim isNumber As Boolean
Dim randNb As Long

For i = 5 To 24
        For Each cell In Range(Cells(3, i), Cells(20, i))
            If cell.Value = "" Then
                y2 = cell.End(xlDown).Value
                y1 = cell.End(xlUp).Value
                If i < 17 Or i > 20 Then
                    y = (y1 + y2) / 2
                    cell.Value = y
                Else
                    If Rnd > 0.5 Then
                        Z = y2
                    Else
                        Z = y1
                    End If
                    If Z = 0 Then
                        cell.Value = "FALSE"
                    Else
                        cell.Value = "TRUE"
                    End If
                End If
            End If
        Next
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
Sub Calc_Random()
'
' Calc_Random Macro
'
    Range("E23").Select
    ActiveCell.FormulaR1C1 = "=R[-21]C+(RAND()-0.5)*R[-21]C*0.25"
    Selection.AutoFill Destination:=Range("E23:E42"), Type:=xlFillDefault
    Range("E23:E42").Select
    Selection.AutoFill Destination:=Range("E23:P42"), Type:=xlFillDefault
    Range("E23:P42").Select
End Sub

Sub AAD_ReportFinalRandom()
'
' AAD_ReportFinalRandom Macro
'
    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    Sheets(4).Name = "F_SHORT"

    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    Sheets(5).Name = "F_LONG"
    
    CopyData
    CopyRandomData
    
End Sub
Sub CopyData()
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
Sub CopyRandomData()
Attribute CopyRandomData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopyRandomData Macro
'

'
    Sheets("ShortS").Select
    Range("E23:G42").Select
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("ShortS").Select
    Range("M23:P42").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F_SHORT").Select
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("LongS").Select
    Range("E23:G42").Select
    Range("E42").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F_LONG").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("LongS").Select
    Range("M23:P42").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F_LONG").Select
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
