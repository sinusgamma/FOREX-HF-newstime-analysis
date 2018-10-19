Attribute VB_Name = "Module8"
Sub AAB_ReportSort()
Attribute AAB_ReportSort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AAB_ReportSort Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets(2).Select
    Sheets(2).Name = "ShortS"
    Sheets.Add After:=ActiveSheet
    Sheets(3).Select
    Sheets(3).Name = "LongS"
    
    Sheets("Data").Select
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "AE1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range("A:BB").AutoFilter Field:=31, Criteria1:=">=1005" _
        , Operator:=xlAnd
    ActiveSheet.Range("A:BB").AutoFilter Field:=46, Criteria1:=">=0.45" _
        , Operator:=xlAnd
    ActiveSheet.Range("A:BB").AutoFilter Field:=54, Criteria1:=">=1.3" _
        , Operator:=xlAnd
    ActiveSheet.Range("A:BB").AutoFilter Field:=9, Criteria1:="SHORT"
    
    'True: copy header
    TopXFilteredCopy 10, True
    Sheets("ShortS").Select
    ActiveSheet.Paste
    
    Sheets("Data").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "AS:AS"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom '        .SortMethod = xlPinYin
        .Apply
    End With
    
    TopXFilteredCopy 5, False
    Sheets("ShortS").Select
    Range("A12").Select
    ActiveSheet.Paste
    
    Sheets("Data").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "BB:BB"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    TopXFilteredCopy 5, False

    Sheets("ShortS").Select
    Range("A17").Select
    ActiveSheet.Paste
    
    Sheets("Data").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=9, Criteria1:="LONG"
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "AE:AE"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'True: copy header
    TopXFilteredCopy 10, True
    Sheets("LongS").Select
    ActiveSheet.Paste
    
    Sheets("Data").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "AT:AT"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    TopXFilteredCopy 5, False
    Sheets("LongS").Select
    Range("A12").Select
    ActiveSheet.Paste

    Sheets("Data").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "BB:BB"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    TopXFilteredCopy 5, False
    Sheets("LongS").Select
    Range("A17").Select
    ActiveSheet.Paste
    
    Sheets("Data").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=9
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add key:=Range( _
        "AE:AE"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AE2").Select
    ActiveSheet.Range("A:BB").AutoFilter Field:=31
    ActiveSheet.Range("A:BB").AutoFilter Field:=46
    ActiveSheet.Range("A:BB").AutoFilter Field:=54
End Sub

Sub TopXFilteredCopy(x As Long, withHeader As Boolean)

Dim r As Range, rC As Range
Dim j As Long
Dim LastRow As Long
Dim HeaderSet As Long
Dim FirstRange As String

If withHeader = True Then
    HeaderSet = 1
    FirstRange = "F1"
Else
    HeaderSet = 0
    FirstRange = "F2"
End If

Set r = Nothing
Set rC = Nothing
j = 0


Set r = Range(FirstRange, Range("F" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible)

For Each rC In r
    j = j + 1
    If j = (x + HeaderSet) Or j = r.Count Then Exit For
Next rC

LastRow = rC(rC.Count).Row

Range(r(1), "BB" & LastRow).SpecialCells(xlCellTypeVisible).Copy

End Sub
