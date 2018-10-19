Attribute VB_Name = "Module1"
Sub a1_newsheets()
Attribute a1_newsheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' t_newsheet Macro
'

'
    Sheets(1).Name = "Data"
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Derived"
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Correl"

End Sub
