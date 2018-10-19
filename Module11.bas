Attribute VB_Name = "Module11"
Sub tester()
Attribute tester.VB_ProcData.VB_Invoke_Func = " \n14"
'
' tester Macro
'

'
    Rows("1:33").Select
    Range("F1").Activate
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
End Sub
