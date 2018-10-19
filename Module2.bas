Attribute VB_Name = "Module2"
Sub a3_CountCorrelation()

Dim LastRow As Integer
Dim LastColumnNbr As Integer
Dim LastColumnLtr As String

Dim i As Integer
Dim j As Integer

Dim chTitle As String
Dim series1 As String
Dim series2 As String
Dim series3 As String
Dim series4 As String
Dim series5 As String
Dim series6 As String
Dim xth As Integer

Dim ActColumnLtrC As String
Dim ActColumnLtrCn As String 'n - next
Dim ActColumnLtrD As String
Dim ActColumnLtrDn As String


    ' Got to the Correl sheet
    Sheets("Correl").Activate

    ' Sheet limits on Derived sheet
    LastRow = Sheets("Derived").UsedRange.Rows.Count 'Number of Rows in Derived sheet
    LastColumnNbr = Sheets("Derived").UsedRange.Columns.Count 'Number of columns in Derived sheet
    LastColumnLtr = ConvertToLetter(LastColumnNbr)
    
    ' Make Header
    Sheets("Derived").Range("N1:" & LastColumnLtr & "1").Copy Destination:=Sheets("Correl").Range("B1")
    
    ' Make Row names
    Sheets("Correl").Range("A2").Value = "actual"
    Sheets("Correl").Range("A3").Value = "forecast"
    Sheets("Correl").Range("A4").Value = "prev"
    Sheets("Correl").Range("A5").Value = "mod_from"
    Sheets("Correl").Range("A6").Value = "Act-For"
    Sheets("Correl").Range("A7").Value = "Act-Pre"
    Sheets("Correl").Range("A8").Value = "Pre-MFr"
    
    'Format first column
    FormatFirstColumn (Sheets("Correl").Range("A2:A8"))
    
    ' Sheet limits on Correl sheet after first column and header
    LastRowHere = Sheets("Correl").UsedRange.Rows.Count
    LastColumnNbrHere = Sheets("Correl").UsedRange.Columns.Count
    LastColumnLtrHere = ConvertToLetter(LastColumnNbr)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Count Correlations from Derived sheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Count correlations ACTUAL
    FString = "=CORREL(Derived!$F2:$F" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B2").Formula = FString
    'Count for row
    Sheets("Correl").Range("B2").AutoFill Destination:=Sheets("Correl").Range("B2:" & LastColumnLtrHere & 2)
    
    'Count correlations FORECAST
    FString = "=CORREL(Derived!$G2:$G" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B3").Formula = FString
    'Count for row
    Sheets("Correl").Range("B3").AutoFill Destination:=Sheets("Correl").Range("B3:" & LastColumnLtrHere & 3)
    
    'Count correlations PREVIEW
    FString = "=CORREL(Derived!$H2:$H" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B4").Formula = FString
    'Count for row
    Sheets("Correl").Range("B4").AutoFill Destination:=Sheets("Correl").Range("B4:" & LastColumnLtrHere & 4)
    
    'Count correlations MOD_FROM
    FString = "=CORREL(Derived!$I2:$I" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B5").Formula = FString
    'Count for row
    Sheets("Correl").Range("B5").AutoFill Destination:=Sheets("Correl").Range("B5:" & LastColumnLtrHere & 5)
    
    'Count correlations ACT-FOR
    FString = "=CORREL(Derived!$K2:$K" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B6").Formula = FString
    'Count for row
    Sheets("Correl").Range("B6").AutoFill Destination:=Sheets("Correl").Range("B6:" & LastColumnLtrHere & 6)
    
    'Count correlations ACT-PRE
    FString = "=CORREL(Derived!$L2:$L" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B7").Formula = FString
    'Count for row
    Sheets("Correl").Range("B7").AutoFill Destination:=Sheets("Correl").Range("B7:" & LastColumnLtrHere & 7)
    
    'Count correlations PRE-MFr
    FString = "=CORREL(Derived!$M2:$M" & LastRow & ",Derived!N2:N" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B8").Formula = FString
    'Count for row
    Sheets("Correl").Range("B8").AutoFill Destination:=Sheets("Correl").Range("B8:" & LastColumnLtrHere & 8)

    
    ' Conditional Format Correlation
    ConditionalFormatCorrelation (Sheets("Correl").Range("B2:" & LastColumnLtrHere & 8))
    
    'Delet formula if error is given
    DeletFormulaError (Sheets("Correl").Range("B2:" & LastColumnLtrHere & 8))
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Count maxcor - maxpredictor - R2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Make Row Names
    Sheets("Correl").Range("A10").Value = "AbsMaxCor"
    Sheets("Correl").Range("A11").Value = "Cor-With"
    Sheets("Correl").Range("A12").Value = "R2"
    
    'Count Maximum Correlation Absolute Value
    FString = "=MAX(ABS(B2:B8))"
    'Count for first cell
    Sheets("Correl").Range("B10").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B10").AutoFill Destination:=Sheets("Correl").Range("B10:" & LastColumnLtrHere & 10)
    
    'Conditional Format Absolute correlation Row
    ConditionalFormatAbsCor (Sheets("Correl").Range("B10:" & LastColumnLtrHere & 10))
    
    'Find the name of maximum correlation datasource
    FString = "=IF(ABS(MAX(B2:B8))>ABS(MIN(B2:B8)),INDEX($A2:$A8,MATCH(MAX(B2:B8),B2:B8,0)),INDEX($A2:$A8,MATCH(MIN(B2:B8),B2:B8,0)))"
    'Count for first cell
    Sheets("Correl").Range("B11").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B11").AutoFill Destination:=Sheets("Correl").Range("B11:" & LastColumnLtrHere & 11)
    
    'Count R2
    FString = "=POWER(B10,2)"
    'Count for first cell
    Sheets("Correl").Range("B12").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B12").AutoFill Destination:=Sheets("Correl").Range("B12:" & LastColumnLtrHere & 12)
    ' Conditional format R2
    ConditionalFormatRsq (Sheets("Correl").Range("B12:" & LastColumnLtrHere & 12))
 
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Count scatter plot quarter percentage
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Make Row Names
    Sheets("Correl").Range("A14").Value = "P(x>0|y>0)"
    Sheets("Correl").Range("A15").Value = "P(x<0|y>0)"
    Sheets("Correl").Range("A16").Value = "P(x>0|y<0)"
    Sheets("Correl").Range("A17").Value = "P(x<0|y<0)"
    
    'Count scatter plot upper right
    FString = "=COUNTIFS(Derived!N2:N" & LastRow & ","">0"",Derived!$K2:$K" & LastRow & ","">0"")/COUNTIFS(Derived!$K2:$K" & LastRow & ","">0"",Derived!N2:N" & LastRow & ",""<>""&""*"")"
    'Count for first cell
    Sheets("Correl").Range("B14").Formula = FString
    'Count for row
    Sheets("Correl").Range("B14").AutoFill Destination:=Sheets("Correl").Range("B14:" & LastColumnLtrHere & 14)
    
    'Count scatter plot lower right
    FString = "=COUNTIFS(Derived!N2:N" & LastRow & ",""<0"",Derived!$K2:$K" & LastRow & ","">0"")/COUNTIFS(Derived!$K2:$K" & LastRow & ","">0"",Derived!N2:N" & LastRow & ",""<>""&""*"")"
    'Count for first cell
    Sheets("Correl").Range("B15").Formula = FString
    'Count for row
    Sheets("Correl").Range("B15").AutoFill Destination:=Sheets("Correl").Range("B15:" & LastColumnLtrHere & 15)
    
    'Count scatter plot upper left
    FString = "=COUNTIFS(Derived!N2:N" & LastRow & ","">0"",Derived!$K2:$K" & LastRow & ",""<0"")/COUNTIFS(Derived!$K2:$K" & LastRow & ",""<0"",Derived!N2:N" & LastRow & ",""<>""&""*"")"
    'Count for first cell
    Sheets("Correl").Range("B16").Formula = FString
    'Count for row
    Sheets("Correl").Range("B16").AutoFill Destination:=Sheets("Correl").Range("B16:" & LastColumnLtrHere & 16)
    
    'Count scatter plot lower left
    FString = "=COUNTIFS(Derived!N2:N" & LastRow & ",""<0"",Derived!$K2:$K" & LastRow & ",""<0"")/COUNTIFS(Derived!$K2:$K" & LastRow & ",""<0"",Derived!N2:N" & LastRow & ",""<>""&""*"")"
    'Count for first cell
    Sheets("Correl").Range("B17").Formula = FString
    'Count for row
    Sheets("Correl").Range("B17").AutoFill Destination:=Sheets("Correl").Range("B17:" & LastColumnLtrHere & 17)
    
    'Conditional format quarters
    ConditionalFormatQuarters (Sheets("Correl").Range("B14:" & LastColumnLtrHere & 17))
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Summary of A-F percents
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Name Data
    Sheets("Correl").Range("A19").Value = "%AF>0"
    Sheets("Correl").Range("C19").Value = "%AF<0"
    Sheets("Correl").Range("E19").Value = "%AF=0"
    Sheets("Correl").Range("G19").Value = "%AF Blank"
    Sheets("Correl").Range("I19").Value = "SUM"

    'Count %AF>0
    FString = "=COUNTIF(Derived!K2:K" & LastRow & ","">0"")/" & LastRow - 1 & "*100"
    'Count for first cell
    Sheets("Correl").Range("B19").Formula = FString
    
    'Count %AF<0
    FString = "=COUNTIF(Derived!K2:K" & LastRow & ",""<0"")/" & LastRow - 1 & "*100"
    'Count for first cell
    Sheets("Correl").Range("D19").Formula = FString
    
    'Count %AF=0
    FString = "=COUNTIF(Derived!K2:K" & LastRow & ",""=0"")/" & LastRow - 1 & "*100"
    'Count for first cell
    Sheets("Correl").Range("F19").Formula = FString
    
    'Count %AF Blank
    FString = "=COUNTBLANK(Derived!K2:K" & LastRow & ")/" & LastRow - 1 & "*100"
    'Count for first cell
    Sheets("Correl").Range("H19").Formula = FString
    
    'Check above %
    Sheets("Correl").Range("J19").Formula = "=B19+D19+F19+H19"
    
    ConditionalFormatPercents (Sheets("Correl").Range("A19:J19"))
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Count Averages, Medians, Quartiles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''
    'BLOCK A-F > 0
    
    'Count Q1 A-f>0
    Sheets("Correl").Range("A21").Value = "Q1 A-f>0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & ">0,Derived!N2:N" & LastRow & "),1)"
    'Count for first cell
    Sheets("Correl").Range("B21").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B21").AutoFill Destination:=Sheets("Correl").Range("B21:" & LastColumnLtrHere & 21)

    'Count Avg A-f>0
    Sheets("Correl").Range("A22").Value = "Avg A-f>0"
    FString = "=AVERAGEIFS(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ","">0"")"
    'Count for first cell
    Sheets("Correl").Range("B22").Formula = FString
    'Count for row
    Sheets("Correl").Range("B22").AutoFill Destination:=Sheets("Correl").Range("B22:" & LastColumnLtrHere & 22)

    'Count Med A-f>0
    Sheets("Correl").Range("A23").Value = "Med A-f>0"
    FString = "=MEDIAN(IF(Derived!$K2:$K" & LastRow & ">0,Derived!N2:N" & LastRow & "))"
    'Count for first cell
    Sheets("Correl").Range("B23").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B23").AutoFill Destination:=Sheets("Correl").Range("B23:" & LastColumnLtrHere & 23)

    'Count Q3 A-f>0
    Sheets("Correl").Range("A24").Value = "Q3 A-f>0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & ">0,Derived!N2:N" & LastRow & "),3)"
    'Count for first cell
    Sheets("Correl").Range("B24").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B24").AutoFill Destination:=Sheets("Correl").Range("B24:" & LastColumnLtrHere & 24)
    
    ''''''''''''''''''
    'BLOCK A-F < 0
    
    'Count Q1 A-f<0
    Sheets("Correl").Range("A26").Value = "Q1 A-f<0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "<0,Derived!N2:N" & LastRow & "),1)"
    'Count for first cell
    Sheets("Correl").Range("B26").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B26").AutoFill Destination:=Sheets("Correl").Range("B26:" & LastColumnLtrHere & 26)
    
    'Count Avg A-f<0
    Sheets("Correl").Range("A27").Value = "Avg A-f<0"
    FString = "=AVERAGEIFS(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ",""<0"")"
    'Count for first cell
    Sheets("Correl").Range("B27").Formula = FString
    'Count for row
    Sheets("Correl").Range("B27").AutoFill Destination:=Sheets("Correl").Range("B27:" & LastColumnLtrHere & 27)

    'Count Med A-f<0
    Sheets("Correl").Range("A28").Value = "Med A-f<0"
    FString = "=MEDIAN(IF(Derived!$K2:$K" & LastRow & "<0,Derived!N2:N" & LastRow & "))"
    'Count for first cell
    Sheets("Correl").Range("B28").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B28").AutoFill Destination:=Sheets("Correl").Range("B28:" & LastColumnLtrHere & 28)

    'Count Q3 A-f<0
    Sheets("Correl").Range("A29").Value = "Q3 A-f<0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "<0,Derived!N2:N" & LastRow & "),3)"
    'Count for first cell
    Sheets("Correl").Range("B29").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B29").AutoFill Destination:=Sheets("Correl").Range("B29:" & LastColumnLtrHere & 29)


    ''''''''''''''''''
    'BLOCK A-F = 0
    
    'Count Q1 A-f=0
    Sheets("Correl").Range("A31").Value = "Q1 A-f=0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "=0,Derived!N2:N" & LastRow & "),1)"
    'Count for first cell
    Sheets("Correl").Range("B31").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B31").AutoFill Destination:=Sheets("Correl").Range("B31:" & LastColumnLtrHere & 31)

    'Count Avg A-f=0
    Sheets("Correl").Range("A32").Value = "Avg A-f=0"
    FString = "=AVERAGEIFS(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ",""=0"")"
    'Count for first cell
    Sheets("Correl").Range("B32").Formula = FString
    'Count for row
    Sheets("Correl").Range("B32").AutoFill Destination:=Sheets("Correl").Range("B32:" & LastColumnLtrHere & 32)

    'Count Med A-f=0
    Sheets("Correl").Range("A33").Value = "Med A-f=0"
    FString = "=MEDIAN(IF(Derived!$K2:$K" & LastRow & "=0,Derived!N2:N" & LastRow & "))"
    'Count for first cell
    Sheets("Correl").Range("B33").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B33").AutoFill Destination:=Sheets("Correl").Range("B33:" & LastColumnLtrHere & 33)
    
    'Count Q3 A-f=0
    Sheets("Correl").Range("A34").Value = "Q3 A-f=0"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "=0,Derived!N2:N" & LastRow & "),3)"
    'Count for first cell
    Sheets("Correl").Range("B34").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B34").AutoFill Destination:=Sheets("Correl").Range("B34:" & LastColumnLtrHere & 34)
    
    
    ''''''''''''''''''
    'BLOCK A-F ISBLANK
    
    'Count Q1 BLANK
    Sheets("Correl").Range("A36").Value = "Q1 Blank"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "="""",Derived!N2:N" & LastRow & "),1)"
    'Count for first cell
    Sheets("Correl").Range("B36").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B36").AutoFill Destination:=Sheets("Correl").Range("B36:" & LastColumnLtrHere & 36)
    
    'Count Avg BLANK
    Sheets("Correl").Range("A37").Value = "Avg Blank"
    FString = "=AVERAGEIFS(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ","""")"
    'Count for first cell
    Sheets("Correl").Range("B37").Formula = FString
    'Count for row
    Sheets("Correl").Range("B37").AutoFill Destination:=Sheets("Correl").Range("B37:" & LastColumnLtrHere & 37)

    'Count Med BLANK
    Sheets("Correl").Range("A38").Value = "Med Blank"
    FString = "=MEDIAN(IF(Derived!$K2:$K" & LastRow & "="""",Derived!N2:N" & LastRow & "))"
    'Count for first cell
    Sheets("Correl").Range("B38").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B38").AutoFill Destination:=Sheets("Correl").Range("B38:" & LastColumnLtrHere & 38)

    'Count Q3 BLANK
    Sheets("Correl").Range("A39").Value = "Q3 Blank"
    FString = "=QUARTILE.EXC(IF(Derived!$K2:$K" & LastRow & "="""",Derived!N2:N" & LastRow & "),3)"
    'Count for first cell
    Sheets("Correl").Range("B39").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B39").AutoFill Destination:=Sheets("Correl").Range("B39:" & LastColumnLtrHere & 39)
    
    
    ''''''''''''''''''
    'BLOCK LARGER
    
    'Count Q1 of Larger Absolute Values
    Sheets("Correl").Range("A41").Value = "Q1AbsL"
    FString = "=QUARTILE.EXC(IF(ABS(Derived!N2:N" & LastRow & ")>ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")),1)"
    'Count for first cell
    Sheets("Correl").Range("B41").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B41").AutoFill Destination:=Sheets("Correl").Range("B41:" & LastColumnLtrHere & 41)
    
    'Count Average of Larger Absolute Values
    Sheets("Correl").Range("A42").Value = "AvgAbsL"
    FString = "=AVERAGE(IF(ABS(Derived!N2:N" & LastRow & ")>ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")))"
    'Count for first cell
    Sheets("Correl").Range("B42").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B42").AutoFill Destination:=Sheets("Correl").Range("B42:" & LastColumnLtrHere & 42)

    'Count Median of Larger Absolute Values
    Sheets("Correl").Range("A43").Value = "MedAbsL"
    FString = "=MEDIAN(IF(ABS(Derived!N2:N" & LastRow & ")>ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")))"
    'Count for first cell
    Sheets("Correl").Range("B43").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B43").AutoFill Destination:=Sheets("Correl").Range("B43:" & LastColumnLtrHere & 43)
    
    'Count Q3 of Larger Absolute Values
    Sheets("Correl").Range("A44").Value = "Q3AbsL"
    FString = "=QUARTILE.EXC(IF(ABS(Derived!N2:N" & LastRow & ")>ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")),3)"
    'Count for first cell
    Sheets("Correl").Range("B44").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B44").AutoFill Destination:=Sheets("Correl").Range("B44:" & LastColumnLtrHere & 44)
    
     
    ''''''''''''''''''
    'BLOCK SMALLER
    
    'Count Q1 of Smaller Absolute Values
    Sheets("Correl").Range("A46").Value = "Q1AbsS"
    FString = "=QUARTILE.EXC(IF(ABS(Derived!N2:N" & LastRow & ")<ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")),1)"
    'Count for first cell
    Sheets("Correl").Range("B46").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B46").AutoFill Destination:=Sheets("Correl").Range("B46:" & LastColumnLtrHere & 46)
    
    'Count Average of Smaller Absolute Values
    Sheets("Correl").Range("A47").Value = "AvgAbsS"
    FString = "=AVERAGE(IF(ABS(Derived!N2:N" & LastRow & ")<ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")))"
    'Count for first cell
    Sheets("Correl").Range("B47").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B47").AutoFill Destination:=Sheets("Correl").Range("B47:" & LastColumnLtrHere & 47)

    'Count Median of Smaller Absolute Values
    Sheets("Correl").Range("A48").Value = "MedAbsS"
    FString = "=MEDIAN(IF(ABS(Derived!N2:N" & LastRow & ")<ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")))"
    'Count for first cell
    Sheets("Correl").Range("B48").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B48").AutoFill Destination:=Sheets("Correl").Range("B48:" & LastColumnLtrHere & 48)
    
    'Count Q3 of Smaller Absolute Values
    Sheets("Correl").Range("A49").Value = "Q3AbsS"
    FString = "=QUARTILE.EXC(IF(ABS(Derived!N2:N" & LastRow & ")<ABS(Derived!O2:O" & LastRow & "),ABS(Derived!N2:N" & LastRow & "),ABS(Derived!O2:O" & LastRow & ")),3)"
    'Count for first cell
    Sheets("Correl").Range("B49").FormulaArray = FString
    'Count for row
    Sheets("Correl").Range("B49").AutoFill Destination:=Sheets("Correl").Range("B49:" & LastColumnLtrHere & 49)


    ''''''''''''''''''
    'BLOCK RATE

    'Count AvgAbsL / AvgAbsS
    Sheets("Correl").Range("A51").Value = "AvgL/s"
    FString = "=B42/B47"
    'Count for first cell
    Sheets("Correl").Range("B51").Formula = FString
    'Count for row
    Sheets("Correl").Range("B51").AutoFill Destination:=Sheets("Correl").Range("B51:" & LastColumnLtrHere & 51)

    'Count MedAbsL / MedAbsS
    Sheets("Correl").Range("A52").Value = "MedL/s"
    FString = "=B43/B48"
    'Count for first cell
    Sheets("Correl").Range("B52").Formula = FString
    'Count for row
    Sheets("Correl").Range("B52").AutoFill Destination:=Sheets("Correl").Range("B52:" & LastColumnLtrHere & 52)
    
    'Count LQ1-SQ3 to see the gap between the small larger and large smaller swing - some safety zone, larger = better
    Sheets("Correl").Range("A53").Value = "LQ1-SQ3"
    FString = "=B41-B49"
    'Count for first cell
    Sheets("Correl").Range("B53").Formula = FString
    'Count for row
    Sheets("Correl").Range("B53").AutoFill Destination:=Sheets("Correl").Range("B53:" & LastColumnLtrHere & 53)
    
    'Count (LQ1-SQ3)/LQ1
    Sheets("Correl").Range("A54").Value = "(LQ1-SQ3)/LQ1"
    FString = "=(B41-B49)/B41"
    'Count for first cell
    Sheets("Correl").Range("B54").Formula = FString
    'Count for row
    Sheets("Correl").Range("B54").AutoFill Destination:=Sheets("Correl").Range("B54:" & LastColumnLtrHere & 54)
    
    'Count (LQ1-SQ3)/SQ3
    Sheets("Correl").Range("A55").Value = "(LQ1-SQ3)/SQ3"
    FString = "=(B41-B49)/B49"
    'Count for first cell
    Sheets("Correl").Range("B55").Formula = FString
    'Count for row
    Sheets("Correl").Range("B55").AutoFill Destination:=Sheets("Correl").Range("B55:" & LastColumnLtrHere & 55)
    
    'Count LAvg-SQ3 some expected profig in pip - compare to spread later
    Sheets("Correl").Range("A56").Value = "LAvg-SQ3"
    FString = "=B42-B49"
    'Count for first cell
    Sheets("Correl").Range("B56").Formula = FString
    'Count for row
    Sheets("Correl").Range("B56").AutoFill Destination:=Sheets("Correl").Range("B56:" & LastColumnLtrHere & 56)
    
    'Count LMed-SQ3 some expected profig in pip - compare to spread later
    Sheets("Correl").Range("A57").Value = "LMed-SQ3"
    FString = "=B43-B49"
    'Count for first cell
    Sheets("Correl").Range("B57").Formula = FString
    'Count for row
    Sheets("Correl").Range("B57").AutoFill Destination:=Sheets("Correl").Range("B57:" & LastColumnLtrHere & 57)
    

    ' Clear every other column because they calculate false data for the larger, smaller absolutes
    For i = 3 To LastColumnNbr Step 2
        For j = 41 To 57
            Sheets("Correl").Cells(j, i).Clear
        Next j
    Next i

    ' Format Aboves
    ConditionalFormatAverages (Sheets("Correl").Range("B21:" & LastColumnLtrHere & 39)) ' Color of A M < > = 0 or blank
    ConditionalFormatRate (Sheets("Correl").Range("B51:" & LastColumnLtrHere & 52)) ' 'Color of AvgL/s MedL/s
    ConditionalFormatGapPerSQ3 (Sheets("Correl").Range("B55:" & LastColumnLtrHere & 55))
    Sheets("Correl").Range("B2:" & LastColumnLtrHere & 70).NumberFormat = "0.00000"
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Count Intercept and Slope
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Make Row Names
    Sheets("Correl").Range("A60").Value = "Intercept A-f"
    Sheets("Correl").Range("A61").Value = "Slope A-f"

    'Count Intercept A-f
    FString = "=INTERCEPT(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B60").Formula = FString
    'Count for row
    Sheets("Correl").Range("B60").AutoFill Destination:=Sheets("Correl").Range("B60:" & LastColumnLtrHere & 60)

    'Count Slope A-f
    FString = "=SLOPE(Derived!N2:N" & LastRow & ",Derived!$K2:$K" & LastRow & ")"
    'Count for first cell
    Sheets("Correl").Range("B61").Formula = FString
    'Count for row
    Sheets("Correl").Range("B61").AutoFill Destination:=Sheets("Correl").Range("B61:" & LastColumnLtrHere & 61)

    'Conditional Format Intercept the same as ConditionalFormatAverages
    ConditionalFormatAverages (Sheets("Correl").Range("B60:" & LastColumnLtrHere & 60))
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Generate Opposite Swing Columns
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Range("A64").Select
        
    'Smaller Swing All Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrC = ConvertToLetter(i) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        Range(ActColumnLtrC & "69").Value = Range(ActColumnLtrC & "1")
        
        FString = "=MIN(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2))"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrC & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrC & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrC & "70:" & ActColumnLtrC & 70 + LastRow - 2)
    Next i
    
    'Larger Swing All Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrCn = ConvertToLetter(i + 1) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        FString = "=MAX(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2))"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrCn & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrCn & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrCn & "70:" & ActColumnLtrCn & 70 + LastRow - 2)
    Next i
    
    'Smaller Swing if Larger is Negativ Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrC = ConvertToLetter(i + 2) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        Range(ActColumnLtrC & "69").Value = "SmNegSw"
        
        FString = "=IF(ABS(Derived!" & ActColumnLtrD & "2) > ABS(Derived!" & ActColumnLtrDn & "2), MIN(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2)),NA())"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrC & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrC & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrC & "70:" & ActColumnLtrC & 70 + LastRow - 2)
    Next i
    
    'Larger Swing if Larger is Negativ Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrC = ConvertToLetter(i + 3) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        Range(ActColumnLtrC & "69").Value = "LrNegSw"
        
        FString = "=IF(ABS(Derived!" & ActColumnLtrD & "2) > ABS(Derived!" & ActColumnLtrDn & "2), MAX(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2)),NA())"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrC & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrC & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrC & "70:" & ActColumnLtrC & 70 + LastRow - 2)
    Next i
    
    'Smaller Swing if Larger is Positive Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrC = ConvertToLetter(i + 4) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        Range(ActColumnLtrC & "69").Value = "SmPosSw"
        
        FString = "=IF(ABS(Derived!" & ActColumnLtrD & "2) <= ABS(Derived!" & ActColumnLtrDn & "2), MIN(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2)),NA())"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrC & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrC & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrC & "70:" & ActColumnLtrC & 70 + LastRow - 2)
    Next i
    
    'Larger Swing if Larger is Positiv Column from 1min
    For i = 2 To LastColumnNbrHere Step 8
        
        ActColumnLtrC = ConvertToLetter(i + 5) 'the column here
        ActColumnLtrD = ConvertToLetter(i + 12) 'the column on derived
        ActColumnLtrDn = ConvertToLetter(i + 13) 'the column on derived
        
        Range(ActColumnLtrC & "69").Value = "LrPosSw"
        
        FString = "=IF(ABS(Derived!" & ActColumnLtrD & "2) <= ABS(Derived!" & ActColumnLtrDn & "2), MAX(ABS(Derived!" & ActColumnLtrD & "2),ABS(Derived!" & ActColumnLtrDn & "2)),NA())"

        'Count for first cell
        Sheets("Correl").Range(ActColumnLtrC & "70").Formula = FString
        'Count for row
        Sheets("Correl").Range(ActColumnLtrC & "70").AutoFill Destination:=Sheets("Correl").Range(ActColumnLtrC & "70:" & ActColumnLtrC & 70 + LastRow - 2)
    Next i
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Generate Box-Plots
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets("Correl")

    LastRowHere = Sheets("Correl").UsedRange.Rows.Count ' count again, the table is larger

        For i = 2 To LastColumnNbrHere Step 8
        
        'very ugly, but no time
        
        chTitle = Range(ConvertToLetter(i) & "1").Value
        series1 = "=Correl!$" & ConvertToLetter(i) & "$70:$" & ConvertToLetter(i) & "$" & LastRowHere
        series2 = "=Correl!$" & ConvertToLetter(i + 1) & "$70:$" & ConvertToLetter(i + 1) & "$" & LastRowHere
        series3 = "=Correl!$" & ConvertToLetter(i + 2) & "$70:$" & ConvertToLetter(i + 2) & "$" & LastRowHere
        series4 = "=Correl!$" & ConvertToLetter(i + 3) & "$70:$" & ConvertToLetter(i + 3) & "$" & LastRowHere
        series5 = "=Correl!$" & ConvertToLetter(i + 4) & "$70:$" & ConvertToLetter(i + 4) & "$" & LastRowHere
        series6 = "=Correl!$" & ConvertToLetter(i + 5) & "$70:$" & ConvertToLetter(i + 5) & "$" & LastRowHere

        Call BoxPlotAbs(sh, chTitle, series1, series2, series3, series4, series5, series6, i)

        Next i
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Freeze Panes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Range("B2").Select
    ActiveWindow.FreezePanes = True

End Sub

' Converts Column Number to Letter
Function ConvertToLetter(ColNum As Integer) As String
    ConvertToLetter = Split(Cells(1, ColNum).Address, "$")(1)
End Function

'Formats the first column
Sub FormatFirstColumn(formatRange As Range)
    With formatRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With formatRange.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    formatRange.Font.Bold = True
End Sub

'Format the correlation table
Sub ConditionalFormatCorrelation(formatRange As Range)
    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Value = -0.9
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 255
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
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Value = 0.9
    With formatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 13369344
        .TintAndShade = 0
    End With
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.5"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=-0.5"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
End Sub

' Format Absolute Max Correlation Row
Sub ConditionalFormatAbsCor(formatRange As Range)
    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Value = 0
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Value = 0.4
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Value = 1#
    With formatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 10498160
        .TintAndShade = 0
    End With
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.6"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With formatRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
End Sub

' Conditional format R2
Sub ConditionalFormatRsq(formatRange As Range)

    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.2"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.3"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.4"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.5"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399945066682943
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.6"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
End Sub

'Conditional format quarter percents
Sub ConditionalFormatQuarters(formatRange As Range)

    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.7"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.8"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.9"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399945066682943
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
End Sub

'Conditional format averages
Sub ConditionalFormatAverages(formatRange As Range)

    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
    formatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    With formatRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).StopIfTrue = False
End Sub

' Conditional format L/S rate
Sub ConditionalFormatRate(formatRange As Range)

    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Value = 0
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Value = 2
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Value = 12
    With formatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
End Sub

'Conditional Format Data percents
Sub ConditionalFormatPercents(formatRange As Range)
    formatRange.FormatConditions.AddColorScale ColorScaleType:=2
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Value = 0
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Value = 80
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 255
        .TintAndShade = 0
    End With
    
    formatRange.Borders(xlDiagonalDown).LineStyle = xlNone
    formatRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With formatRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With formatRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With formatRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With formatRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    formatRange.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

' Conditional Format Gap/PQ3
Sub ConditionalFormatGapPerSQ3(formatRange As Range)

    formatRange.FormatConditions.AddColorScale ColorScaleType:=3
    formatRange.FormatConditions(formatRange.FormatConditions.Count).SetFirstPriority
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(1).Value = -2
    With formatRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 255
        .TintAndShade = 0
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With formatRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
    End With
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    formatRange.FormatConditions(1).ColorScaleCriteria(3).Value = 10
    With formatRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 5287936
        .TintAndShade = 0
    End With
End Sub

Sub BoxPlotAbs(shLocal As Worksheet, chTitle As String, series1 As String, series2 As String, series3 As String, series4 As String, series5 As String, series6 As String, i As Integer)

    Dim boxChrt As Shape
                  
    Set boxChrt = shLocal.Shapes.AddChart2(408, xlBoxwhisker)
    boxChrt.Select
    
    Do While ActiveChart.SeriesCollection.Count > 0 ' !!!!!! Prevents unwanted extra data :)
        ActiveChart.SeriesCollection(1).Delete
    Loop
   
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Values = series1

    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Values = series2
    
        ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Values = series3
    
        ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Values = series4
    
        ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(5).Values = series5
    
        ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(6).Values = series6
    
    ActiveChart.ChartTitle.Text = chTitle
    ActiveChart.ChartTitle.Characters.Font.Size = 10
    ActiveChart.HasLegend = False
    ActiveChart.SetElement (330) ' the horizontal grids
    ActiveChart.SetElement (331)
    ActiveChart.Parent.Left = 45 + ((i - 2) * 49)
    ActiveChart.Parent.Top = 935
    ActiveChart.ChartArea.Width = 380
    ActiveChart.ChartArea.Height = 400
    
    ' ActiveChart.DisplayBlanksAs = xlNotPlotted - doesn't work
    
End Sub

' Delete formula from empty or error cells, giving back a really empty cell
Sub DeletFormulaError(rng As Range)

    Dim cell As Range
    For Each cell In rng
        If (IsError(cell)) Then
           cell.ClearContents
        End If
    Next cell
    
End Sub
