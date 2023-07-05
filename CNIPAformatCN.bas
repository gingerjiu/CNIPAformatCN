Attribute VB_Name = "Module3"
Sub FormatCN()
Attribute FormatCN.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatCN Macro
'

'
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Range("F1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "SpecsClient"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Nice"
    Columns("B:B").Select
    Columns("A:A").ColumnWidth = 8
    Selection.ColumnWidth = 13.29
    Columns("D:D").ColumnWidth = 8
    Columns("F:F").ColumnWidth = 8
    Columns("E:E").ColumnWidth = 35.29
    Columns("G:G").ColumnWidth = 10
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
End Sub
