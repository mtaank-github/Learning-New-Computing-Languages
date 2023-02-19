Attribute VB_Name = "Module1"
Sub FormatTable()
Attribute FormatTable.VB_Description = "format table"
Attribute FormatTable.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' FormatTable Macro
' format table
'
' Keyboard Shortcut: Ctrl+j
'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Emp ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Last Name"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "First Name"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Dept"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Email"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Ext"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Location"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Hire Date"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Pay Rate"
    Range("A1:I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Size = 11
    Selection.Font.Size = 12
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Size = 14
    Selection.AutoFilter
    Range("J8").Select
End Sub
