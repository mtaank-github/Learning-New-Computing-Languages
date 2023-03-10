Attribute VB_Name = "Module1"
Sub Format_Headers()
Attribute Format_Headers.VB_Description = "format headers and change values to currency"
Attribute Format_Headers.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' Format_Headers Macro
' format headers and change values to currency
'
' Keyboard Shortcut: Ctrl+j
'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Expense"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A3:F3").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Range("A3:F17").Select
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
    Range("C4:F17").Select
    Selection.NumberFormat = "$#,##0.00"
End Sub
