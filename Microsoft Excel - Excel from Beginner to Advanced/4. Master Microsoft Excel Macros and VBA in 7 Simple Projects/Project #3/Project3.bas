Attribute VB_Name = "Module1"
Sub LoopInsertFormatHeaders()

    ' In this procedure, we will create the loops.
    
    ' Variable representing the next worksheet
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the worksheets collection
    For Each ws In Worksheets
    
        Worksheets(ws.Name).Select
        
        ' Condition to not run the macro if a worksheet has already been formatted.
        If Range("A1").Value <> "Division" Then
        
        ' Call the other Procedures
            InsertHeaders
            FormatHeaders
        End If
            
    Next ws

End Sub



Sub InsertHeaders()
Attribute InsertHeaders.VB_Description = "Insert a row and add list headers"
Attribute InsertHeaders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' InsertHeaders Macro
' Insert a row and add list headers
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Division"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A2").Select
End Sub
Sub FormatHeaders()
Attribute FormatHeaders.VB_Description = "Format the Headers and List Contents"
Attribute FormatHeaders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatHeaders Macro
' Format the Headers and List Contents
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
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
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = "$#,##0.00"
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("A2").Select
End Sub
