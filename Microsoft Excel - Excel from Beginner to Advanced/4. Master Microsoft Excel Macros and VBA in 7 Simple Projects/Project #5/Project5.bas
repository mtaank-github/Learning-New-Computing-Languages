Attribute VB_Name = "Module1"
Sub LoopQuarterlyReport()
    
    ' Variables
    Dim ws As Worksheet
    Dim firstTime As Boolean
    
    ' Set the firstTime variable to be True
    firstTime = True
            
    ' Create the loop over the worksheets
    For Each ws In Worksheets
        
        ' Select the worksheet
        Worksheets(ws.Name).Select
        
        ' Set up condition to not include the Q.Report sheet in loop
        If ws.Name <> "YEARLY REPORT" Then
        
            ' Call the other procedures
            InsertHeaders
            FormatHeaders
            AutomateTotalSUM
            
            ' Select the data on the current worksheet
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
            
            ' Copy the data on the current worksheet
            Selection.Copy
            
            ' Select the Q.Report sheet
            Worksheets("YEARLY REPORT").Select
                            
            ' Paste the data to the Q.Report sheet
            Range("A100000").Select
            Selection.End(xlUp).Select
            
            ' Set up logic so this doesnt run the first time
            If firstTime <> True Then
                ActiveCell.Offset(1, 0).Select
            Else
                firstTime = False
            End If
            
            ' Continue with the pasting
            ActiveSheet.Paste
        
        End If
                                
    ' Move the the next sheet
    Next ws
    
    ' We still want to add formatting and calc. sum on the Q.Report sheet
    Worksheets("YEARLY REPORT").Select
    InsertHeaders
    FormatHeaders
    AutomateTotalSUM

End Sub



Public Sub AutomateTotalSUM()
    Dim lastCell As String
    Dim ws As Worksheet

    'For Each ws In Worksheets
        'Worksheets(ws.Name).Select
        
        Range("F2").Select
        
        Selection.End(xlDown).Select
        
        lastCell = ActiveCell.Address(False, False)
        
        ActiveCell.Offset(1, 0).Select
        
        ActiveCell.Value = "=sum(F2:" & lastCell & ")"
    'Next ws
End Sub

Sub InsertHeaders()
'
' InsertHeaders Macro
' Inserts a new row and add the list headers
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
'
' FormatHeaders Macro
' Formats list headers and list content
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Size = 12
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "Currency"
    Columns("B:F").Select
    Columns("B:F").EntireColumn.AutoFit
    Range("A2").Select
End Sub

