Attribute VB_Name = "Module1"

Public Sub AutomateTotalSum()

    ' Create a procedure to calculate the sum of the total expenses
    ' and place it in the cell directly below the Total Expenses column.
    
    ' Variables
    Dim lastCell As String
    Dim ws As Worksheet
    
    ' Set up the Loop to loop over worksheets
    For Each ws In Worksheets
    
        Worksheets(ws.Name).Select
    
        ' Select the column that the expenses are in
        Range("F2").Select
        
        ' Select the cell at the end of the column
        Selection.End(xlDown).Select
        
        lastCell = ActiveCell.Address(False, False)
        
        ' Select the cell below the last cell in the column with data
        ActiveCell.Offset(1, 0).Select
        
        ' Perform the Sum
        ActiveCell.Value = "=sum(F2:" & lastCell & ")"
        
    Next ws

End Sub
