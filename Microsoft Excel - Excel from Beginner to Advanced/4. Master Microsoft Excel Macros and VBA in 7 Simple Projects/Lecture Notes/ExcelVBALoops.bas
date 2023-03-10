Attribute VB_Name = "Module1"

Public Sub FunWithLogic()

    ' IF Logic to Determine the Age
    
    If ActiveCell.Value < 21 Then
        MsgBox ("User is Under 21")
    ElseIf ActiveCell.Value >= 21 Then
        MsgBox ("User is 21 or over")
    Else
        MsgBox ("Error")
    End If

End Sub

Public Sub FunWithSelect()

    ' Now using the Select Statement to set up multiple cases
    
    Select Case ActiveCell.Value
        
        ' Case #1
        Case Is > 100
            MsgBox ("User is OVER 100?!?!?!?!?!?!")
        ' Case #2
        Case 21 To 99
            MsgBox ("User is between 21 and 99 Years old")
        ' Case #3
        Case Else
            MsgBox ("User is under 21")
    End Select

End Sub

Public Sub FunWithDoWhileLoops()

    Dim i As Integer
    i = 1
    
    ' Set up the Do-While Loop
    Do While ActiveCell.Value <> "" 'i <= 10
    
        ' Call the last Procedure
        FunWithSelect
        
        'Need to update the Active Cell
        ActiveCell.Offset(1, 0).Select
        
        ' Update the index for the loop
        i = i + 1
    Loop

End Sub

Public Sub FunWithForEachLoops()

    Dim user As Range
    
    For Each user In Selection
    
        FunWithSelect
        
        ActiveCell.Offset(1, 0).Select
    
    Next user
        
End Sub

Public Sub FunWithForNextLoop()

    Dim i As Integer
    
    For i = 1 To ActiveSheet.UsedRange.Rows.Count - 1
    
        FunWithSelect
        
        ActiveCell.Offset(1, 0).Select
        
    Next i

End Sub
