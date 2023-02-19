Attribute VB_Name = "Module1"
Public Sub FunWithInputBox()
    ' Practice using the Input box. We will get input from User
    Dim userInput As String
    userInput = InputBox("What is your Favourite Colour?", "Favourite Colour")
End Sub

Public Sub UserSortInput()

    ' Get the sort order from the user
    
    ' Assign all variables here
    Dim sortOrder As Integer
    Dim promptMsg As String
    Dim tryAgain As Integer
    
    ' Error Handling
    On Error GoTo errorHandler
    
    ' Well provide a message to the user so they know how to sort
    promptMsg = "How would you like to sort the list?" & vbCrLf & _
    "1 - Sort by Division" & vbCrLf & _
    "2 - Sort by Category" & vbCrLf & _
    "3 - Sort by Total"
    
    ' Well assign a value to sortOrder variable from the user
    sortOrder = InputBox(promptMsg, "Sort Order")
    
    ' Well now use Logic to call a certain Sorting Macro below
    If sortOrder = 1 Then
        DivisionSort
    ElseIf sortOrder = 2 Then
        CategorySort
    ElseIf sortOrder = 3 Then
        TotalSort
    Else
errorHandler:
        tryAgain = MsgBox("Invalid Selection: Would you like to try again?", vbYesNo)
        ' Need a condition for if they selected "Yes"
        If tryAgain = 6 Then
            UserSortInput
        End If
    End If
End Sub

Public Sub DivisionSort()
    ' Sorts the List by the Division
    Columns("A:F").Sort key1:=Range("A2"), order1:=xlDescending, Header:=xlYes
End Sub

Public Sub CategorySort()
    ' Sorts the List by the Category
    Columns("A:F").Sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes
End Sub

Public Sub TotalSort()
    ' Sorts the List by the Total
    Columns("A:F").Sort key1:=Range("F2"), order1:=xlDescending, Header:=xlYes
End Sub

