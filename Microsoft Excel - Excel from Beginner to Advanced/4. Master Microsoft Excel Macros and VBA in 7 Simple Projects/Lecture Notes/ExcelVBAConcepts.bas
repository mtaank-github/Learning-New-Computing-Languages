Attribute VB_Name = "Module1"

Public Sub FunWithVBAProcedures()

    ' This is a VBA Comment
    
    ' This line changes the value of Cell A1
    ActiveSheet.Range("A1").Value = "Hello World"
    
    ' Lets make a call to the Message box
    MsgBox (ActiveSheet.Range("A1").Value)
    
        
End Sub

Public Sub FunWithVariables()

    ' Here, we will use Variables to write out some stuff
    
    Dim userName As String ' This is a string (text) variable
    Dim userAge As Integer 'This is a numeric (integer) variable
    
    ' Now we can assign Values to these Variables
    userName = "Mukesh"
    userAge = 24
    
    ' Here, we write out the messages we want the user to see
    MsgBox ("Hello " & userName & "! You are " & userAge & " years old.")
    MsgBox (userName & ", you were born in " & Year(Now()) - userAge & ".")

End Sub
