Attribute VB_Name = "Module1"

Public Sub ImportTextFile()
    
    ' Variables
    Dim textFile As Workbook
    Dim openFiles() As Variant
            
    ' Get Files using the Application GetOpenFiles
    'openFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
    ' Get files using the function
    openFiles = GetFiles
    
    ' Turn off screen flicker
    Application.ScreenUpdating = False
    
    ' Set up a Loop to open multiple textfiles
    For i = 1 To Application.CountA(openFiles)
        
        ' Assign textfile to the variable and open text file
        Set textFile = Workbooks.Open(openFiles(i))
        
        ' Copy the contents of this text file and place it in this workbook
        textFile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        
        ' Add a new sheet to paste into
        Workbooks(1).Worksheets.Add
        
        ' Paste the contents
        ActiveSheet.Paste
        
        ' Rename the new worksheet
        ActiveSheet.Name = textFile.Name
        
        ' Clear the clipboard after we paste
        Application.CutCopyMode = False
                        
        ' Remember to close the text file
        textFile.Close
       
    Next i
    
    ' Turn it back on
    Application.ScreenUpdating = True
    
End Sub

Public Function GetFiles() As Variant

    ' Procedure to use the application and prompt the user to open file(s).
    GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
        
End Function
