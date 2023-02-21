VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Welcome to the Report Form"
   ClientHeight    =   2800
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4288
   OleObjectBlob   =   "Project6_frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddWorksheet_Click()
    ' Set up variable to catch errors
    Dim tryAgain As Integer
    ' Proceed to error handler if they make a mistake
    On Error GoTo errHandler
    ' Add a new worksheet and name it:
    Worksheets.Add before:=Worksheets(1)
    ActiveSheet.Name = InputBox("Please Name the new Worksheet:")
    ' Handle the error
errHandler:
    tryAgain = MsgBox("Invalid Worksheet Name. Would you like to Try Again?", vbYesNo)
    ' If they said yes
    If tryAgain = 6 Then
        btnAddWorksheet_Click
    Else
        Application.DisplayAlerts = False
        ActiveSheet.Delete
    End If
End Sub

Private Sub btnRunReport_Click()
    LoopYearlyReport
End Sub

Private Sub cboWhichSheet_Change()
    ' Take the value the user chose and select that worksheet
    Worksheets(Me.cboWhichSheet.Value).Select
End Sub

Private Sub UserForm_Click()
    MsgBox ("Hello!")
End Sub

Private Sub UserForm_Initialize()
    ' Define the index variable
    Dim i As Integer
    ' initialize the index variable
    i = 1
    ' Set up Do-While Loop
    Do While i <= Worksheets.Count
        Me.cboWhichSheet.AddItem Worksheets(i).Name
        i = i + 1
    Loop
End Sub


