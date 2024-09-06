Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    Call SendEmailForTodayTasks
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred during Workbook_Open: " & Err.Description
End Sub
