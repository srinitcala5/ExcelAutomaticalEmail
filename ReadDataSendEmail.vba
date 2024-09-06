' In a standard module (e.g., Module1)
Option Explicit

Sub SendEmailForTodayTasks()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Remainders") 'Change to your sheet name if different
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim outlookApp As Object
    Dim emailItem As Object
    Set outlookApp = CreateObject("Outlook.Application")
    
    Dim i As Long
    Dim taskDate As Date
    Dim task As String
    Dim sendTo As String
    
    For i = 2 To lastRow 'Assumes headers in row 1
        taskDate = ws.Cells(i, 1).Value
        task = ws.Cells(i, 2).Value
        sendTo = ws.Cells(i, 3).Value
        
        If DateValue(taskDate) = DateValue(Date) Then
            Set emailItem = outlookApp.CreateItem(0)
            With emailItem
                .To = sendTo
                .Subject = "Task for Today: " & task
                .Body = "This is a reminder for the following task: " & task
                .Send
            End With
        End If
    Next i
    
CleanUp:
    Set outlookApp = Nothing
    Set emailItem = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    Resume CleanUp
End Sub
